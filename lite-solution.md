# PPT 资产提取 · 轻量实施方案 Lite v1.8

> 定位：**先跑起来，边用边调**。单人/小团队、可信内网、试点规模（几百份 PPT / 万页以内）。
> 与 `tech-solution.md`（完整方案）的关系：**Lite 是唯一的第一期实施目标，完整方案是规模化蓝图**。
> v1.8 = v1.7 + 复审收尾：媒体失败补偿确定规则（防误删共享文件）、备份改停服复制/Backup API、
> fslock 范围收窄（仅共享规范路径）、GC 注释修正、mids_dead 行展开、预检魔数分流补全。
> 五条铁律（一切设计服从）：
> ① 人工成果（review/design_system）与解析产物分表，重跑永不丢失；
> ② 数据库是唯一真相（队列、状态、生效版本），内存信号仅用于唤醒，进程重启零丢失；
> ③ 所有读取面只经统一入口取"生效数据"（`v_active_slide`），禁止自行拼状态条件；
> ④ **文件先发布、DB 后切换**：目录级原子发布，DB 提交时目标文件必然已完整存在；
> ⑤ **文件生命周期三原则**：extraction ID 永不复用（autoincrement）；**唯一锁序** = 外层取 `fslock` → 短 DB 事务 → commit → **仍在锁内**把待删路径搬入 `trash/{uuid}/` → 解锁（helper 一律不再自行取锁）；延迟删除必先进 trash（uuid 路径与规范路径不可能碰撞）。
>
> 状态迁移一律带来源条件并断言 `rowcount=1`；GC 只在启动停流窗口运行（无独立 CLI），永不与处理并发。

---

## 1. 技术栈（6 个依赖 + LibreOffice）

| 用途 | 选型 | 说明 |
| --- | --- | --- |
| 解析 | python-pptx + lxml | 文本/形状/表格/图表/坐标/样式；主题色读 `theme1.xml` |
| OLE/宏检测 | olefile | 旧 `.ppt` 检测 `_VBA_PROJECT_CUR`；加密容器检测 `EncryptionInfo` |
| 渲染 | LibreOffice + **pypdfium2**（Apache-2.0） | LO 整文件转 PDF → pypdfium2 每页 PNG + 缩略图 |
| 图像 | Pillow + imagehash | sha256 精确去重 + phash 近似浏览 |
| 存储 | SQLite（WAL）+ FTS5 **trigram**（普通内容表） | 中文子串检索；短查询/异常回退 LIKE |
| 服务 | FastAPI + Jinja2 | 服务端渲染，无前端构建链；**单 uvicorn 进程** |

版本钉死（pyproject 锁定）：python-pptx / pypdfium2 / LO / Pillow —— 同时写进 `extraction` 表，可追溯数据由哪个版本产出。

## 2. 目录结构

```text
ppt-lite/
├── main.py            # FastAPI：上传/浏览/搜索/审核/删除（启动序列见 §4.10）
├── pipeline.py        # 解析+渲染 → ExtractedDocument（§7）
├── render.py          # LO 调用 + PDF→PNG（每文件独立 profile / 超时 / kill 进程树；锁外）
├── db.py              # schema / 领取 / finalization / mark_failed / 删除 / GC（§4 全部入口）
├── design_draft.py    # spec → design_system 草稿（幂等；锁外调用）
├── search.py          # FTS / LIKE 回退
├── fslock.py          # 进程级文件锁（threading.Lock）+ trash 搬运 + after-commit 协议
├── templates/         # 4 个页面
└── data/
    ├── app.db
    ├── raw/{sha256}.{ext}          # 原件永不改动；.ppt 转换件 raw/derived/{sha256}.pptx
    ├── media/{sha256}.{ext}        # 提取媒体（原格式；sha256 即规范名）
    ├── tmp/                        # 一切锁外工作的私有目录（uuid 子目录）
    │   ├── upload-{uuid}/          #   上传暂存（哈希/预检/LO 转换）
    │   └── {extraction_id}/        #   解析工作目录：pdf/、preview/{slides,thumbs}/、media_staging/
    ├── previews/
    │   ├── .staging/{ex_id}/       # 发布中转（校验完整后原子 rename）
    │   └── {ex_id}/slides/{no}.png、{ex_id}/thumbs/{no}.png   # 不可变版本目录
    └── trash/{uuid}/               # 延迟删除中转（uuid 隔离，杜绝 TOCTOU）
```

## 3. 数据模型（8 张表 + 1 视图 + 1 FTS）

```sql
create table file (
  id integer primary key,
  sha256 text unique not null,
  name text not null,                  -- 最近一次上传原名（展示用）
  category text not null check (category in ('material','spec','sample')),
  status text not null default 'idle'
              check (status in ('idle','queued','working','done','failed')),
                                      -- 仅展示"最近状态"；真相在 extraction（§4.6 状态表）
  error text,
  raw_path text not null,
  derived_path text,
  note text,
  created_at text default (datetime('now'))
);

create table extraction (
  id integer primary key autoincrement,   -- ★ 永不复用（previews 目录键与 design_system 幂等键依赖此性质；
  file_id integer not null references file(id) on delete cascade,   -- 禁止重置 sqlite_sequence）
  parser_version text, renderer_version text,
  status text not null default 'queued'
              check (status in ('queued','running','done','failed')),
  started_at text, finished_at text, error text
);
create unique index uq_ex_one_active on extraction (file_id)
  where status in ('queued','running');           -- active 每文件至多一个（DB 强制）
create unique index uq_ex_one_done on extraction (file_id) where status='done';

create table slide (
  id integer primary key,
  extraction_id integer not null references extraction(id) on delete cascade,
  file_id integer not null references file(id) on delete cascade,
  no integer not null,
  title text,
  search_text text,
  texts json, struct json,             -- struct 含 warnings/hidden/notes
  png text not null, thumb text not null,   -- previews/{ex_id}/... 初次写入即定
  unique (extraction_id, no)
);
-- slide.file_id 由 writer 从 extraction 行带出（不信任调用方传两份）

create table media (
  id integer primary key,
  sha256 text unique not null,
  phash text, fmt text, w integer, h integer,
  path text, preview text
);

create table media_ref (
  slide_id integer not null references slide(id) on delete cascade,
  media_id integer not null references media(id),
  role text not null default 'content',
  primary key (slide_id, media_id, role)
);

-- ===== 人工成果（重跑不碰） =====
create table slide_review (
  file_id integer not null references file(id) on delete cascade,
  no integer not null,
  status text not null default 'pending'
       check (status in ('pending','approved','reference','forbidden')),
  slide_type text, quality integer,
  tags text,                          -- 规范化 ',a,b,'（§5）
  is_template integer default 0,
  template_json json,
  note text,
  updated_at text default (datetime('now')),
  primary key (file_id, no)
);
create table media_review (
  sha256 text primary key,            -- 内容级结论；文件删除时保留（同媒体再现自动带出）
  status text default 'pending'
       check (status in ('pending','approved','reference','forbidden')),
  tags text, note text,
  updated_at text default (datetime('now'))
);
create table design_system (
  id integer primary key,
  version integer not null unique,
  json text not null,
  status text default 'draft' check (status in ('draft','confirmed')),
  confirmed_by text, confirmed_at text,
  source_file_id integer references file(id) on delete set null,
  source_extraction_id integer,        -- 幂等键（唯一）；故意不设 FK：源文件删除后草稿/确认版保留；
  schema_version integer default 1,   -- 依赖 extraction.id autoincrement 永不复用
  unique (source_extraction_id)
);

-- FTS：普通内容表；slide 删除触发器同步清理
create virtual table slide_fts using fts5(title, body, tokenize='trigram');
create trigger slide_fts_del after delete on slide begin
  delete from slide_fts where rowid = old.id;
end;

-- 生效数据统一入口：页面墙/详情/搜索/媒体反查/模板跳转一律用它
create view v_active_slide as
select s.* from slide s
join extraction e on e.id = s.extraction_id and e.status='done';
```

## 4. 流水线与状态规则

### 4.0 fslock 与 after-commit 协议（铁律⑤的实现约定）

```python
# fslock.py 提供的唯一模式——复合操作模板（所有**共享规范路径**的发布/替换/trash 搬运都必须套用；
# 私有 tmp/{uuid}/ 内的锁外工作不受此限）：
with fslock:                                  # ① 最外层取锁，全程只取这一次
    trash_paths = []
    with db_transaction(busy="short") as tx:  # ② 短 DB 事务（busy_timeout 压短，快速失败）
        ...业务 SQL...
        trash_paths = collect(...)            #    事务函数返回"待搬运清单"，不用隐式回调
    # ③ 事务已 commit 且仍持锁：同步搬入 trash（失败仅记日志，留待启动 GC）
    for p in trash_paths: to_trash(p)
# ④ 解锁
# 规则：helper（gc_media/mark_failed 内部等）一律提供 *_locked 变体，绝不在内部再取锁；
#       BUSY 重试=重试整个复合操作（重试前丢弃 trash_paths）。
```

### 4.1 上传与预检（重活全在锁外私有目录；发布与登记在锁内一气呵成）

```text
POST /upload（或 watch 目录扫描）
── 锁外（tmp/upload-{uuid}/）──
  1. 扩展名白名单 .pptx/.ppt；大小上限 500MB；落暂存并算 sha256
  2. 预检（全部针对暂存文件；按魔数三级分流，错误原因才准确）：
     读文件头魔数：
     PK\x03\x04（ZIP）→ .pptx 路线：zipfile 试开 → 查 ppt/vbaProject.bin（有→拒收'含宏'）
       （.pptx 扩展名但 ZIP 无法打开 → 拒收'文件损坏'）
     D0 CF 11 E0（OLE CFB）→ .ppt 路线：olefile 查 _VBA_PROJECT_CUR（有→拒收'含宏'）
       ；查 EncryptionInfo 流（有→拒收'加密'——含 .pptx 扩展名但实为 CFB 的新格式加密）
       ；否则视为旧 .ppt → LO 转 .pptx → tmp/upload-{uuid}/derived.pptx（★LO 在锁外）
         → zipfile+python-pptx 试开校验；失败→拒收'转换失败，请手工另存为 .pptx'
     其他魔数 → 拒收'无法识别的格式'
── 锁内（fslock + 单 DB 事务 + 文件发布）──
  3. 查 file by sha256：
     已存在 → 复用（不重复发布文件）；category 不同→提示"已按{旧类}入库，可修改"
     不存在 → raw 暂存 rename→raw/{sha256}.{ext}（已存在则删暂存用现成，内容寻址幂等）；
              derived 同理 → raw/derived/{sha256}.pptx
     同事务：insert file(+raw_path/derived_path)；insert extraction(queued)；file.status='queued'
  4. 预检失败路径：**不创建 file 行**（避免 raw_path NOT NULL 与无原件记录的矛盾）——
     返回结构化上传错误（HTTP 4xx + 原因：含宏/加密/转换失败/格式不支持），
     暂存目录锁内 → trash；界面"失败历史"不记录预检拒绝（它从未入库）
     （如需留档被拒文件：另行人工放入 raw/quarantine/，属运维操作，不在系统内）
  5. 提交后解锁 → 唤醒后台线程
  ※ raw 发布后 DB 事务重试耗尽：锁内复查 DB 无此 sha256 行 → 已发布的 raw/derived → trash
    （另由启动 GC 步骤 d 对账兜底，防空间泄漏）
```

### 4.2 领取协议（DB 为真相；单事务；无需文件锁；含防饿死）

```sql
begin;
update extraction
set status='running', started_at=datetime('now'), error=null
where id = (select id from extraction
            where status='queued'
              and file_id not in (select file_id from extraction where status='running')
            order by id limit 1)
returning id, file_id;        -- 0 行=队空，回滚并 sleep 等唤醒
update file set status='working', error=null where id=:上一步的 file_id;
commit;                        -- 整个事务参与 SQLITE_BUSY 重试
-- 防饿死守卫：file_id not in (running) 跳过"该文件仍有任务在跑"的行；
-- uq_ex_one_active 保证这种情况本不该存在——守卫仅防御"重跑入队与领取竞态"的极端时序
```

### 4.3 处理（后台单线程；重活在锁外）

```text
tmp/{ex_id}/ 工作目录布局：
  pdf/                          LO 产物（LO 在锁外跑）
  preview/slides/{no}.png、preview/thumbs/{no}.png     ← 仅此子树参与发布
  media_staging/                媒体暂存

pipeline.py 产出 ExtractedDocument（§7；slide.png/thumb 命名为 previews/{ex_id}/...，
  文件此刻在 tmp/{ex_id}/preview/ 下）
逐页写 DB（短事务，无文件操作，无需锁）：
  slide 行 + slide_fts 行；slide.file_id 取自 extraction 行
媒体落盘（铁律⑤模板，顺序=文件先发布、DB 后提交，对齐铁律④）：
  规范路径 = media/{sha256}.{fmt}
  fmt 解析规则（确定的，无歧义）：
    已有 media 行 → 沿用该行持久化的 path（权威来源；同内容不同扩展名不再生成第二份）
    新媒体       → 首次发现时的实际格式 fmt（来自 zip entry 扩展名，小写归一）
  with fslock:
      if not exists(规范路径):                 # need_write 精确语义：仅看规范路径是否存在
          rename media_staging/x → 规范路径    # ① 文件先发布（内容寻址，幂等）
      with db_transaction(busy="short") as tx: # ② DB 后提交
          media_id = insert … on conflict(sha256) do update
                     set phash=coalesce(media.phash, excluded.phash) returning id
          insert media_ref(...)
  # DB 失败补偿（确定的规则，防误删共享文件）：
  #   重试耗尽后锁内复查——仅当 DB 中不存在该 sha256 的 media 行时，规范文件才是孤儿 → trash；
  #   media 行存在（旧 done 引用中的共享媒体）→ 绝不搬运规范文件；
  #   同批次前页已发布、本次无暂存文件 → 同样不搬运（该文件属于已成功的发布）
  #   兜底：启动 GC 步骤 c 对账（无 media 行的文件）
  # 无需写盘时：暂存文件随 tmp/{ex_id} 清理
  # 特例：media_staging 中无该 sha 的暂存文件（同批次前页已发布过）→ 跳过写盘即可
完成后：§4.4 发布 + finalization；任何异常 → §4.5 失败清理
```

### 4.4 发布与 Finalization（文件先发布、DB 后切换；§4.0 模板）

```python
def finalize_extraction(ex_id):            # 幂等：重复调用安全
    with fslock:
        ex = db.one("select file_id, status from extraction where id=?", ex_id)
        if ex is None: to_trash(tmp/{ex_id}); return          # 已删除：清残留
        if ex.status == 'done':
            if not dir_complete(previews/{ex_id}):
                raise InvariantViolation(f"done 目录缺失/不完整: {ex_id}")   # 不改状态、不重发；
            return                                            # 记高优先级日志，停机人工处理或重跑生成新 ex
        if ex.status != 'running': raise InvalidState

        trash = []
        # A. 文件发布（rename 前任何失败 → raise，由调用方走 §4.5；旧 done 未动）
        staging = previews/.staging/{ex_id}
        to_trash(staging)                        # staging 是私有中转，可安全重来
        copy tmp/{ex_id}/preview/* → staging/
        verify_complete(staging, slide_count)    # 页数=slide 数、文件齐全
        if exists(previews/{ex_id}):             # 上次发布后、DB 切换前崩溃的孤儿
            to_trash(previews/{ex_id})           # 进 trash，不盲删
        os.rename(staging, previews/{ex_id})     # 同卷原子

        # B. DB finalization（单事务；失败 raise → §4.5，previews/{ex_id} 成为 GC 孤儿）
        with db_transaction(busy="short") as tx:
            mids_old, old_ex_id = [], None                        # ★ 先初始化（首跑无旧 done）
            row = tx.one("select id from extraction where file_id=? and status='done' and id<>?",
                         ex.file_id, ex_id)
            if row:
                old_ex_id = row.id
                mids_old = collect_media_ids(tx, extraction=old_ex_id)   # 取证（级联删除前）
                tx.exec("delete from extraction where id=?", old_ex_id)  # 级联 slide(触发器清FTS)/media_ref
            cur = tx.exec("update extraction set status='done', finished_at=datetime('now')
                          where id=? and status='running'", ex_id)
            assert cur.rowcount == 1             # 条件迁移（迟到回调在此止步）
            tx.exec("update file set status='done', error=null where id=?", ex.file_id)
            mids_dead = gc_media_locked(tx, mids_old)             # 返回待删文件清单（仅 dead，去重）
        # C. 提交成功、仍持锁：搬运（失败仅记日志，留给启动 GC；绝不动已 done 状态）
        to_trash(tmp/{ex_id})
        if old_ex_id is not None: to_trash(previews/{old_ex_id})
        for row in mids_dead:                          # rows of (path, preview)
            to_trash(row.path); to_trash_if_exists(row.preview)
    # D. 锁外：规范草稿（独立失败面，绝不影响 done）
    design_draft.maybe_generate(ex_id)           # §4.9
```

### 4.5 失败清理 mark_failed（旧 done 永不动；幂等；两个入口）

```python
def mark_failed(ex_id, reason):                # 入口 1：正常失败路径（自取锁）
    with fslock:
        _mark_failed_locked(ex_id, reason)

def _mark_failed_locked(ex_id, reason):        # 入口 2：启动恢复等外层已持锁者调用（★不再取锁）
        row = db.one("select file_id from extraction where id=? and status='running'", ex_id)
        if not row: to_trash_if_exists(tmp/{ex_id}); return   # 已终态：幂等退出，顺手清残留
        with db_transaction(busy="short") as tx:
            mids = collect_media_ids(tx, extraction=ex_id)
            tx.exec("delete from slide where extraction_id=?", ex_id)   # 触发器清FTS；级联 media_ref
            cur = tx.exec("update extraction set status='failed', error=?, finished_at=datetime('now')
                          where id=? and status='running'", reason, ex_id)
            assert cur.rowcount == 1
            tx.exec("update file set status='failed', error=? where id=?", reason, row.file_id)
            mids_dead = gc_media_locked(tx, mids)
        # 提交成功、仍持锁（由外层保证）：
        to_trash(tmp/{ex_id}); to_trash_if_exists(previews/{ex_id})    # 发布失败的孤儿一并进 trash
        for row in mids_dead:                                          # rows of (path, preview)
            to_trash(row.path); to_trash_if_exists(row.preview)
```

**异常分流（调用方约定）**：`SQLITE_BUSY` → 回滚、退出锁作用域、**重试整个复合操作**（最多 3 次）；仅重试耗尽或非重试错误才调 `mark_failed`。禁止在持有 finalize 锁时调用公开 `mark_failed()`。

### 4.6 状态与转换总表（file.status 展示值 / extraction 真相）

| 场景 | extraction | file.status | 读取面行为 |
| --- | --- | --- | --- |
| 上传入队 | queued | queued | "排队中"；有旧 done 同时展示旧数据 |
| 领取 | queued→running | working | 同上 |
| 首跑成功 | running→done（无旧） | done | v_active_slide 切新数据 |
| 重跑成功 | 旧 done 级联删除 + running→done（原子） | done | 同上 |
| 首跑失败 | running→failed（slide 已清） | failed | 无数据，显示错误 |
| 重跑失败 | running→failed（旧 done 保留生效） | failed | **展示旧数据**+横幅"重跑失败：{error}，以下为上次成功数据" |
| 发布/finalization 失败 | running→failed | failed | 同上两条（旧 done 未动） |
| 删除（终态） | 级联全删 | — | §4.8 |
| 启动恢复 | running→failed | failed | 同"重跑失败"；error='进程中断' |

重跑/入队接口：显式检查无 active extraction；`uq_ex_one_active` 冲突 → 409"已有任务在排队或处理中"。

### 4.7 gc_media_locked（helper；调用方必须已持 fslock；返回待删清单）

```python
def gc_media_locked(tx, mids):
    dead = dedup([m for m in mids
                  if not tx.one("select 1 from media_ref where media_id=? limit 1", m)])
    if not dead: return []
    paths = tx.all("select path, preview from media where id in ({dead})")
    tx.exec("delete from media where id in ({dead})")
    return paths            # ★ 只返回清单；搬运由外层在 commit 后、持锁状态下执行
```

### 4.8 删除文件（仅终态：无 active extraction；否则 409；§4.0 模板）

```python
with fslock:
    with db_transaction(busy="short") as tx:
        info = tx.one("select raw_path, derived_path from file where id=?", fid)   # None→404
        ex_ids = tx.all("select id from extraction where file_id=?", fid)
        mids   = collect_media_ids(tx, file=fid)
        cur = tx.exec("""delete from file where id=?
                         and not exists (select 1 from extraction e where e.file_id=file.id
                                         and e.status in ('queued','running'))""", fid)
        if cur.rowcount == 0: raise Conflict(409)   # 守卫在 DELETE 内：与并发重传/重跑竞态安全
        mids_dead = gc_media_locked(tx, mids)
        # media_review 按 sha256 保留；design_system.source_file_id 置空（FK on delete set null）
    # 提交成功、仍持锁：
    to_trash(info.raw_path); to_trash_if_exists(info.derived_path)
    for ex in ex_ids: to_trash(previews/{ex}, tmp/{ex})
    for row in mids_dead:                                          # rows of (path, preview)
        to_trash(row.path); to_trash_if_exists(row.preview)
```

### 4.9 规范草稿（design_draft.py；锁外；幂等、不覆盖人工）

```text
时机：finalize 步骤 D（锁外）；仅当 file.category='spec'
幂等：design_system.source_extraction_id 唯一 → 同一解析结果至多一份自动草稿
规则：无该 source_extraction_id 行 → insert draft（version=max(version)+1）
     人工编辑/confirmed 的行永不被自动流程触碰
失败：记日志（详情页提示"草稿生成失败，可重跑"），不影响已 done 的 extraction
```

### 4.10 启动序列与 GC（停流窗口内完成，永不与处理并发）

```text
app 启动序列（顺序固定；整个恢复+GC 在同一持锁窗口内，内部全部用 *_locked 变体）：
 1. 取 fslock
 2. 启动恢复：枚举 status='running' 的 extraction
    → 逐个调用 _mark_failed_locked(ex,'进程中断')      # ★锁内变体，不重复取锁（消自锁）
    （复用 §4.5 完整清理流程；queued 不动——持久队列，worker 启动后自然消费）
 3. GC（此刻 worker/HTTP 均未启动，天然互斥）：
    a. previews/{ex_id}：DB 无此 extraction 或 status='failed' → trash
       （done 保留；queued 仍存在但无 previews 目录——预览只在 finalize 时发布，无需判定）
    b. previews/.staging/*、tmp/* → trash
    c. media/ 无对应 media 行的文件 → trash
    d. raw/ 与 raw/derived/ 无对应 file 行（按 sha256 匹配文件名）的文件 → trash（对账）
    e. trash/* → 物理删除（uuid 路径与规范路径不可能碰撞）
 4. 解锁 → 启动后台 worker 线程 → 启动 HTTP

维护 GC：**不提供独立 CLI**（进程内 threading.Lock 无法阻止应用与 CLI 并发，
跨进程锁超出 Lite 范围）。GC 仅存在于上述启动停流窗口；需要深度清理时停服重启即可。
```

## 5. 检索（FTS + LIKE 回退；两条完整 SQL）

```sql
-- FTS 路（q 为连续 ≥3 中文字符：re.fullmatch(r'[\u4e00-\u9fff]{3,}', q.strip())）
-- 绑定 :phrase = '"' + q.replace('"','""') + '"'
select s.id, s.file_id, s.no, s.title, s.thumb, sr.status as review_status
from slide_fts f
join v_active_slide s on s.id = f.rowid
left join slide_review sr on sr.file_id = s.file_id and sr.no = s.no
where slide_fts match :phrase
  and coalesce(sr.status, 'pending') <> 'forbidden'
  -- [and sr.slide_type = :type] [and sr.tags like '%,'||:tag||',%']
order by rank limit 100;

-- LIKE 路（其余一切查询：英文/短中文/混排；转义 \\ % _ 后 '%…%' ESCAPE '\'）
select s.id, s.file_id, s.no, s.title, s.thumb, sr.status as review_status
from v_active_slide s
left join slide_review sr on sr.file_id = s.file_id and sr.no = s.no
where s.search_text like :pattern escape '\'
  and coalesce(sr.status, 'pending') <> 'forbidden'
  -- [and sr.slide_type = :type] [and sr.tags like '%,'||:tag||',%']
limit 100;
```

标签规范化：`tags` 存 `,a,b,`；写入时拆分→trim→丢空→去重→重组；标签值禁含逗号。

## 6. 界面（4 个页面）

1. **文件列表**：状态徽标/错误/重跑（终态才可用）/删除（终态才可用）/编辑 category、note；failed-but-old-data 徽标。
2. **页面墙 + 详情**：数据一律 v_active_slide；筛选（文件/类型/标签/审核状态）→ 详情（大图 + 文本 + 对象简表 + 媒体引用 + warnings + 重跑失败横幅如有）+ 审核表单（upsert slide_review）。
3. **素材网格**：phash 聚簇；同 sha256 聚合；"被引用于"走 v_active_slide 反查；打标写 media_review。
4. **模板与规范**：模板候选（slide_review.is_template=1 + template_json；缩略图+容量表单+来源页跳转）；规范草稿（自动 draft → 人工编辑 → 确认 → 导出 JSON）。

## 7. ExtractedDocument 契约（迁移桥 + 内部接口）

**解析器不直接写库**：`pipeline.py` 输入文件路径与 tmp 目录，产出可序列化 dict：

```text
ExtractedDocument v1:
  source{name, sha256, category, raw_path, derived_path?}
  versions{parser, renderer}
  slides[{no, title, texts, struct, notes, hidden, warnings,
          png, thumb}]                  # 最终路径 previews/{ex_id}/...（发布后生效）
  media[{sha256, fmt, w, h, path, preview, phash}]
  occurrences[{slide_no, media_sha256, role}]
  design{theme_colors, theme_fonts, page_size, common_rules}   # spec 类由 design_draft.py 消费
  #   theme_colors/theme_fonts：主题声明；font_usage/fill_usage：run 级字体与形状填充实采（规范草稿数据源）
  render_error?                         # 文档级渲染失败
```

`db.py` 的 SQLite writer 消费它（逐页短事务）；未来 PG writer 消费同一契约 → 解析核心复用。

**SQLite → PostgreSQL 迁移是领域重构而非导表**：冻结写入 → 哈希校验 → ID 映射 → file/extraction 拆完整方案表 → review 表按键合并 → 重建 FTS → 对账。人工成果（review×2 + design_system）可无损保留，是迁移的主要价值。

## 8. 明确砍掉的内容（及补回触发）

| 完整方案组件 | Lite 做法 | 补回触发 |
| --- | --- | --- |
| run/stage/artifact 状态机 + 租约 fencing | extraction 表兼持久队列 + 部分唯一索引协议 | 第二个写进程/第二台主机 |
| PostgreSQL + pgvector | SQLite + FTS5 trigram | >3 万页或检索 P95 不可接受 |
| RBAC/ACL/审计 | 无登录 + 自由备注 | 多部门或密级强制 |
| 媒体双层（内容/来源业务） | media + media_review | 跨部门共享素材且权限不同 |
| Saga 发布/对账 | 文件先发布 DB 后切换 + 停流 GC + trash | 数据成为正式资产底座 |
| 黄金样本 + SSIM 基线 | 肉眼抽查 + "预览仅供参考" | 预览质量被投诉 |
| OCR / 视觉向量 | 关（图片页搜不到为已知降级） | 成为主诉痛点 |
| Docker 安全加固 | 裸机进程 | 上生产/安全检查 |
| 受控标签词表 | 规范化文本标签 | 同义词污染搜索 |

## 9. 里程碑（单人）

| 时间 | 交付 |
| --- | --- |
| 第 1–3 天 | 预检/宏检测 + 解析 + ExtractedDocument + tmp 渲染 + 媒体落盘 |
| 第 1 周末 | 领取/finalization/mark_failed/启动恢复 + Web 浏览 + 中文搜索 + 文件管理 |
| 第 2 周末 | 审核打标 + 模板候选 + 素材网格 + 规范草稿自动生成与确认 |
| 第 3 周 | 试点真实批次（1 素材 + 1 规范 + 5–10 范例），按体验调整 |

**验收即故障注入矩阵**（第 1 周末自测，全部必须符合预期）：

| 故障点 | 预期 |
| --- | --- |
| 领取事务提交前崩溃 | 任务仍 queued，重启后继续 |
| running 中崩溃 | 启动后 failed（走 mark_failed 完整清理）；旧 done 保留并展示 |
| 预览发布（rename）前崩溃 | 新 extraction 不 done；旧数据不受影响；重启 GC 清 staging/tmp |
| 发布后、DB 切换前崩溃 | previews/{ex} 为孤儿；重启 GC（恢复转 failed 后）回收 |
| 媒体文件发布后、DB 提交前崩溃 | 文件为内容寻址孤儿；启动 GC 步骤 c 对账回收，无悬挂引用 |
| 预检失败（宏/加密/转换失败） | 不建 file/extraction；HTTP 4xx 结构化错误；暂存进 trash |
| BUSY 瞬时冲突（finalize 中） | 重试整个复合操作（≤3 次），不降级为 failed；耗尽才 mark_failed |
| DB 切换后、步骤 C 崩溃 | 新数据可读；tmp/旧目录/trash 由下次 GC 回收（done 状态不变） |
| 并发双击重跑 | 一个入队成功，一个 409 |
| 删除与重跑并发 | 只能一个成功；不删除活动任务 |
| 删除与同 sha256 上传并发 | 串行于 fslock：先到先得，后到者复用/重建规范文件，无误删（uuid trash 隔离） |
| 重复调用 finalize（同一 ex） | done 且目录完整→幂等成功；目录缺失→InvariantViolation（不改状态，人工处理） |
| 删除文件后重传同 sha256 | 新 file 正常；trash 中旧文件不碰撞新数据 |
| 删除最大 id extraction 后新建 | 新 id 不复用（autoincrement），目录/幂等键不碰撞 |
| GC 与处理并发 | 不可能：GC 仅存在于启动停流窗口（无独立 CLI，跨进程互斥超出 Lite） |
| mark_failed 迟到回调 | 幂等退出（非 running），仅清残留 |

## 10. 运行纪律（硬边界）

- SQLite 每连接：`PRAGMA journal_mode=WAL; PRAGMA foreign_keys=ON; PRAGMA busy_timeout=5000;`；fslock 内的短事务用更短的 busy 配置并快速失败重试**整个复合操作**。
- **单 writer 现实**：同一时刻仅一个写事务；HTTP 审核写与后台解析写串行竞争，全部短事务；统一封装 `SQLITE_BUSY` 重试（3 次随机退避），耗尽返回明确错误页。
- **文件锁（fslock）与锁序**：唯一锁序 = 外层 fslock → 短 DB 事务 → commit → 锁内搬 trash → 解锁；helper 用 `*_locked` 变体，绝不在内部再取锁；LO 转换/哈希/预检等重活一律锁外。
- 线程模型：每线程独立 sqlite 连接；HTTP 线程池 + 1 个后台解析线程，无跨线程共享连接。
- **单实例**：一个 uvicorn 进程、一个后台线程；不部署第二份（会破坏领取协议、单 writer 与进程内文件锁假设）。
- app.db 只放本地盘（**不放 NAS/同步盘**）；备份二选一：**停服后复制** .db（+ -wal/-shm，若存在）；
  或在线时用 **SQLite Backup API**（`sqlite3.Connection.backup`）。单次 checkpoint 后继续开放写入再复制
  **不是**受支持的备份方式（三文件非原子快照，可能得到不一致副本）。
- 上传 500MB / 解压总量 2GB / 单媒体 512MB 上限；zip entry 规范化后必须落在目标目录内（防穿越）。
- LO：soffice.exe 绝对路径、每文件独立 profile（`-env:UserInstallation=file:///...lo-<uuid>`）、超时 kill 进程树、任务后删 profile。
