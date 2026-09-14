# PPT 资产提取与检索系统 · 技术方案 v2.2

> 依据 `ppt-pipeline.md`（需求/思路）+ 三轮外部架构评审修订。
> **定位调整（第三轮评审采纳）：本方案为规模化蓝图；第一期实施目标是 `lite-solution.md`（加固版 Lite）。**
> 第一期范围：**解析、分类、预览、检索、人工审核**，不含 PPT 自动生成。
> 全部组件离线/内网运行，禁止任何外部 API 调用。

## v2.1 → v2.2 变更摘要（对照第三轮复审）

| # | 第三轮 P0 | v2.2 处理 |
| --- | --- | --- |
| 1 | fencing 未覆盖业务写入与产物发布 | §6.5 业务写入围栏：写事务内 `stage_run` 租约校验（for update）；产物 key 含 run_id 形成命名空间隔离；artifact 发布 CAS 同步校验租约 |
| 2 | 同配置重跑被 `(document_id,pipeline_version,config_hash)` 唯一约束阻断 | 改为部分唯一索引 `where status in ('running','succeeded')`；reprocess 幂等语义在 §6.6 重写（running 复用，终态新建） |
| 3 | `:param::type` 在 SQLAlchemy 下不可靠 | 全部改 `cast(:param as type)`；§9 增加绑定约定与"按是否提供向量组装固定查询" |
| 4 | 取消/失租/发布/supersede 状态转换不闭环 | §6.4 增加取消接管 SQL（换 token）；§6.8 状态机补 `running→cancelled`、GC 行；失败转 pending 清理租约字段 |
| 5 | `active_run_id` 可指向他文档的 run | 改复合外键 `(id, active_run_id) → pipeline_run(document_id, id)` |
| P1 | 领取未查父 run 状态 / 媒体检索漏 rejected / ACL schema 与 SQL 不对齐 / Compose 密码双源 | 领取 SQL 加 `pr.status='running'`；媒体检索补 `d.review_state<>'rejected'`；ACL 注释固化"document read + 子资产 deny"语义；Compose 密码单源说明（entrypoint 生成 secret 文件） |

## v1 → v2 变更摘要（对照第一/二轮评审）

| # | 评审问题 | 处理 |
| --- | --- | --- |
| 1 | 关键词召回 `ORDER BY rn DESC` 取最差结果 | 修正为 `order by rn`（升序）；三路召回排序统一复核 |
| 2 | 队列无阶段依赖，S1–S5 可乱序 | `stage_run.stage_order` + 领取 SQL 内前驱全成功检查（§6.3） |
| 3 | 单文件互斥 MVCC 竞态 | 部分唯一索引 `unique(document_id) where status='running'` 兜底 + 23505 冲突重试；`NOT EXISTS` 降级为预过滤（§6.3） |
| 4 | 租约无 fencing | `lease_token`(uuid) 全程 CAS：heartbeat/complete/fail 均校验 token；租约重置即换 token（§6.4） |
| 5 | S0 自举循环 | 拆出 **API 同步 bootstrap**（建 blob/document/run/stage，单事务）；S0 仅做深检与 .ppt 派生（§6.1） |
| 6 | 业务实体无 run 版本，新旧产物混用 | `slide/slide_object/media_occurrence` 增加 `run_id`；`source_document.active_run_id` 指向当前生效 run；检索/UI 一律走 active run（§5/§9） |
| 7 | "重建 pending" 违反 `unique(run_id, stage)` | 重跑语义改为**新建 run + 切换 active**，旧 run 置 `superseded` 待 GC；同 run 内不再复活 stage（§6.6） |
| 8 | DB+NAS"同一事务"不成立 | 改为显式 **发布 Saga**（staged→published→abandoned 状态机 + 补偿 + 对账）（§6.7） |
| 9 | source_document 唯一键吞掉密级差异 | 唯一索引加入 `classification, usage_scope, coalesce(copyright_status,'')`，不同密级各自成档（§5） |
| 10 | 全局 media 无法继承多源密级 | 拆分 `media_blob`（内容层，全局去重）/ `document_media`（业务层，按文档的密级/审核/用途）；occurrence 引用业务层（§5） |
| 11 | 检索 SQL 未执行 asset_acl | slide/media 检索均实现：部门匹配 ∨ admin ∨ read 授权，且非 deny；补文档级 review/format 硬过滤；`q_text_emb` NULL 时语义路显式关闭（§9） |
| 12 | 旧 .ppt 的 VBA 未检测 | S0 深检 OLE `_VBA_PROJECT_CUR` 流，命中先隔离再转换（§6.2） |
| 13 | Compose 不可运行 | 完整基线：密码 secret、连接 env、健康检查、迁移入口、internal 网络、非 root/只读根/tmpfs/资源限制/cap_drop（§12.1） |

同步落实的 P1：外键索引补齐、跨表一致性复合外键（run↔stage↔artifact、父子对象同页、occurrence 同页、组件同页）、`media_occurrence` NULL 去重（`unique nulls not distinct`）、完整状态机（§6.7）、部分成功/降级定义、素材检索可执行 SQL、模板"同 family 仅一个 active"部分唯一索引、OCR 归属与置信度模型。P2 落实：enum 顺序契约、标签触发器校验、usage 口径分层、验收指标口径定义、法务依赖从 M0 工期剥离。

---

## 1. 目标与范围

| 输入 | 产出资产 |
| --- | --- |
| ① 多元素素材 PPT | A. 原始素材库（图片/图标/图形组件/图表/表格/背景/色板字体） |
| ② PPT 规范示例 | B. 品牌规范库（`design-system.json`：自动草稿 + 人工确认 + 版本化） |
| ③ 优秀 PPT 范例 | C. 页面模板与版式库（页面资产 + 组件资产 + 内容容量 + 版本） |
| — | D. 检索库（文本/结构/视觉/语义/业务标签混合检索，支持页面与素材两级） |

明确不做（第一期）：动画自动复用、SmartArt 拆解重建、PPT 自动生成、跨部门共享（权限模型就位但先单部门试用）。

---

## 2. 总体架构：演进式单体

**第一期原则：能不引入的组件就不引入。** 队列用 PostgreSQL（部分唯一索引互斥 + `FOR UPDATE SKIP LOCKED`），存储用内网 NAS/本地卷（MinIO 可选），对象检测用 Pillow + imagehash（无 OpenCV），向量检索用精确扫描（无 HNSW）。规模化路径见 §12.4。

```text
        ┌────────────────────────────────────┐
        │ 上传(multipart) / 受控扫描目录(watch) │  ← 不接受任意服务器路径
        └──────────────────┬─────────────────┘
                           ▼
        ┌────────────────────────────────────┐
        │ API (FastAPI)                       │
        │  bootstrap(同步事务): 魔数初检→blob→  │
        │  document→run→stage×6 登记          │
        │  检索 / 审核 / 模板 / 规范 / 监控     │
        │  + 审核台 SSR(Jinja2+htmx)          │
        └───────┬──────────────┬─────────────┘
                │ stage 队列     │ 产物读
                ▼              ▼
  ┌───────────────────────────────────────────────┐
  │ Worker（单镜像多角色，PG 队列领取+租约+fencing）  │
  │  extract: s0,s1,s2   render: s3   index: s4,s5 │
  └───────┬───────────────────────┬───────────────┘
          ▼                       ▼
  ┌──────────────┐      ┌─────────────────────┐
  │ PostgreSQL16 │      │ 存储抽象 Storage      │
  │ pgvector     │      │ local/NAS/MinIO      │
  │ pg_trgm      │      │ 原件只读区+发布Saga   │
  └──────────────┘      └─────────────────────┘
```

关键设计：

- **bootstrap 与执行分离**：API 在单个 DB 事务内完成登记与任务创建，杜绝 S0 自举循环（§6.1）。
- **run 级版本**：所有解析产物归属 `pipeline_run`，`active_run_id` 指定生效版本，重跑 = 新 run（§6.6）。
- **互斥与 fencing 由数据库保证**：部分唯一索引 + lease token CAS，不依赖应用自觉（§6.3/§6.4）。
- **render 角色独立**：LibreOffice 进程重、易泄漏，独立容器与队列；解析 worker 与渲染/索引网络隔离（§12.1）。

---

## 3. 技术选型

| 层 | 选型 | 版本基线 | 说明 |
| --- | --- | --- | --- |
| 语言 | Python | 3.11+ | — |
| 页面解析 | python-pptx | 1.0.x（钉死小版本） | 文本/形状/表格/基础图表；主题继承、动画、SmartArt、OLE 由 OOXML 层补充 |
| OOXML/OLE 解析 | zipfile + lxml + olefile | — | olefile 用于 OLE 魔数分流与流检测（加密/VBA）；带解压限额（§7.3） |
| 渲染 | LibreOffice Headless | 生产钉死单一版本（Linux 容器） | 仅作基线预览；黄金样本 SSIM 度量（§8） |
| PDF→PNG | **候选**：pypdfium2（Apache-2.0）首选；PyMuPDF（AGPL）备选 | — | M0 性能对比后定版；PyMuPDF 待法务确认 |
| 图像处理 | Pillow + imagehash | — | SHA-256 + pHash 去重；不引入 OpenCV |
| OCR | 候选：RapidOCR（onnxruntime，Apache-2.0） | — | 无可提取文本页/截图/图表图补偿；结果带置信度（§6.9） |
| 文本向量 | 候选冻结：bge-large-zh-v1.5（1024 维） | 权重入镜像 | M0 评测后冻结版本，随行记录 `text_emb_model` |
| 图像向量 | 候选冻结：Chinese-CLIP ViT-B/16（512 维） | 权重入镜像 | 评测不达标则降级为仅以图搜图 |
| 元数据/向量/队列 | PostgreSQL 16 + pgvector + pg_trgm | pgvector/pgvector:pg16 | 精确扫描起步；HNSW 暂缓（§12.4） |
| 对象存储 | NAS 只读共享/本地卷起步；MinIO 可选 | — | MinIO 为 AGPL，引入前过法务 |
| API/前端 | FastAPI + Pydantic v2 + SQLAlchemy 2 + Alembic；审核台 Jinja2 + htmx | — | Vue 延后到流程稳定 |
| 部署 | Docker Compose → 可迁 K8s | — | 镜像全部入内网仓库 |

商业备选：Aspose.Slides 不预先集成。仅当 M0 黄金样本验证中开源链路失败率/视觉差异超标时，经统一 `Parser/Renderer` 接口替换 S2/S3 实现（`ppt_pipeline/parsers/`），数据结构不变。

### 3.1 授权清单（M0 发起法务确认，结论为 M1 准入门槛）

| 组件 | 许可 | 风险与动作 |
| --- | --- | --- |
| PyMuPDF | AGPL-3.0 | **待确认**；不通过则用 pypdfium2 |
| MinIO（如引入） | AGPL-3.0 | **待确认**；不通过则 NAS/本地卷 |
| LibreOffice | MPL-2.0 | 确认镜像分发方式 |
| 企业字体 | 商业授权 | 确认"打包进镜像"是否在授权范围；备替代字体映射表 |
| BGE / Chinese-CLIP / RapidOCR 权重 | MIT/Apache 系 | 核对模型卡用途限制 |
| 其余依赖 | MIT/BSD/PostgreSQL/Apache | 低风险 |

---

## 4. 工程结构

```text
ppt-pipeline/
├── pyproject.toml
├── Dockerfile                  # 单镜像（api/worker 共用，command 区分；含 LibreOffice+字体+模型）
├── docker-compose.yml
├── entrypoint.sh               # app 角色: alembic upgrade head && uvicorn
├── secrets/                    # db_password 等（git 忽略，示例文件提供）
├── ppt_pipeline/
│   ├── core/                   # 配置、日志(trace_id)、存储抽象(Storage)、ID 生成
│   ├── parsers/                # Parser/Renderer 接口 + pptx 实现（未来 aspose）
│   ├── extract/
│   │   ├── s0_intake.py        # 深检（VBA/加密/OLE）、.ppt 派生、元信息回填
│   │   ├── s1_ooxml.py         # 受限解包、媒体/主题/母版/版式/嵌入清单
│   │   ├── s2_structure.py     # 页级/对象级解析
│   │   ├── s3_render.py        # LO→PDF→PNG、媒体预览
│   │   ├── s4_index.py         # 向量、OCR、聚类去重、规范草稿
│   │   └── s5_report.py        # 盘点报告（usage 两口径聚合）
│   ├── pipeline/               # bootstrap、领取/租约/fencing CAS、Saga 发布、状态机、GC
│   ├── search/                 # 多路召回 + RRF（slide/media 两套）
│   ├── design_system.py        # 规范草稿（confidence/evidence/extractor_version）
│   ├── db/                     # 模型 + Alembic + triggers/（标签校验、审计不可变）
│   ├── api/                    # REST + 审核台 SSR
│   └── worker/                 # worker 入口（--stages 角色领取）
├── templates/                  # Jinja2 审核台页面
├── alembic/
├── models/                     # 冻结权重 + MANIFEST.json（随镜像构建 COPY）
├── golden-samples/             # 黄金样本与 SSIM 基线报告
└── tests/                      # 含队列并发/租约 fencing/重跑隔离的集成测试
```

---

## 5. 数据模型（按序可执行 DDL）

> 设计要点：
> - 内容与业务分离：`file_blob`/`source_document`（文件）、`media_blob`/`document_media`（媒体）；
> - 执行历史与业务数据分离：`pipeline_run`/`stage_run`/`artifact`；业务表带 `run_id`，`active_run_id` 指定生效版本；
> - 跨表一致性用**复合外键**落地（stage 归属 run、父子同页、occurrence 同页、artifact 归属 stage）；
> - 密级 `classification_level` enum **按声明顺序比较（public < internal < confidential < secret），此顺序为接口契约，迁移只允许追加**；
> - 子资产密级 `NULL = 继承源文档`（有效密级 = `coalesce(子, 母)`）。

```sql
create extension if not exists vector;
create extension if not exists pg_trgm;

create type classification_level as enum ('public','internal','confidential','secret');

-- ============ 身份与权限 ============
create table department (
  id    bigint generated always as identity primary key,
  code  text unique not null,
  name  text not null
);

create table app_user (
  id            bigint generated always as identity primary key,
  username      text unique not null,
  display_name  text not null,
  department_id bigint not null references department(id),
  role          text not null default 'reviewer'
                check (role in ('viewer','reviewer','admin')),
  is_active     boolean not null default true,
  created_at    timestamptz not null default now()
);
create index on app_user (department_id);

-- ============ 文件内容层与业务层 ============
create table file_blob (
  id          bigint generated always as identity primary key,
  sha256      char(64) unique not null check (sha256 ~ '^[0-9a-f]{64}$'),
  size_bytes  bigint not null check (size_bytes > 0),
  object_key  text not null,
  stored_at   timestamptz not null default now()
);

create table source_document (
  id             bigint generated always as identity primary key,
  blob_id        bigint not null references file_blob(id),
  derived_blob_id bigint references file_blob(id),   -- .ppt 转换派生的 .pptx（S0 回填）
  active_run_id  bigint,                             -- 当前生效 run；FK 在 pipeline_run 建表后补加
  file_name      text not null,
  category       text not null check (category in ('material','spec','sample')),
  owner_department_id bigint not null references department(id),
  classification classification_level not null default 'internal',
  usage_scope    text not null default 'internal_use'
                 check (usage_scope in ('internal_use','client_delivery','public','restricted')),
  copyright_status text,
  slide_count    int check (slide_count is null or slide_count > 0),
  page_width_emu bigint check (page_width_emu is null or page_width_emu > 0),
  page_height_emu bigint check (page_height_emu is null or page_height_emu > 0),
  producer       text,
  converted      boolean not null default false,
  format_status  text not null default 'ok'
                 check (format_status in ('ok','macro_quarantined','encrypted','corrupt')),
  tags           text[] not null default '{}',
  review_state   text not null default 'pending'
                 check (review_state in ('pending','in_review','approved','rejected')),
  created_at     timestamptz not null default now()
);
-- 同一内容：部门/类别/密级/用途/版权任一不同即独立成档（避免降密合并）
create unique index uq_source_document
  on source_document (blob_id, category, owner_department_id, classification, usage_scope,
                      coalesce(copyright_status, ''));
create index on source_document (category);
create index on source_document (owner_department_id);
create index on source_document using gin (tags);

create table ingest_record (
  id           bigint generated always as identity primary key,
  document_id  bigint not null references source_document(id),
  ingested_by  bigint not null references app_user(id),
  source_kind  text not null check (source_kind in ('upload','watch_dir','admin_import')),
  source_path  text,
  ingested_at  timestamptz not null default now()
);
create index on ingest_record (document_id);
create index on ingest_record (ingested_by);

-- ============ 执行层：run / stage / artifact ============
create table pipeline_run (
  id               bigint generated always as identity primary key,
  document_id      bigint not null references source_document(id),
  pipeline_version text not null,
  config_hash      text not null,
  status           text not null default 'running'
                   check (status in ('running','succeeded','failed','cancelled','superseded')),
  started_at       timestamptz not null default now(),
  finished_at      timestamptz,
  -- 幂等说明：同一参数重跑不是复用旧 run，而是"沿用 pending 状态"或建新 run，
  -- 用 (status, attempt 语义) 而非唯一约束表达——见 §6.5；
  -- 唯一约束仅防止 bootstrap 并发双击产生的重复 run，reprocess 走显式路径：
  unique (document_id, pipeline_version, config_hash) where status in ('running','succeeded'),
  unique (id, document_id)                                -- 支撑下游复合外键
);
create index on pipeline_run (document_id, status);

create table stage_run (
  id               bigint generated always as identity primary key,
  run_id           bigint not null references pipeline_run(id) on delete cascade,
  document_id      bigint not null references source_document(id),
  stage            text not null check (stage in ('s0','s1','s2','s3','s4','s5')),
  stage_order      int not null check (stage_order between 0 and 5),
  attempt          int not null default 0 check (attempt >= 0),
  status           text not null default 'pending'
                   check (status in ('pending','running','succeeded','failed','cancelled')),
  lease_owner      text,
  lease_token      uuid,
  lease_expires_at timestamptz,
  error_code       text,
  error_detail     text,
  warnings         jsonb not null default '[]',   -- 页级失败/降级明细（阶段仍成功时）
  started_at       timestamptz,
  finished_at      timestamptz,
  unique (run_id, stage),
  unique (run_id, stage_order),
  unique (id, run_id),                             -- 支撑 artifact 复合外键
  foreign key (run_id, document_id) references pipeline_run (id, document_id)  -- stage 必属该 run 的文档
);
create index on stage_run (status, stage);
create index on stage_run (document_id, status);
-- ★ 文件级互斥的真保证：同一文档同时至多一个 running stage（并发写入冲突 → 23505 → 重试）
create unique index uq_stage_running_per_document
  on stage_run (document_id) where status = 'running';

create table artifact (
  id            bigint generated always as identity primary key,
  run_id        bigint not null references pipeline_run(id) on delete cascade,
  stage_run_id  bigint not null references stage_run(id),
  kind          text not null,                    -- slide_png/thumb/media/json/report
  tmp_key       text,                             -- Saga 中转 key
  object_key    text not null,                    -- 正式 key
  sha256        char(64) not null,
  size_bytes    bigint not null check (size_bytes >= 0),
  generator     text not null,                    -- python-pptx 1.0.2 / soffice 7.6.4 / bge-large-zh-v1.5
  status        text not null default 'staged'
                check (status in ('staged','published','abandoned')),
  created_at    timestamptz not null default now(),
  unique (stage_run_id, kind, object_key),        -- 同 stage 幂等登记
  foreign key (stage_run_id, run_id) references stage_run (id, run_id)  -- artifact 必属该 stage 的 run
);
create index on artifact (kind, status);
create index on artifact (run_id);

alter table source_document
  add constraint fk_active_run foreign key (id, active_run_id)
  references pipeline_run (document_id, id);   -- 复合外键：active run 必属本文档

-- ============ 媒体：内容层（全局去重）/ 业务层（按文档） ============
create table media_blob (
  id             bigint generated always as identity primary key,
  sha256         char(64) unique not null check (sha256 ~ '^[0-9a-f]{64}$'),
  phash          char(16),
  fmt            text not null,
  mime_type      text not null,
  width int, height int, has_alpha boolean,
  color_stats    jsonb,
  object_key     text not null,                   -- 原格式保留
  preview_key    text,
  img_emb        vector(512),                     -- 内容向量（内容派生，全局共享）
  ocr_text       text,
  ocr_confidence numeric(4,3) check (ocr_confidence is null or ocr_confidence between 0 and 1),
  caption        text,
  dedup_group    int                              -- 内容相似聚类（按媒体类型分阈值）
);
create index on media_blob (phash);
create index on media_blob (dedup_group);
create index on media_blob using gin (ocr_text gin_trgm_ops);

create table document_media (
  id             bigint generated always as identity primary key,
  document_id    bigint not null references source_document(id),
  blob_id        bigint not null references media_blob(id),
  classification classification_level,            -- NULL = 继承文档
  usage_scope    text check (usage_scope is null or usage_scope in ('internal_use','client_delivery','public','restricted')),
  copyright_status text,
  review_state   text not null default 'pending'
                 check (review_state in ('pending','in_review','approved','rejected')),
  usage_policy   text not null default 'reference_only'
                 check (usage_policy in ('reusable','reference_only','forbidden')),
  generation_allowed boolean not null default false,
  tags           text[] not null default '{}',
  created_at     timestamptz not null default now(),
  unique (document_id, blob_id)                   -- 同文档同内容一条；跨文档各自成档（密级/审核独立）
);
create index on document_media (blob_id);
create index on document_media (review_state, usage_policy);
create index on document_media using gin (tags);

-- ============ 页面层（run 版本化） ============
create table slide (
  id             bigint generated always as identity primary key,
  run_id         bigint not null references pipeline_run(id) on delete cascade,
  document_id    bigint not null references source_document(id),
  slide_no       int not null check (slide_no > 0),
  is_hidden      boolean not null default false,
  layout_part    text,
  master_part    text,
  title          text,
  texts          jsonb not null default '[]',
  search_text    text not null default '',
  notes_text     text,                            -- 默认不拼接进 search_text/向量（§7.4）
  struct         jsonb not null default '{}',
  png_key        text,
  thumb_key      text,
  text_emb       vector(1024),
  text_emb_model text,
  img_emb        vector(512),
  img_emb_model  text,
  review_state   text not null default 'pending'
                 check (review_state in ('pending','in_review','approved','rejected')),
  usage_policy   text not null default 'reference_only'
                 check (usage_policy in ('reusable','reference_only','forbidden')),
  generation_allowed boolean not null default false,
  template_candidate boolean not null default false,
  slide_type     text,
  quality        int check (quality is null or quality between 1 and 5),
  classification classification_level,            -- NULL = 继承文档
  tags           text[] not null default '{}',
  unique (run_id, document_id, slide_no),
  foreign key (run_id, document_id) references pipeline_run (id, document_id)
);
create index on slide (document_id, run_id);      -- active run 查询入口
create index on slide (review_state, usage_policy);
create index on slide using gin (tags);
create index on slide using gin (search_text gin_trgm_ops);
-- 向量索引暂缓：万页级精确扫描足够（§12.4）

-- ============ 对象层（父子同页复合外键） ============
create table slide_object (
  id           bigint generated always as identity primary key,
  slide_id     bigint not null references slide(id) on delete cascade,
  parent_id    bigint,
  seq          int not null check (seq >= 0),
  obj_type     text not null
               check (obj_type in ('placeholder','textbox','autoshape','picture','table',
                                   'chart','group','smartart','ole','media','note_shape')),
  name         text,
  placeholder_type text,
  x_emu bigint not null check (x_emu >= 0), y_emu bigint not null check (y_emu >= 0),
  w_emu bigint not null check (w_emu >= 0), h_emu bigint not null check (h_emu >= 0),
  rotation     numeric,
  style        jsonb not null default '{}',
  content      jsonb not null default '{}',
  editable     boolean not null default true,
  unique (slide_id, seq),
  unique (id, slide_id),                          -- 支撑复合外键
  foreign key (parent_id, slide_id) references slide_object (id, slide_id)  -- 父必同页
);
create index on slide_object (slide_id);
create index on slide_object (obj_type);
create index on slide_object (parent_id);

-- ============ 媒体出现关系（run 版本化；对象同页复合外键；NULL 参与去重） ============
create table media_occurrence (
  id          bigint generated always as identity primary key,
  slide_id    bigint not null references slide(id) on delete cascade,
  object_id   bigint,                             -- NULL = 页面背景/母版引用
  document_media_id bigint not null references document_media(id),
  rel_id      text,
  part_uri    text,
  role        text not null default 'content'
              check (role in ('content','background','master','layout','fill','poster')),
  crop        jsonb,
  unique nulls not distinct (slide_id, object_id, document_media_id, role),  -- 背景(NULL)也不重复
  foreign key (object_id, slide_id) references slide_object (id, slide_id)   -- 对象必属该页
);
create index on media_occurrence (document_media_id);
create index on media_occurrence (slide_id);
create index on media_occurrence (object_id);

-- ============ 组件资产 ============
create table asset_component (
  id             bigint generated always as identity primary key,
  source_slide_id   bigint not null references slide(id),
  source_object_id  bigint not null,
  name           text not null,
  kind           text not null check (kind in ('kpi_card','diagram','timeline','arrow_set','badge','other')),
  preview_key    text,
  structure      jsonb not null default '{}',
  review_state   text not null default 'pending'
                 check (review_state in ('pending','in_review','approved','rejected')),
  usage_policy   text not null default 'reference_only'
                 check (usage_policy in ('reusable','reference_only','forbidden')),
  generation_allowed boolean not null default false,
  classification classification_level,
  tags           text[] not null default '{}',
  created_at     timestamptz not null default now(),
  foreign key (source_object_id, source_slide_id) references slide_object (id, slide_id)  -- 组件对象必属该页
);
create index on asset_component (source_slide_id);
create index on asset_component (review_state, usage_policy);

-- ============ 模板库（多版本；同 family 仅一个 active） ============
create table page_template (
  id                bigint generated always as identity primary key,
  family_key        text not null,
  version           int not null check (version > 0),
  source_slide_id   bigint not null references slide(id),
  layout            jsonb not null,
  capacity          jsonb not null,
  status            text not null default 'draft'
                    check (status in ('draft','active','retired')),
  supersedes_id     bigint references page_template(id),
  approved_by       bigint references app_user(id),
  approved_at       timestamptz,
  effective_from    timestamptz,
  effective_to      timestamptz,
  created_at        timestamptz not null default now(),
  unique (family_key, version),
  check (effective_to is null or effective_from is null or effective_to > effective_from)
);
create unique index uq_template_one_active on page_template (family_key) where status = 'active';
create index on page_template (source_slide_id);
create index on page_template (supersedes_id);
create index on page_template (approved_by);

-- ============ 规范库 ============
create table design_system (
  id           bigint generated always as identity primary key,
  version      int not null check (version > 0),
  schema_version int not null default 1,
  payload      jsonb not null,
  draft_diff   jsonb,
  status       text not null default 'draft' check (status in ('draft','confirmed','retired')),
  confirmed_by bigint references app_user(id),
  confirmed_at timestamptz,
  created_at   timestamptz not null default now(),
  unique (version)
);
create index on design_system (confirmed_by);

create table design_system_source (
  design_system_id bigint not null references design_system(id) on delete cascade,
  document_id      bigint not null references source_document(id),
  primary key (design_system_id, document_id)
);
create index on design_system_source (document_id);

-- ============ 受控标签词表 ============
create table tag (
  id       bigint generated always as identity primary key,
  scope    text not null check (scope in ('slide_type','business','media_type','style')),
  key      text not null,
  label    text not null,
  aliases  text[] not null default '{}',
  parent_id bigint,
  unique (scope, key),
  unique (id, scope),                          -- 支撑复合外键（父标签同域）
  foreign key (parent_id, scope) references tag (id, scope)
);
create index on tag (parent_id);
-- 各资产表 tags 中的 key 必须存在于 tag 表：由 db/triggers/validate_tags.sql 触发器强制
-- （text[] 无法直接外键；触发器覆盖 API/脚本/批量导入所有写入路径）

-- ============ 使用记录（口径分层：见 §6.9） ============
create table asset_usage_log (
  id         bigint generated always as identity primary key,
  asset_type text not null check (asset_type in ('slide','media','component','template')),
  asset_id   bigint not null,                  -- 多态引用，应用层保证
  action     text not null check (action in ('search_click','export','template_ref','view')),
  actor_id   bigint references app_user(id),
  at         timestamptz not null default now()
);
create index on asset_usage_log (asset_type, asset_id, at);
create index on asset_usage_log (actor_id);

-- ============ ACL 覆盖与审计 ============
-- 默认权限：资产归属源文档部门；本表只登记显式覆盖。
-- 语义约定：read 仅支持 document 粒度（子资产授权一律通过文档授权表达）；
-- deny 支持 document/slide/media 粒度（就近收紧）。schema 与检索 SQL 保持一致，
-- 如未来需要子资产 read，需同步修改 §9 的 allow 判定。
create table asset_acl (
  id            bigint generated always as identity primary key,
  asset_type    text not null check (asset_type in ('document','slide','media','component','template')),
  asset_id      bigint not null,
  department_id bigint not null references department(id),
  permission    text not null check (permission in ('deny','read')),
  created_by    bigint not null references app_user(id),
  created_at    timestamptz not null default now(),
  unique (asset_type, asset_id, department_id)
);
create index on asset_acl (department_id, asset_type);

create table audit_event (
  id         bigint generated always as identity primary key,
  actor_id   bigint not null references app_user(id),
  action     text not null,
  asset_type text not null,
  asset_id   bigint not null,
  old_value  jsonb,
  new_value  jsonb,
  trace_id   text,
  at         timestamptz not null default now()
);
create index on audit_event (asset_type, asset_id, at);
create index on audit_event (actor_id);
-- 不可变：应用角色无 UPDATE/DELETE 权限（迁移中 revoke），另加触发器兜底拒绝
```

> 本节用于评审数据模型，正式执行以 Alembic 迁移（含触发器与 revoke）为准。

---

## 6. 流水线设计

### 6.1 Bootstrap（API 内，单 DB 事务，同步）

解决 v2 的 S0 自举循环：任务创建先于执行，且不依赖任何 worker。

```text
1. 魔数初检（读文件头，olefile/zipfile）：
   PK\x03\x04            → OOXML 走 2
   OLE CFB + EncryptionInfo 流 → format_status='encrypted'，登记后终止（不建 run）
   OLE CFB 其他          → 旧 .ppt，走 2
   其他                  → format_status='corrupt'，登记后终止（不建 run）
2. sha256 → file_blob upsert（已存在则复用）
3. source_document insert（业务属性：类别/部门/密级/用途/版权必填）
   ※ 唯一索引冲突 = 同参数已登记 → 复用既有 document，仅追加 ingest_record
4. ingest_record insert（who/where/when）
5. pipeline_run insert（当前 pipeline_version + config_hash）
   ※ 幂等键冲突 = 同参数 run 已存在 → 直接返回该 run（不重复创建）
6. source_document.active_run_id = 新 run
7. 6 条 stage_run insert（s0..s5，stage_order 0..5，status='pending'）
```

加密文件的人工密码流程：解密产生新 blob → 走正常 bootstrap（记 ingest_record 来源 `admin_import`）。

### 6.2 各阶段职责

| 阶段 | order | 角色 | 内容 | 终止/降级 |
| --- | --- | --- | --- | --- |
| s0 intake | 0 | extract | 深检：ZIP 内 `ppt/vbaProject.bin`、OLE `_VBA_PROJECT_CUR` 流（**旧 .ppt 宏同样拦截**）→ 隔离区 + run 置 cancelled；`.ppt` → LO 转换派生 `.pptx`（`derived_blob_id` 回填）；页数/尺寸/producer 回填 source_document | 宏/损坏 → run cancelled，下游 stage 全部 cancelled |
| s1 unpack | 1 | extract | 受限解包（限额见 §7.3）；媒体/主题/母版/版式/图表/嵌入/备注清单 → 文件级 JSON artifact | 可重试 |
| s2 structure | 2 | extract | 页/对象解析入 `slide`/`slide_object`/`media_occurrence`（带 run_id） | 页级失败 → `stage_run.warnings` + `slide.struct.parse_errors`，阶段仍 succeeded |
| s3 render | 3 | render | LO 整文件转 PDF（超时 → stage 失败重试）；pypdfium2 按页出 PNG（≥1600px）+ 缩略图；单页渲染失败跳过记 warnings | PDF 转换重试上限后 failed |
| s4 index | 4 | index | 文本/视觉向量（行内记录模型 ID）、OCR、phash 聚类、规范草稿（confidence/evidence） | 单页向量失败 → 该页 emb 为 NULL + warning，阶段 succeeded；补算走新 run |
| s5 report | 5 | index | 盘点报告（usage 两口径聚合：出现次数 occurrence / 复用动作 usage_log） | 依赖 s0–s4 全 succeeded |

worker 角色：extract 领 `s0,s1,s2`；render 领 `s3`；index 领 `s4,s5`（report 轻量，与 index 同角色，修正 v2 表述不一致）。

### 6.3 任务领取（互斥 + 依赖 + fencing 前置）

互斥的真保证是部分唯一索引 `uq_stage_running_per_document`；并发冲突以 23505 暴露，worker 捕获后回滚重试。`NOT EXISTS running` 仅作廉价预过滤。

```sql
-- 领取（单语句事务）
with candidate as (
  select sr.id
  from stage_run sr
  where sr.status = 'pending'
    and sr.stage = any(cast(:stages as text[]))
    -- ★ 阶段依赖：同 run 内前驱必须全部 succeeded，且父 run 仍在 running
    and exists (select 1 from pipeline_run pr
                where pr.id = sr.run_id and pr.status = 'running')
    and not exists (
      select 1 from stage_run pre
      where pre.run_id = sr.run_id
        and pre.stage_order < sr.stage_order
        and pre.status <> 'succeeded')
  order by sr.id
  limit 1
  for update skip locked
)
update stage_run sr
set status = 'running',
    lease_token  = gen_random_uuid(),
    lease_owner  = :worker_id,
    lease_expires_at = now() + interval '15 minutes',
    attempt = sr.attempt + 1,
    started_at = now()
from candidate c
where sr.id = c.id
returning sr.id, sr.run_id, sr.document_id, sr.stage, sr.attempt, sr.lease_token;
-- 并发同文档领取 → unique_violation(23505) → 回滚、随机退避后重试（这是互斥语义的一部分，不是错误路径）
```

### 6.4 租约、心跳与 CAS 完成（fencing）

所有状态推进均校验 `lease_token`；租约被接管后 token 已换，旧 worker 的任何写入必然 0 行——这就是 fencing。

```sql
-- 心跳（长任务周期执行）
update stage_run set lease_expires_at = now() + interval '15 minutes'
where id = :id and lease_token = :token and status = 'running';
-- 0 行 → 已被接管：worker 必须立即终止，放弃全部未发布产物

-- 崩溃恢复（任意 worker 周期执行；重置即换 token）
update stage_run
set status = 'pending', lease_token = null, lease_owner = null,
    lease_expires_at = null, error_code = 'lease_expired'
where status = 'running' and lease_expires_at < now();

-- 失败（CAS）
update stage_run
set status = case when attempt >= :max_attempts then 'failed' else 'pending' end,
    error_code = :code, error_detail = :detail,
    lease_owner = null, lease_expires_at = null,   -- 保留 lease_token 供失败路径排查
    finished_at = now()
where id = :id and lease_token = :token and status = 'running';
-- attempt 达上限 → failed；编排器同事务内：
--   下游 pending → cancelled；run → failed, finished_at = now()

-- 取消（用户/编排器）：先原子接管租约（换 token），旧 worker 后续 CAS 必然失败
update stage_run
set status = 'cancelled', lease_token = gen_random_uuid(),
    lease_owner = 'orchestrator', finished_at = now(),
    error_code = :reason
where id = :id and status in ('pending','running');
-- 编排器同事务内：同 run 其余 pending → cancelled；run → cancelled

-- 成功（CAS；与最后一批 artifact 发布同事务，见 §6.7）
update stage_run set status = 'succeeded', finished_at = now()
where id = :id and lease_token = :token and status = 'running';
-- 编排器同事务内：若 s5 成功 → run → succeeded
```

### 6.5 业务写入围栏（fencing 覆盖副作用）

stage 状态 CAS 只保护状态本身；业务数据与产物按以下规则纳入 fencing：

1. **业务写事务必须携带租约校验**：slide/slide_object/media_occurrence 等所有 run 产物的写事务，先在同事务内执行
   `select id from stage_run where id=:sid and lease_token=:token and status='running' for update`
   （0 行 → 立即回滚放弃），再写业务表。worker 数据访问层统一封装此校验，禁止裸写。
2. **产物 key 天然含 run_id**（`slides/{run_id}/{no}.png` 等）：reprocess 产生新 run 即新命名空间，
   旧 worker 即便写存储也写不进新 run 的 key；旧 run 产物由 GC 回收。
3. **artifact 发布 CAS**（§6.7 步骤 4）同时校验 stage 租约：
   `update artifact ... where id=:aid and status='staged' and exists (select 1 from stage_run sr where sr.id=:sid and sr.lease_token=:token and sr.status='running')`。
4. 结果：失租 worker 最坏情况是把数据写进**已被 supersede 的旧 run**（检索/UI 走 active_run_id 不可见，GC 清理），不可能污染新 run 或覆盖新产物。

### 6.6 重跑语义：新 run + active 切换（消除 v2 矛盾）

```text
POST /api/documents/{id}/reprocess
1. 若同参数 run 存在且 status='running' → 幂等返回该 run（bootstrap 部分唯一索引保证不重复创建）
   否则：新建 pipeline_run（同参数旧 run 已是 superseded/failed/cancelled，
        部分唯一索引 where status in ('running','succeeded') 不阻塞新 run）
2. source_document.active_run_id → 新 run（原子）
3. 旧 run 置 superseded
效果：
- 检索/UI 一律 join active_run_id → 即刻只见新 run 数据，天然杜绝新旧混用
- 旧 run 的 slide/object/occurrence/artifact 保留供审计，GC 任务在保留期（默认 30 天）后删除
- 不存在"同 run 内复活 stage"，unique(run_id, stage) 永不违反
- "from=s3" 参数忽略：S0–S2 重解析成本（秒级）远低于跨 run 部分继承的状态维护成本，
  全量重建是最简单且不可出错的选择（显式决策，非遗漏）
```

GC 规则：删除 `pipeline_run.status in ('superseded','cancelled')` 且 `finished_at < now()-保留期` 的 run 级业务数据（slide 级联 slide_object/media_occurrence）及其 artifact（含 abandoned 残留）；`document_media` 按文档保留（跨 run 持久，见 §6.9）。

### 6.7 产物发布 Saga（替代"同一事务"的不实表述）

DB 事务无法覆盖存储操作，显式状态机 + 补偿 + 对账：

```text
1. 存储写 tmp 对象（tmp/{stage_run_id}/{uuid}.{ext}）
2. DB 事务 A：insert artifact(status='staged', tmp_key, object_key=正式key, sha256, generator)
3. 存储操作：copy tmp → 正式 key（或 NAS rename）
4. DB 事务 B（CAS，校验 lease_token）：
   a. update artifact set status='published' where id=:aid and status='staged';
   b. 全部产物已 published 且工作完成 → update stage_run → succeeded（§6.4）
5. 清理 tmp 对象
失败补偿：
- 事务 B 前失败 → 删 tmp，artifact 置 abandoned，stage 走失败/重试
- 事务 B 后清理 tmp 失败 → 对账任务回收（幂等）
对账任务（周期）：
- staged 超时未 published → 置 abandoned 并清理对象
- published 但对象缺失/哈希不符 → 告警 + 触发 reprocess
- 存储孤儿对象（无 artifact 行）→ 回收
```

### 6.8 状态机（正式定义）

**pipeline_run**

| 当前 | 目标 | 触发（守卫） |
| --- | --- | --- |
| running | succeeded | s5 succeeded（编排器在 s5 CAS 同事务内） |
| running | failed | 任一 stage 达 attempt 上限 failed（同事务级联下游 cancelled） |
| running | cancelled | 用户取消（接管租约换 token 后级联，§6.4）；s0 发现宏/损坏（级联下游 cancelled） |
| running/succeeded/failed | superseded | reprocess 创建新 run 并切换 active |
| superseded → （GC 后物理删除） | — | 保留期满，GC 删除 run 级数据与产物 |

**stage_run**

| 当前 | 目标 | 触发（守卫） |
| --- | --- | --- |
| pending | running | 领取成功（父 run running + 依赖满足 + 互斥索引不冲突） |
| running | succeeded | worker CAS（lease_token 匹配；产物发布完成） |
| running | pending | 失败且 attempt < max（CAS） |
| running | failed | 失败且 attempt ≥ max（CAS） |
| running | pending | 租约过期被重置（token 作废 → 原 worker 后续 CAS 必然 0 行） |
| running/pending | cancelled | 编排器接管租约（换 token）后置 cancelled；上游终态失败 / run cancelled / run superseded 级联 |

部分成功语义：页级失败/单页渲染失败/单页向量失败 → `stage_run.warnings` + 阶段仍 `succeeded`；补数统一走 reprocess（新 run），不在线修补。

### 6.9 媒体两层协作与 usage 口径

```text
S1/S2 提取:
  media_blob: 内容层 upsert（sha256 全局唯一）—— phash/尺寸/向量/OCR/caption/dedup_group 均内容派生
  document_media: (document_id, blob_id) upsert —— 密级/审核/用途/版权/generation_allowed 按文档独立
  重跑时同 blob 复现 → document_media 既有审核结论自动保留（不丢失人工成果）
  文档不再引用的 document_media → 标记 orphan（S5 报告列出，人工决定清理）
usage 两口径（S5 报告分列）：
  出现次数 = media_occurrence 计数（素材在源文件中被引用多少次）
  复用动作 = asset_usage_log 按 action 分组（search_click/export/template_ref/view）
```

OCR 归属：整页无可提取文本时，对页面 PNG 做 OCR → 存 `slide.struct.ocr_text`（含平均置信度，标注来源 `page_ocr`）；图片素材 OCR 存 `media_blob.ocr_text/ocr_confidence`（内容派生，全局共享）。两者拼接进检索文本的开关独立配置。

---

## 7. 安全设计

### 7.1 认证与权限
- 内网 SSO/LDAP（M0 确认方案），本地账号兜底；RBAC：viewer/reviewer/admin。
- 有效权限算法（应用层统一函数 + 检索 SQL 内联）：
  - 可见部门集 = 源文档部门 ∨ admin ∨ `asset_acl(document, read)` 授予；
  - 可见密级 = `coalesce(资产密级, 文档密级) ≤ 用户密级`；
  - 显式 `deny`（document/slide/media 级）一票否决；
  - deny 只能由 admin 撤销（审计）。
- 检索 SQL 强制内联上述条件（§9），不依赖前端过滤。

### 7.2 输入面收敛
- 仅 multipart 上传 + 预配置 watch 目录；不提供任意服务器路径参数；watch 扫描对路径做规范化 + 根目录前缀校验。
- 上传大小上限；来源登记必填（部门/类别/密级初值）。

### 7.3 恶意文档防护
- 宏检测双通道：ZIP 内 `ppt/vbaProject.bin`（.pptm/.pptx）+ OLE `_VBA_PROJECT_CUR` 流（旧 .ppt，**转换前拦截**）→ 隔离区（独立只读目录）+ 内网 AV；OLE/嵌入对象提取后先过 AV。
- 解析 worker：`internal: true` 网络（无外网出口）、非 root、只读根文件系统 + tmpfs、cap_drop ALL、CPU/内存/时限限制（Compose 落地见 §12.1）。
- 解包限额：解压总体积 ≤2GB、单文件 ≤512MB、XML 节点数上限、entry 名规范化后必须落在目标目录内（防路径穿越）。
- 代码无外部 HTTP 出口：CI 静态检查 + 网络层（internal 网络）双重保障。

### 7.4 敏感内容
- 备注单独存储，默认不入 `search_text`/向量，详情页按 ACL 展示（可配置开启拼接，需 admin）。
- `secret` 默认排除检索，仅密级允许且显式勾选可见。

---

## 8. 渲染治理（LibreOffice）

### 8.1 隔离与自愈
- 每任务：`-env:UserInstallation=file:///tmp/lo-<uuid>` + 独立输出目录；硬超时（单文件 120s 默认）后 kill 进程树，profile 随任务清理。
- 每 worker：N=20 任务后进程自愈重启；render 角色独立容器/队列；单容器内 LO 串行，横向扩容器。

### 8.2 黄金样本基线（M0 关键交付）
1. 30–50 页样本：中文字体、母版继承、SmartArt、图表、透明度、EMF/WMF、WPS、旧 `.ppt`。
2. PowerPoint/WPS 人工导出基准 PNG。
3. 生产链路（钉版 LO + 企业字体镜像）渲染，计算 SSIM + 像素差异。
4. **阈值 M0 冻结（初始建议 SSIM ≥ 0.90，冻结后即为验收值）**；低于阈值的对象类型进"预览仅供参考"清单。基线报告存 `golden-samples/`；LO/字体/镜像任何变更重跑基线。

### 8.3 平台策略
- 生产钉死 Linux 容器 + 固定 LO 版本 + 固定字体包；Windows 仅本地开发，渲染结果不作质量依据。
- 企业字体授权确认后打入镜像并 `fc-cache`；未授权字体不进镜像，维护替代映射表并在预览标注。
- EMF/WMF/SVG 预览独立适配层，转失败仅记录（原件保留）。

---

## 9. 检索设计（多路召回 + RRF）

原则：硬过滤先行（active run / 审核状态 / 密级 / ACL / 文档状态），三路独立召回（每路 limit 100），RRF 融合（只依赖排名）；纯文本查询关闭视觉路，反之关闭文本路；权重为 M3 评测项。

参数绑定约定（SQLAlchemy `text()` + `bindparam` 显式类型）：SQL 文本统一用 `cast(:param as type)` 而非 `:param::type`（后者会被 SQLAlchemy 误解析为绑定参数名）；`:q_text_emb → vector(1024)`、`:q_img_emb → vector(512)`、`:user_clearance → classification_level`、`:stages/:tags → text[]`、`:is_admin → boolean`。语义路/视觉路的 NULL 关闭在 SQL 层表达；实现上按"是否提供向量"组装两条固定查询（hybrid / 纯 kw），避免 NULL 参数导致计划不稳定。

### 9.1 页面检索（可执行）

```sql
-- :qt 检索词, :q_text_emb, :q_img_emb(可空), :user_dept, :user_clearance, :is_admin,
-- :slide_type(可空), :tags(可空 text[])
with base as (                       -- 硬过滤 + ACL（所有召回路公共底座）
  select s.id, s.title, s.thumb_key, s.slide_type
  from slide s
  join source_document d
    on d.id = s.document_id
   and s.run_id = d.active_run_id    -- ★ 仅当前生效 run 的数据
  where s.review_state = 'approved'
    and s.usage_policy <> 'forbidden'
    and d.format_status = 'ok'       -- 隔离/损坏文档不进检索
    and d.review_state <> 'rejected'
    and coalesce(s.classification, d.classification) <= :user_clearance
    and ( d.owner_department_id = :user_dept
          or :is_admin
          or exists (select 1 from asset_acl a
                     where a.asset_type = 'document' and a.asset_id = d.id
                       and a.department_id = :user_dept and a.permission = 'read'))
    and not exists (select 1 from asset_acl a
                    where a.asset_type = 'document' and a.asset_id = d.id
                      and a.department_id = :user_dept and a.permission = 'deny')
    and not exists (select 1 from asset_acl a
                    where a.asset_type = 'slide' and a.asset_id = s.id
                      and a.department_id = :user_dept and a.permission = 'deny')
    and (:slide_type is null or s.slide_type = cast(:slide_type as text))
    and (:tags is null or s.tags @> cast(:tags as text[]))
),
sem as (                             -- 语义路（查询向量为空 → 整路为空）
  select b.id, row_number() over (order by s.text_emb <=> :q_text_emb) rn
  from base b join slide s on s.id = b.id
  where s.text_emb is not null
    and cast(:q_text_emb as vector) is not null
  order by s.text_emb <=> :q_text_emb
  limit 100
),
kw as (                              -- 关键词路（rn 升序 = 相似度降序）
  select b.id, row_number() over (
           order by similarity(coalesce(s.title,'') || ' ' || s.search_text, :qt) desc) rn
  from base b join slide s on s.id = b.id
  where similarity(coalesce(s.title,'') || ' ' || s.search_text, :qt) > 0.1
  order by rn
  limit 100
),
vis as (                             -- 视觉路（以图搜图时提供 :q_img_emb）
  select b.id, row_number() over (order by s.img_emb <=> :q_img_emb) rn
  from base b join slide s on s.id = b.id
  where s.img_emb is not null
    and cast(:q_img_emb as vector) is not null
  order by s.img_emb <=> :q_img_emb
  limit 100
)
select b.id, b.title, b.thumb_key, b.slide_type,
       coalesce(1.0/(60+sem.rn), 0)
     + coalesce(1.0/(60+kw.rn), 0)
     + coalesce(1.0/(60+vis.rn), 0) * :vis_w   -- 初始 1.0；以图搜图场景上调（M3 评测定版）
     as rrf_score
from base b
left join sem on sem.id = b.id
left join kw  on kw.id  = b.id
left join vis on vis.id = b.id
where sem.id is not null or kw.id is not null or vis.id is not null
order by rrf_score desc
limit 20;
```

### 9.2 素材检索（可执行，v2.1 新增）

```sql
-- :qt, :q_img_emb(可空), :user_dept, :user_clearance, :is_admin, :tags(可空)
with base as (
  select dm.id as media_id, mb.id as blob_id, mb.preview_key,
         coalesce(dm.classification, d.classification) as eff_cls
  from document_media dm
  join media_blob mb on mb.id = dm.blob_id
  join source_document d on d.id = dm.document_id
  where dm.review_state = 'approved'
    and dm.usage_policy <> 'forbidden'
    and d.format_status = 'ok'
    and d.review_state <> 'rejected'
    and coalesce(dm.classification, d.classification) <= :user_clearance
    and ( d.owner_department_id = :user_dept
          or :is_admin
          or exists (select 1 from asset_acl a
                     where a.asset_type = 'document' and a.asset_id = d.id
                       and a.department_id = :user_dept and a.permission = 'read'))
    and not exists (select 1 from asset_acl a
                    where a.asset_type = 'media' and a.asset_id = dm.id
                      and a.department_id = :user_dept and a.permission = 'deny')
    and not exists (select 1 from asset_acl a
                    where a.asset_type = 'document' and a.asset_id = d.id
                      and a.department_id = :user_dept and a.permission = 'deny')
    and (:tags is null or dm.tags @> cast(:tags as text[]))
),
kw as (
  select b.media_id, row_number() over (
           order by similarity(coalesce(mb.ocr_text,'') || ' ' || coalesce(mb.caption,''), :qt) desc) rn
  from base b join media_blob mb on mb.id = b.blob_id
  where similarity(coalesce(mb.ocr_text,'') || ' ' || coalesce(mb.caption,''), :qt) > 0.1
  order by rn
  limit 100
),
vis as (
  select b.media_id, row_number() over (order by mb.img_emb <=> :q_img_emb) rn
  from base b join media_blob mb on mb.id = b.blob_id
  where mb.img_emb is not null
    and cast(:q_img_emb as vector) is not null
  order by mb.img_emb <=> :q_img_emb
  limit 100
)
select b.media_id, b.blob_id, b.preview_key,
       coalesce(1.0/(60+kw.rn), 0) + coalesce(1.0/(60+vis.rn), 0) as rrf_score
from base b
left join kw  on kw.media_id  = b.media_id
left join vis on vis.media_id = b.media_id
where kw.media_id is not null or vis.media_id is not null
order by rrf_score desc
limit 20;
```

- 结构检索：`struct` 特征（栏数/图表类型/形状数）作为**过滤条件**（查询归一化层解析"三栏"→ `column_count=3`），不参与打分。
- 查询归一化：内网 LLM（可选）产出 `slide_type + tags + 结构过滤 + 检索词`；无 LLM 时纯检索词 + 手选过滤，功能完整。

### 9.3 检索评测（M3 交付）

- 标注查询集 ≥50 条真实业务查询 + 人工标注相关页。
- 指标与初始门槛：`Recall@20 ≥ 0.8`、`nDCG@10 ≥ 0.7`、审核员 Top-5 接受率 ≥ 60%。
- RRF 权重（`vis_w` 等）与各路阈值以评测定版，写入版本化配置。

---

## 10. 审核控制台与 API

### 10.1 页面（Jinja2 + htmx）
- 文档列表：run/stage 状态、失败原因、reprocess 入口；血缘视图（同 blob 多 document）。
- 页面缩略图墙 → 单页详情：原图 + 对象树（组合层级）+ JSON 对照 + 媒体引用 + warnings。
- 页面操作：打标（词表）、`review_state` 流转、`usage_policy`、`slide_type`、`quality`、`generation_allowed`、模板候选（框选区域 → layout + capacity）。
- 对象晋升组件 → `asset_component`（kind/标签/preview）。
- 素材操作：按 `media_blob.dedup_group` 浏览近似组 → 合并 document_media 标签/结论、密级、版权、`generation_allowed`。
- 规范确认：design-system 草稿（confidence/evidence）逐条确认 → 版本发布。
- admin：ACL 管理、审计查询、GC/对账任务状态。

### 10.2 API（REST，认证 + 审计）
```text
POST   /api/documents                      上传（multipart，必填部门/类别/密级/用途）
POST   /api/documents/watch-scan           触发受控目录扫描（admin）
GET    /api/documents?category=&state=
POST   /api/documents/{id}/reprocess       重跑（新 run + active 切换；from 参数已废弃）
POST   /api/documents/{id}/cancel          取消 run（接管租约 + 级联 cancelled）
GET    /api/search/slides?q=&type=&tags=&struct=
GET    /api/search/media?q=&tags=&img=     素材检索/以图搜图
GET    /api/slides/{id}                     详情（对象树/媒体引用；备注按 ACL）
PATCH  /api/slides/{id}                     审核（review_state/usage_policy/type/quality/generation_allowed）
POST   /api/slides/{id}/promote-template    模板候选（family_key+layout+capacity）
POST   /api/objects/{id}/promote-component  对象晋升组件
GET    /api/media?group=&tags=
POST   /api/media/{id}/review               素材审核（密级/用途/版权/合并结论）
GET    /api/templates?family=&status=
GET    /api/design-system
PUT    /api/design-system/{version}/confirm
GET    /api/reports/inventory
GET    /api/health/pipeline                 队列积压/失败率/租约接管/对账状态
```

---

## 11. 存储布局

```text
storage/                          # NAS / 本地卷 / MinIO（同一 Storage 抽象）
├── raw/          {blob_sha256}.pptx          # 原件：应用写入后置只读；生产用 NAS 只读子目录挂载或 MinIO 对象锁
├── quarantine/   宏/OLE 隔离（AV 后处置）     # 独立只读挂载
├── media/        原格式媒体（svg/emf/wmf 不转码）
├── previews/     媒体 PNG 预览
├── slides/       每页 PNG（≥1600px 宽）
├── thumbnails/   缩略图（480px 宽）
├── json/         文件/页/对象级 JSON 产物
├── reports/      盘点报告
└── tmp/          Saga 中转（对账可回收）
# 解包工作区在 worker 本地 tmpfs（read_only 容器的 /tmp），不入存储层
```

---

## 12. 部署与运维

### 12.1 Compose 基线（可启动，安全控制落地）

```yaml
name: ppt-pipeline

x-service-base: &service-base
  build: .
  environment: &service-env
    DATABASE_URL: postgresql+psycopg://ppt_app:${DB_PASSWORD}@db:5432/ppt_assets
    # DB_PASSWORD 与 db 服务的 secrets/db_password.txt 必须同源：
    # 部署基线统一由 .env 注入（DB_PASSWORD=...），entrypoint.sh 生成
    # secrets/db_password.txt（单配置源），避免两处手工维护
    STORAGE_ROOT: /storage
    PIPELINE_VERSION: "1.0.0"
  user: "10001:10001"                    # 非 root（镜像内建同 uid 用户）
  read_only: true
  tmpfs: ["/tmp:size=2g,mode=1777"]
  security_opt: ["no-new-privileges:true"]
  cap_drop: [ALL]
  networks: [backend]                    # internal 网络：无外网出口
  depends_on:
    db:
      condition: service_healthy
  restart: unless-stopped

services:
  db:
    image: pgvector/pgvector:pg16
    environment:
      POSTGRES_USER: ppt_app
      POSTGRES_PASSWORD_FILE: /run/secrets/db_password
      POSTGRES_DB: ppt_assets
    secrets: [db_password]
    volumes:
      - pgdata:/var/lib/postgresql/data
    networks: [backend]
    healthcheck:
      test: ["CMD-SHELL", "pg_isready -U ppt_app -d ppt_assets"]
      interval: 5s
      timeout: 3s
      retries: 12
    restart: unless-stopped

  app:
    <<: *service-base
    command: sh -c "alembic upgrade head && uvicorn ppt_pipeline.api:app --host 0.0.0.0 --port 8000"
    volumes:
      - storage:/storage
    ports:
      - "8000:8000"
    networks: [backend, frontend]        # 仅 app 可被宿主访问；app 自身代码无外呼
    deploy:
      resources:
        limits: { cpus: "2", memory: 2g }

  worker-extract:
    <<: *service-base
    command: python -m ppt_pipeline.worker --stages s0,s1,s2
    volumes:
      - storage:/storage
    deploy:
      resources:
        limits: { cpus: "2", memory: 2g }

  worker-render:
    <<: *service-base
    command: python -m ppt_pipeline.worker --stages s3
    volumes:
      - storage:/storage
    shm_size: 1g
    tmpfs: ["/tmp:size=6g,mode=1777"]    # LO profile + 渲染输出
    deploy:
      resources:
        limits: { cpus: "4", memory: 6g }

  worker-index:
    <<: *service-base
    command: python -m ppt_pipeline.worker --stages s4,s5
    volumes:
      - storage:/storage
    deploy:
      resources:
        limits: { cpus: "4", memory: 8g }   # 模型推理

networks:
  backend:
    internal: true                       # ★ 解析/渲染/DB 全部无外网出口
  frontend: {}                           # 仅承载 app 发布端口

volumes:
  pgdata: {}
  storage: {}

secrets:
  db_password:
    file: ./secrets/db_password.txt
```

说明：模型权重与字体 `COPY` 进镜像（`models/`、`fonts/`，见 Dockerfile），不使用空挂载卷；`DB_PASSWORD` 经 `.env` 注入（git 忽略）；`storage` 为命名卷，本地基线中原件只读由应用层保证，**生产替换为 NAS 只读子目录挂载 `raw/`/`quarantine/` 或 MinIO 对象锁**（文件系统级强制）；迁移由 app 入口执行（单实例部署下无并发迁移竞争，多实例需迁移锁或独立 job）。

### 12.2 运维基线
- 备份：PG 每日全量 + WAL 归档；存储按 NAS 快照/MinIO 版本化；权重与配置入制品库。
- 迁移：Alembic 必须含回滚；升级前存储与 DB 双快照。
- 指标（`/api/health/pipeline` + Prometheus）：队列积压、各阶段耗时分布、attempt/失败率、**租约接管次数**、23505 冲突率、LO 超时率、磁盘、对账孤儿数。
- 日志：结构化 JSON，`trace_id/document_id/run_id/stage_run_id` 全链路。
- 版本清单（`models/MANIFEST.json` + 镜像标签）：python-pptx/LO/字体/BGE/CLIP/OCR/pipeline_version/config_hash；渲染相关变更触发黄金样本重跑。
- 容量：按页均估（PNG ≈0.5–2MB、缩略图 ≈50KB、向量 ≈6KB/页、媒体原样），试点后按实测修订。

### 12.3 性能认知
瓶颈需 M0/M2 实测：对象级批量写库（单页可达数千对象 → batch insert + 单事务页级提交）、OOXML 大包、CPU 推理吞吐、OCR、存储 IO、渲染。万页试点下渲染仍是大头，index 角色单独容量规划。

### 12.4 规模化路径（触发条件 → 动作）
| 触发 | 动作 |
| --- | --- |
| 页数 > 5 万或检索 P95 > 500ms | 启用 HNSW |
| 队列竞争/吞吐不足（23505 率高） | 引入 Celery+Redis（领取/租约语义不变） |
| 多部门共享 | NAS → MinIO（对象锁/版本化/IAM），过法务 |
| CLIP 版式理解不达标 | 视觉路降级仅以图搜图，或内网微调模型 |
| 黄金样本差异超标 | Parser/Renderer 接口接 Aspose 替换 S2/S3 |

---

## 13. 里程碑

**假设：2 后端 + 1 前端（兼）+ 业务审核兼职（M0/M6）。** 法务确认为**外部依赖**：第 0 天发起，不占研发工期，但其结论是 **M1 准入门槛**（Go/No-Go 决策点）。

| 阶段 | 内容 | 关键交付 | 工期 |
| --- | --- | --- | --- |
| M0 技术验证 | 黄金样本集 + SSIM 基线（阈值冻结）、pypdfium2/PyMuPDF 对比、BGE/CLIP/OCR 评测、解析吞吐实测；**并行发起法务** | 基线报告 + 冻结选型清单 + 法务跟踪单 | 1–2 周 |
| M1 数据与流水线底座 | §5 全量迁移、bootstrap、领取/租约/fencing、Saga 发布、S0–S2、媒体两层 | 三级 JSON 入库；并发/fencing 集成测试绿 | 3 周 |
| M2 渲染与安全 | S3（隔离/超时/自愈）、字体镜像、Compose 安全基线（§12.1 全项） | 页面 PNG/缩略图稳定产出；安全用例通过 | 2–3 周 |
| M3 检索基线 | S4、素材检索、评测集与调参 | Recall@20 ≥ 0.8 / nDCG@10 ≥ 0.7 | 2–3 周 |
| M4 审核台 | 文档/页面/素材/组件审核、模板候选、规范确认、ACL+审计 | 业务可日常打标 | 3–4 周（与 M3 部分并行） |
| M5 稳定性与报告 | S5、reprocess/GC/对账验证、监控、备份恢复演练、压测 | 试运行就绪 | 2 周 |
| M6 业务试点 | 1 素材 + 1 规范 + 5–10 范例全流程；首批资产（30–50 素材、8–12 页面类型、10–20 模板候选） | 首版"设计资产底座" | 2–3 周 |

合计约 **13–17 周**（并行后）；WPS/旧格式/SmartArt/OLE 占比高 → 追加 3–6 周兼容治理；1 名全栈 → 16–24 周。

---

## 14. 风险与对策

| 风险 | 对策 |
| --- | --- |
| LO 渲染保真度 | 黄金样本 SSIM 量化（阈值 M0 冻结）；预览标注"仅供参考"；原始 XML/原件永久保留；超标 → Aspose 接口替换 |
| 企业字体授权限制打包 | M0 法务；未授权不进镜像，替代映射 + 预览标注差异 |
| AGPL 组件 | 首选非 AGPL 替代（pypdfium2/NAS）；确需使用过法务 |
| python-pptx 覆盖不全 | OOXML 层补充；SmartArt/动画只检测；页级异常不阻塞 |
| WPS/旧格式 | 魔数分流 + OLE 宏检测 + 转换日志 + 复核清单；黄金样本含 WPS |
| 中文检索质量 | 向量为主 + 受控标签 + trigram；评测驱动；人工标签质量是第一保障 |
| 恶意文档 | 双通道宏检测 + 隔离区 + AV + internal 网络 + 非 root/只读根/限额 |
| DB/存储不一致 | Saga 发布 + 对账任务 + artifact 状态机 |
| 并发异常（互斥/双写） | 唯一部分索引互斥 + lease fencing CAS + 集成测试（kill -9、并发领取、租约过期） |
| 解析器/模型升级混用 | run 版本化 + active_run_id 切换 + GC |
| 审核瓶颈 | 审核台效率优先（快捷键/批量/近似组聚合）；M6 真实审核员验证吞吐 |

---

## 15. 验收标准（量化 + 口径定义）

| # | 需求（原文档功能要求） | 指标与口径 | 里程碑 |
| --- | --- | --- | --- |
| 1 | 文件/页面/对象级 JSON | 进入流水线的 run 解析成功率 ≥95%（= succeeded run / (succeeded+failed)，不含 cancelled/superseded；拒收的加密/损坏文件单列"导入拒收率"）；失败均带 error_code 进复核清单 | M1 |
| 2 | 页面全要素提取 | 黄金样本全页人工比对固定字段清单（页码/尺寸/版式/文本/字体/字号/对齐/填充/线条/坐标/层级/表格/图表/备注/隐藏页/超链接），完整率 ≥98%（缺失字段数/应填字段总数） | M1–M2 |
| 3 | 原始媒体导出保留格式 | media 原样导出率 100%（svg/emf/wmf 不转码，抽检 SHA 一致） | M1 |
| 4 | 主题/母版/页脚/Logo 解析 | 规范草稿每条带 confidence/evidence/extractor_version；确认走完整版本流 | M3/M4 |
| 5 | PNG/缩略图/视觉向量 | 渲染成功率 ≥99%（成功页/应渲染页；转换级失败重试用尽计入失败）；SSIM ≥ M0 冻结阈值；向量覆盖率 ≥98%（有文本页 text_emb 非 NULL 比例） | M2–M3 |
| 6 | 去重 | 精确去重 100%（sha256 全局唯一由约束保证）；近似组（照片阈值≤8/图标≤4）人工合并误合并率 0 | M3 |
| 7 | 文件/页面/素材打标与可用范围 | 文件级审核字段全链路生效；密级/用途/版权受 ACL 过滤；标签仅可取词表 key（触发器强制） | M4 |
| 8 | 优秀页面标注（类型/模板/容量/组件/允许生成） | 组件晋升与 generation_allowed 落库可查；模板同 family 仅一个 active（约束保证） | M4 |
| 9 | 盘点报告 | 报告自动生成；出现次数与复用动作两口径分列 | M5 |
| 10 | 检索质量 | Recall@20 ≥ 0.8、nDCG@10 ≥ 0.7（≥50 条标注查询集） | M3 |
| 11 | 幂等/重跑/并发 | 集成测试：并发领取同文件无并行 stage（23505 退避路径覆盖）；kill -9 后租约接管且无重复产物；reprocess 后检索/UI 仅见新 run；GC 后无悬挂引用；对账无孤儿 | M5 |
| 12 | 安全 | 宏（含旧 .ppt VBA）/加密 100% 分流（按黄金样本恶意用例集）；ACL 渗透用例 0 越权（deny 生效/跨部门 read 生效/密级继承）；容器网络无外网出口（出网探测用例） | M2/M5 |
