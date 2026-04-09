# DWG Project Orchestrator v2 — Architecture & Implementation Plan

## What We're Building

A **web-based CAD project management platform** that replaces the PyQt6 desktop prototype. It runs on your Proxmox infrastructure, accessible from any machine at B&R (or home), and separates concerns cleanly:

- **Web UI** → React dashboard (runs anywhere with a browser)
- **API Server** → FastAPI backend (runs on Proxmox VM)
- **Database** → PostgreSQL (your existing Proxmox DB infrastructure)
- **CAD Worker** → Lightweight agent on Windows machines with AutoCAD (handles COM automation + accoreconsole)
- **AI Layer** → OpenClaw integration for intelligent auditing (future phase)

---

## Why This Stack

| Choice | Why |
|--------|-----|
| **FastAPI** | You know Python, it's async, auto-generates OpenAPI docs, trivial to add endpoints |
| **PostgreSQL** | You already run it for ACAD-GIS, you know it cold, PostGIS for future spatial queries |
| **SQLModel** | Pydantic + SQLAlchemy in one — models are both DB schemas AND API request/response types |
| **React + Tailwind** | Claude Code / Gemini CLI eat this for breakfast. Vibe-code the entire frontend |
| **Redis (optional)** | Job queue for async CAD worker tasks, real-time progress via WebSockets |
| **Docker** | Everything containerized on Proxmox. One `docker compose up` and it's running |

---

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                        BROWSER (Any Machine)                │
│  ┌───────────────────────────────────────────────────────┐  │
│  │              React + Tailwind Dashboard                │  │
│  │  ┌─────────┐ ┌──────────┐ ┌─────────┐ ┌───────────┐  │  │
│  │  │Projects │ │Automation│ │Standards│ │ Analysis  │  │  │
│  │  │Manager  │ │  Hub     │ │Dashboard│ │ Viewer    │  │  │
│  │  └─────────┘ └──────────┘ └─────────┘ └───────────┘  │  │
│  └───────────────────────┬───────────────────────────────┘  │
└──────────────────────────┼──────────────────────────────────┘
                           │ HTTP/WebSocket
┌──────────────────────────┼──────────────────────────────────┐
│                    PROXMOX VM (Linux)                        │
│  ┌───────────────────────┴───────────────────────────────┐  │
│  │                 FastAPI Backend                         │  │
│  │  /api/projects    /api/standards   /api/recipes         │  │
│  │  /api/analysis    /api/jobs        /api/health          │  │
│  │  /ws/jobs/{id}    (WebSocket for live progress)         │  │
│  └──────────┬────────────────────────┬───────────────────┘  │
│             │                        │                       │
│  ┌──────────┴──────────┐  ┌─────────┴─────────┐            │
│  │    PostgreSQL       │  │   Redis (Queue)    │            │
│  │  ┌──────────────┐   │  │  Job dispatch +    │            │
│  │  │ projects     │   │  │  progress tracking │            │
│  │  │ drawings     │   │  └───────────┬────────┘            │
│  │  │ layer_stds   │   │              │                     │
│  │  │ recipes      │   │              │                     │
│  │  │ filename_rules│  │              │                     │
│  │  │ analyses     │   │              │                     │
│  │  │ audit_log    │   │              │                     │
│  │  └──────────────┘   │              │                     │
│  └─────────────────────┘              │                     │
└───────────────────────────────────────┼─────────────────────┘
                                        │ Job pickup (HTTP poll or Redis sub)
┌───────────────────────────────────────┼─────────────────────┐
│              WINDOWS WORKSTATION (AutoCAD)                    │
│  ┌────────────────────────────────────┴──────────────────┐  │
│  │              CAD Worker Agent (Python)                  │  │
│  │  • Polls API for pending jobs                          │  │
│  │  • Executes AutoCAD COM automation                     │  │
│  │  • Runs accoreconsole for headless ops                 │  │
│  │  • Uploads results (DXF, metadata) back to API         │  │
│  │  • Reports progress via WebSocket                      │  │
│  └────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────┘
```

---

## Database Schema (Phase 1 — Kill the JSON Files)

This maps every JSON config file to proper relational tables.

### Core Tables

```sql
-- ============================================================
-- PROJECTS & DRAWINGS
-- ============================================================

CREATE TABLE projects (
    id              SERIAL PRIMARY KEY,
    project_number  VARCHAR(20) NOT NULL,
    sub_number      VARCHAR(10) NOT NULL,
    project_name    TEXT,
    client_name     TEXT,
    project_manager TEXT,
    lead_designer   TEXT,
    project_date    DATE,
    project_status  VARCHAR(10) DEFAULT 'SD',  -- SD, DD, CD
    setup_config    VARCHAR(50),               -- School_Small, BR_Plan, etc.
    tb_size         VARCHAR(10),               -- 11x17, 22x34, 24x36, 30x42
    tb_type         VARCHAR(10),               -- BR, EXHIBIT, DSA, QKA, SR
    coordinate_system TEXT,
    vertical_datum  TEXT,
    root_path       TEXT NOT NULL DEFAULT 'J:\J',
    archive_path    TEXT DEFAULT 'R:\J',
    created_at      TIMESTAMPTZ DEFAULT NOW(),
    updated_at      TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(project_number, sub_number)
);

CREATE TABLE drawings (
    id              SERIAL PRIMARY KEY,
    project_id      INTEGER REFERENCES projects(id) ON DELETE CASCADE,
    filename        TEXT NOT NULL,
    file_type_code  VARCHAR(20) NOT NULL,      -- DESIGN, PLAN, C-TOPO, etc.
    description     TEXT,
    phase           VARCHAR(20),
    folder_path     TEXT,
    file_size_bytes BIGINT,
    dwg_version     VARCHAR(20),
    last_modified   TIMESTAMPTZ,
    last_analyzed   TIMESTAMPTZ,
    status          VARCHAR(20) DEFAULT 'active', -- active, archived, deleted
    created_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_drawings_project ON drawings(project_id);
CREATE INDEX idx_drawings_file_type ON drawings(file_type_code);

-- ============================================================
-- LAYER STANDARDS (replaces layer_standards.json — 361 rows)
-- ============================================================

CREATE TABLE layer_standards (
    id                    SERIAL PRIMARY KEY,
    name                  VARCHAR(100) NOT NULL UNIQUE,
    color_code            INTEGER NOT NULL,
    linetype              VARCHAR(50) DEFAULT 'CONTINUOUS',
    lineweight            INTEGER DEFAULT -3,
    is_plottable          BOOLEAN DEFAULT TRUE,
    plot_style_name       VARCHAR(50),
    category              VARCHAR(50),          -- Sheet Layout, Utility, Survey, etc.
    discipline            VARCHAR(50),          -- Civil, Structural, etc.
    status                VARCHAR(20),          -- General, Required, Optional
    description           TEXT,
    typical_object_types  TEXT[],               -- PostgreSQL array
    notes                 TEXT,
    standards_revision_id INTEGER DEFAULT 1,
    created_at            TIMESTAMPTZ DEFAULT NOW(),
    updated_at            TIMESTAMPTZ DEFAULT NOW()
);

-- ============================================================
-- FILENAME RULES (replaces dwg_filename_rules.json — 38 rows)
-- ============================================================

CREATE TABLE filename_rules (
    id                          SERIAL PRIMARY KEY,
    file_type_code              VARCHAR(20) NOT NULL UNIQUE,
    file_type_description       TEXT,
    folder_path_template        TEXT,           -- J:\J\[ProjectNumber]\dwg\...
    filename_pattern            TEXT,           -- [ProjectNumber].[Subnumber] ...
    phase_required              BOOLEAN DEFAULT FALSE,
    phase_source                VARCHAR(50),
    phase_allowed_list_source   VARCHAR(50),
    phase_format                VARCHAR(100),
    description_required        BOOLEAN DEFAULT FALSE,
    description_format          VARCHAR(100),
    multi_instance_allowed      BOOLEAN DEFAULT TRUE,
    notes                       TEXT,
    created_at                  TIMESTAMPTZ DEFAULT NOW()
);

-- ============================================================
-- AUTOMATION RECIPES (replaces automation_recipes.json)
-- ============================================================

CREATE TABLE recipe_categories (
    id          SERIAL PRIMARY KEY,
    name        VARCHAR(100) NOT NULL UNIQUE,
    description TEXT,
    sort_order  INTEGER DEFAULT 0
);

CREATE TABLE recipes (
    id              SERIAL PRIMARY KEY,
    category_id     INTEGER REFERENCES recipe_categories(id),
    name            VARCHAR(200) NOT NULL,
    runner          VARCHAR(20) NOT NULL,       -- core_console, pyautocad, python_direct
    script_file     TEXT,
    command         VARCHAR(200),
    description     TEXT,
    script_content  TEXT,                       -- optional: store script inline
    is_active       BOOLEAN DEFAULT TRUE,
    created_at      TIMESTAMPTZ DEFAULT NOW()
);

-- ============================================================
-- VIEWPORT PRESETS (replaces viewport_presets.json)
-- ============================================================

CREATE TABLE viewport_presets (
    id          SERIAL PRIMARY KEY,
    tb_type     VARCHAR(10) NOT NULL,           -- QKA, BR, DSA, etc.
    tb_size     VARCHAR(10) NOT NULL,           -- 24x36, 22x34, etc.
    layout_code VARCHAR(20) NOT NULL,           -- COVER, PLAN, DEMO, etc.
    viewports   JSONB NOT NULL,                 -- Array of viewport configs
    UNIQUE(tb_type, tb_size, layout_code)
);

-- ============================================================
-- PROJECT PRESETS (replaces project_presets.json)
-- ============================================================

CREATE TABLE project_presets (
    id          SERIAL PRIMARY KEY,
    name        VARCHAR(100) NOT NULL UNIQUE,   -- School_Small, BR_Plan, etc.
    description TEXT,
    drawings    JSONB NOT NULL                  -- [{code, description}, ...]
);

-- ============================================================
-- DXF ANALYSIS RESULTS (new — stores what dxf_analyzer.py produces)
-- ============================================================

CREATE TABLE dxf_analyses (
    id              SERIAL PRIMARY KEY,
    drawing_id      INTEGER REFERENCES drawings(id) ON DELETE CASCADE,
    analysis_data   JSONB NOT NULL,             -- Full analysis JSON
    entity_count    INTEGER,
    layer_count     INTEGER,
    block_count     INTEGER,
    analyzed_at     TIMESTAMPTZ DEFAULT NOW()
);

-- ============================================================
-- AUDIT LOG (new — tracks all operations for health checks)
-- ============================================================

CREATE TABLE audit_log (
    id          SERIAL PRIMARY KEY,
    project_id  INTEGER REFERENCES projects(id),
    action      VARCHAR(50) NOT NULL,           -- create_drawing, run_recipe, etc.
    details     JSONB,
    user_name   VARCHAR(100),
    created_at  TIMESTAMPTZ DEFAULT NOW()
);

-- ============================================================
-- JOB QUEUE (for CAD worker tasks)
-- ============================================================

CREATE TABLE jobs (
    id          SERIAL PRIMARY KEY,
    project_id  INTEGER REFERENCES projects(id),
    job_type    VARCHAR(50) NOT NULL,           -- recipe_run, dwg_convert, batch_plot
    status      VARCHAR(20) DEFAULT 'pending',  -- pending, running, completed, failed
    payload     JSONB NOT NULL,                 -- Job-specific parameters
    result      JSONB,                          -- Output/errors
    progress    INTEGER DEFAULT 0,              -- 0-100
    worker_id   VARCHAR(100),                   -- Which CAD worker picked it up
    created_at  TIMESTAMPTZ DEFAULT NOW(),
    started_at  TIMESTAMPTZ,
    completed_at TIMESTAMPTZ
);

CREATE INDEX idx_jobs_status ON jobs(status);
CREATE INDEX idx_jobs_project ON jobs(project_id);
```

### Data Migration Script Concept

```python
# migrate_json_to_pg.py — Run once to seed the database from your JSON files
# Claude Code can generate this completely from the schema above
```

---

## API Endpoints (Phase 2)

```
# Projects
GET    /api/projects                    List all projects
POST   /api/projects                    Create project (+ folder structure)
GET    /api/projects/{id}               Get project details
PUT    /api/projects/{id}               Update project metadata
GET    /api/projects/{id}/drawings      List drawings in project
POST   /api/projects/{id}/drawings      Register a new drawing

# Standards
GET    /api/standards/layers            List all layer standards
GET    /api/standards/layers?category=X Filter by category
PUT    /api/standards/layers/{id}       Update a layer standard
GET    /api/standards/filename-rules    List filename rules
GET    /api/standards/filename-rules/{code}  Get specific rule

# Recipes & Automation
GET    /api/recipes                     List all recipes (with categories)
GET    /api/recipes/{id}                Get recipe details
POST   /api/jobs                        Submit a job (recipe execution)
GET    /api/jobs/{id}                   Get job status
GET    /api/jobs?project_id=X           List jobs for project
WS     /ws/jobs/{id}                    Live progress stream

# Analysis
POST   /api/analysis/dxf               Upload & analyze a DXF file
GET    /api/analysis/{id}               Get analysis results
GET    /api/analysis?drawing_id=X       Get analyses for a drawing

# Health Check
GET    /api/health/project/{id}         Run project health audit
GET    /api/health/drawing/{id}         Check single drawing compliance

# Presets
GET    /api/presets/projects            List project presets
GET    /api/presets/viewports           List viewport presets
```

---

## Project Structure

```
dwg-orchestrator/
├── docker-compose.yml              # PostgreSQL + API + Redis + frontend
├── backend/
│   ├── Dockerfile
│   ├── pyproject.toml
│   ├── alembic/                    # DB migrations
│   │   └── versions/
│   ├── app/
│   │   ├── main.py                 # FastAPI app entry
│   │   ├── config.py               # Settings (DB URL, paths, etc.)
│   │   ├── database.py             # SQLModel engine + session
│   │   ├── models/                 # SQLModel models (= DB + API schemas)
│   │   │   ├── project.py
│   │   │   ├── drawing.py
│   │   │   ├── layer_standard.py
│   │   │   ├── filename_rule.py
│   │   │   ├── recipe.py
│   │   │   ├── viewport_preset.py
│   │   │   ├── project_preset.py
│   │   │   ├── analysis.py
│   │   │   ├── job.py
│   │   │   └── audit_log.py
│   │   ├── routers/                # FastAPI routers (one per domain)
│   │   │   ├── projects.py
│   │   │   ├── standards.py
│   │   │   ├── recipes.py
│   │   │   ├── analysis.py
│   │   │   ├── jobs.py
│   │   │   └── health.py
│   │   ├── services/               # Business logic
│   │   │   ├── project_service.py
│   │   │   ├── analysis_service.py # DXF parsing (ezdxf)
│   │   │   ├── health_service.py   # Audit/compliance checks
│   │   │   └── job_service.py
│   │   └── migrations/             # JSON→DB seed scripts
│   │       └── seed_from_json.py
│   └── tests/
├── frontend/
│   ├── Dockerfile
│   ├── package.json
│   ├── src/
│   │   ├── App.jsx
│   │   ├── api/                    # API client (auto-generated from OpenAPI)
│   │   ├── components/
│   │   │   ├── ProjectManager/
│   │   │   ├── AutomationHub/
│   │   │   ├── StandardsDashboard/
│   │   │   ├── AnalysisViewer/
│   │   │   ├── LayerManager/
│   │   │   └── JobMonitor/
│   │   └── pages/
│   └── tailwind.config.js
├── worker/                         # Runs on Windows with AutoCAD
│   ├── worker_agent.py             # Polls API, executes jobs
│   ├── autocad_engine.py           # COM automation (from your existing code)
│   ├── accoreconsole_runner.py     # Headless operations
│   └── config.yaml                 # Worker config (API URL, AutoCAD paths)
└── seed_data/                      # Your existing JSON files for initial migration
    ├── layer_standards.json
    ├── dwg_filename_rules.json
    ├── automation_recipes.json
    ├── project_presets.json
    └── viewport_presets.json
```

---

## Implementation Phases

### Phase 1: Database + API Foundation (Week 1-2)
**Claude Code prompt target: backend/**
1. PostgreSQL schema creation via Alembic migrations
2. SQLModel models for all tables
3. Seed script to import all 6 JSON files
4. Basic CRUD endpoints for projects, standards, recipes
5. Docker Compose for local dev (PG + API)

### Phase 2: Web Dashboard (Week 2-3)
**Claude Code prompt target: frontend/**
1. React app with router + Tailwind
2. Projects list/create/detail pages
3. Layer Standards browser (searchable, filterable, editable)
4. Filename Rules viewer
5. Recipes browser with categories

### Phase 3: DXF Analysis (Week 3-4)
**Claude Code prompt target: backend/app/services/analysis_service.py**
1. Port dxf_analyzer.py to a FastAPI service
2. Upload DXF → analyze → store results in PG
3. Analysis viewer in frontend (layer breakdown, entity stats, block inventory)
4. Health check endpoint (compare drawing layers against standards table)

### Phase 4: CAD Worker Agent (Week 4-5)
**Claude Code prompt target: worker/**
1. Worker agent that polls /api/jobs for pending tasks
2. Port AutomationEngine class to worker (COM, accoreconsole)
3. WebSocket progress reporting
4. Job monitor page in frontend

### Phase 5: OpenClaw Integration (Future)
1. AI-powered drawing audit (send analysis JSON to LLM)
2. Natural language project setup via Telegram
3. Anomaly detection across project drawings

---

## What Gets Preserved vs. Rewritten

| Original Component | What Happens |
|---|---|
| `dwg_project_orchestrator.py` (2654 lines) | **Decomposed** → models, routers, services |
| `config_manager.py` | **Replaced** by SQLModel + database.py |
| `dxf_analyzer.py` | **Preserved** mostly as-is inside analysis_service.py |
| `backup_json/*.json` | **Migrated** into PostgreSQL via seed script, then retired |
| `recipes/*.scr` + `*.lsp` | **Stored** in DB (script_content column) or kept as files in worker/ |
| `styles.qss` | **Gone** — Tailwind handles all styling |
| PyQt6 UI tabs | **Rewritten** as React components |
| AutoCAD COM engine | **Moved** to worker agent (Windows-only piece) |
| File path logic (J:\, R:\) | **Configurable** via project.root_path column |

---

## Key Design Decisions

**1. Why not just put scripts in the DB?**
You can do both. The `recipes` table has both `script_file` (path) and `script_content` (inline). Short scripts go inline, complex LISP files stay as files in the worker's filesystem.

**2. Why a separate worker agent?**
AutoCAD COM only works on Windows with AutoCAD installed. By making it a separate process that talks to your API, the web app stays platform-independent. You could even have multiple workers (one per AutoCAD seat) for parallel processing.

**3. Why JSONB for some columns?**
Viewport configs and analysis results are deeply nested and variable in structure. JSONB lets you store them as-is while still being queryable (`analysis_data->>'entity_count'`). The structured data (layers, rules, projects) gets proper columns.

**4. Why not Next.js / full-stack JS?**
Your strength is Python. FastAPI lets you write the backend in the language you know. The frontend is just a thin client — React + fetch calls. Claude Code generates both equally well.

---

## First Claude Code Prompt (Ready to Go)

When you're ready to start, create a new repo and give Claude Code this prompt:

```
Create the backend/ directory for a FastAPI application with:
- SQLModel models matching the schema in ARCHITECTURE_V2.md
- Alembic migration setup for PostgreSQL
- A seed script that reads JSON files from seed_data/ and populates all tables
- CRUD routers for: projects, layer_standards, filename_rules, recipes, project_presets, viewport_presets
- Docker Compose with PostgreSQL 16 and the FastAPI app
- Use async SQLAlchemy with asyncpg
- Include proper error handling and pagination on list endpoints
```
