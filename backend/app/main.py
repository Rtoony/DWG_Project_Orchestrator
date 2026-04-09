from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.config import settings
from app.routers import projects, standards, recipes, jobs
from contextlib import asynccontextmanager
from app.database import engine
from sqlmodel import SQLModel

# Import all models to ensure they are registered with SQLModel metadata
from app.models import project, drawing, layer_standard, filename_rule, recipe, viewport_preset, project_preset, analysis, job, audit_log

@asynccontextmanager
async def lifespan(app: FastAPI):
    async with engine.begin() as conn:
        # Create all tables
        await conn.run_sync(SQLModel.metadata.create_all)
    yield

app = FastAPI(
    title=settings.PROJECT_NAME,
    openapi_url=f"{settings.API_V1_STR}/openapi.json",
    lifespan=lifespan
)

# Set all CORS enabled origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(projects.router, prefix=f"{settings.API_V1_STR}/projects", tags=["projects"])
app.include_router(standards.router, prefix=f"{settings.API_V1_STR}/standards", tags=["standards"])
app.include_router(recipes.router, prefix=f"{settings.API_V1_STR}/recipes", tags=["recipes"])
app.include_router(jobs.router, prefix=f"{settings.API_V1_STR}/jobs", tags=["jobs"])

@app.get("/")
def root():
    return {"message": "Welcome to DWG Project Orchestrator API"}

@app.get("/health")
def health_check():
    return {"status": "ok"}
