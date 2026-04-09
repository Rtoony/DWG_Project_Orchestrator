import json
import asyncio
import os
from sqlmodel import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.database import engine
from app.models.layer_standard import LayerStandard
from app.models.filename_rule import FilenameRule
from app.models.recipe import Recipe, RecipeCategory
from app.models.project_preset import ProjectPreset
from app.models.viewport_preset import ViewportPreset

SEED_DATA_DIR = "/seed_data"

async def seed_layer_standards(session: AsyncSession):
    file_path = os.path.join(SEED_DATA_DIR, "layer_standards.json")
    if not os.path.exists(file_path):
        print(f"Skipping layer standards: {file_path} not found")
        return
    
    with open(file_path, 'r') as f:
        data = json.load(f)
        # Assuming data is a list of dicts
        for item in data:
            # Check if exists
            result = await session.execute(select(LayerStandard).where(LayerStandard.name == item['name']))
            if not result.scalars().first():
                session.add(LayerStandard(**item))
    print("Seeded layer standards")

async def seed_filename_rules(session: AsyncSession):
    file_path = os.path.join(SEED_DATA_DIR, "dwg_filename_rules.json")
    if not os.path.exists(file_path):
        print(f"Skipping filename rules: {file_path} not found")
        return
    
    with open(file_path, 'r') as f:
        data = json.load(f)
        for item in data:
            result = await session.execute(select(FilenameRule).where(FilenameRule.file_type_code == item['file_type_code']))
            if not result.scalars().first():
                session.add(FilenameRule(**item))
    print("Seeded filename rules")

async def seed_recipes(session: AsyncSession):
    file_path = os.path.join(SEED_DATA_DIR, "automation_recipes.json")
    if not os.path.exists(file_path):
        print(f"Skipping recipes: {file_path} not found")
        return
    
    with open(file_path, 'r') as f:
        data = json.load(f)
        # Assuming format { "categories": [...], "recipes": [...] }
        for cat_data in data.get('categories', []):
            result = await session.execute(select(RecipeCategory).where(RecipeCategory.name == cat_data['name']))
            if not result.scalars().first():
                session.add(RecipeCategory(**cat_data))
        
        await session.flush() # Ensure categories are saved to get IDs
        
        for rec_data in data.get('recipes', []):
            result = await session.execute(select(Recipe).where(Recipe.name == rec_data['name']))
            if not result.scalars().first():
                # Map category name to ID if needed
                session.add(Recipe(**rec_data))
    print("Seeded recipes and categories")

async def main():
    async with AsyncSession(engine) as session:
        await seed_layer_standards(session)
        await seed_filename_rules(session)
        await seed_recipes(session)
        await session.commit()

if __name__ == "__main__":
    asyncio.run(main())
