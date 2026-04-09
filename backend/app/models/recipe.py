from datetime import datetime
from typing import Optional, List
from sqlmodel import SQLModel, Field, Relationship

class RecipeCategoryBase(SQLModel):
    name: str = Field(unique=True, index=True)
    description: Optional[str] = None
    sort_order: int = Field(default=0)

class RecipeCategory(RecipeCategoryBase, table=True):
    __tablename__ = "recipe_categories"
    id: Optional[int] = Field(default=None, primary_key=True)
    recipes: List["Recipe"] = Relationship(back_populates="category")

class RecipeBase(SQLModel):
    category_id: Optional[int] = Field(default=None, foreign_key="recipe_categories.id")
    name: str
    runner: str  # core_console, pyautocad, python_direct
    script_file: Optional[str] = None
    command: Optional[str] = None
    description: Optional[str] = None
    script_content: Optional[str] = None
    is_active: bool = Field(default=True)

class Recipe(RecipeBase, table=True):
    __tablename__ = "recipes"
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    category: Optional[RecipeCategory] = Relationship(back_populates="recipes")

class RecipeCreate(RecipeBase):
    pass

class RecipeRead(RecipeBase):
    id: int
    created_at: datetime
