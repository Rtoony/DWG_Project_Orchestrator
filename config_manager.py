# Configuration Manager for DWG Project Orchestrator
# This module centralizes all configuration loading from both JSON files and database
# while preserving 100% of the existing functionality and interface.

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple, Union

# Import database manager for database functionality
try:
    from database_manager import DatabaseManager
    DATABASE_AVAILABLE = True
except ImportError:
    DATABASE_AVAILABLE = False

# Re-define Rule dataclass to avoid circular imports
@dataclass
class Rule:
    code: str
    folder_short: str 
    Description_Required: bool = False
    Multi_Instance_Allowed: bool = False
    filename_pattern: str | None = None

def _last_segment(path: str) -> str:
    """Extract the last segment of a folder path for rule processing."""
    return Path(path).name if path else ""

class ConfigurationManager:
    """Centralized configuration management for the DWG Project Orchestrator.
    
    This class handles all JSON file loading with the exact same interface
    and error handling as the original implementation.
    """
    
    def __init__(self, app_dir: Path, use_database: bool = True):
        """Initialize the configuration manager with the application directory.
        
        Args:
            app_dir: Application directory containing JSON files
            use_database: Whether to try database first (falls back to JSON if unavailable)
        """
        self.app_dir = app_dir
        self.use_database = use_database and DATABASE_AVAILABLE
        
        # Backup JSON file paths (exact table name matching)
        self.recipes_config = app_dir / "backup_json" / "automation_recipes.json"
        self.presets_config = app_dir / "backup_json" / "project_presets.json" 
        self.templates_default = app_dir / "backup_json" / "templates.json"
        self.rules_default = app_dir / "backup_json" / "dwg_filename_rules.json"
        
        # Initialize database manager if available
        self.db_manager = None
        if self.use_database and DATABASE_AVAILABLE:
            try:
                self.db_manager = DatabaseManager(app_dir)
                if not self.db_manager.is_connected():
                    print("Database not connected, falling back to JSON files")
                    self.use_database = False
            except Exception as e:
                print(f"Database initialization failed, using JSON files: {e}")
                self.use_database = False
    
    def load_recipes(self) -> Tuple[Dict[str, Any], Dict[str, Any], Optional[str]]:
        """Load recipes from database or backup_json/automation_recipes.json with support for both flat and categorized formats.
        
        Returns:
            Tuple of (recipes_dict, recipes_categorized_dict, error_message)
            - recipes_dict: Flattened recipes for backward compatibility
            - recipes_categorized_dict: Categorized structure for UI
            - error_message: None if successful, error string if failed
        """
        # Try Supabase database first (automation_recipes + recipe_categories tables)
        db_error = None
        if self.use_database and self.db_manager and self.db_manager.is_connected():
            try:
                automation_recipes_flat, automation_recipes_categorized, error = self.db_manager.load_automation_recipes_from_recipe_categories_tables()
                if not error and automation_recipes_flat:
                    return automation_recipes_flat, automation_recipes_categorized, None
                # If Supabase tables fail, capture error for reporting
                if error:
                    db_error = f"Supabase automation_recipes/recipe_categories tables failed: {error}"
                    print(f"{db_error}, falling back to backup_json/automation_recipes.json")
            except Exception as e:
                db_error = f"Supabase database error: {e}"
                print(f"{db_error}, falling back to backup_json/automation_recipes.json")
        
        # Fallback to backup JSON file  
        if not self.recipes_config.exists():
            return {}, {}, f"Backup automation_recipes.json file not found:\n{self.recipes_config}"
        
        try:
            raw_data = json.loads(self.recipes_config.read_text(encoding="utf-8"))
            
            # Check if this is the new categorized format (exact same logic as original)
            if raw_data.get("_format") == "categorized" and "categories" in raw_data:
                # Store the categorized structure for UI
                recipes_categorized = raw_data["categories"]
                
                # Flatten recipes for backward compatibility
                recipes = {}
                for category_name, category_data in raw_data["categories"].items():
                    if "recipes" in category_data:
                        for recipe_name, recipe_data in category_data["recipes"].items():
                            recipes[recipe_name] = recipe_data
            else:
                # Legacy flat format - create a single category
                recipes = raw_data
                recipes_categorized = {
                    "All Recipes": {
                        "description": "All available automation recipes",
                        "recipes": raw_data
                    }
                }
            
            # If we successfully loaded from backup automation_recipes.json but Supabase failed, report it
            warning = f"Using backup_json/automation_recipes.json fallback - {db_error}" if db_error else None
            return recipes, recipes_categorized, warning
            
        except Exception as e:
            # Both Supabase and backup automation_recipes.json failed
            if db_error:
                return {}, {}, f"Supabase failed ({db_error}) and backup_json/automation_recipes.json failed: {e}"
            return {}, {}, f"Error loading backup_json/automation_recipes.json:\n{e}"
    
    def load_presets(self) -> Tuple[Dict[str, Any], Optional[str]]:
        """Load presets from database or backup_json/project_presets.json.
        
        Returns:
            Tuple of (presets_dict, error_message)
        """
        # Try Supabase database first (project_presets + preset_drawings tables)
        db_error = None
        if self.use_database and self.db_manager and self.db_manager.is_connected():
            try:
                project_presets_dict, error = self.db_manager.load_project_presets_from_preset_drawings_tables()
                if not error and project_presets_dict:
                    return project_presets_dict, None
                # If Supabase tables fail, capture error for reporting
                if error:
                    db_error = f"Supabase project_presets/preset_drawings tables failed: {error}"
                    print(f"{db_error}, falling back to backup_json/project_presets.json")
            except Exception as e:
                db_error = f"Supabase database error: {e}"
                print(f"{db_error}, falling back to backup_json/project_presets.json")
        
        # Fallback to backup JSON file
        if not self.presets_config.exists():
            return {}, f"Backup project_presets.json file not found:\n{self.presets_config}"
        
        try:
            presets = json.loads(self.presets_config.read_text(encoding="utf-8"))
            # If we successfully loaded from backup project_presets.json but Supabase failed, report it
            warning = f"Using backup_json/project_presets.json fallback - {db_error}" if db_error else None
            return presets, warning
        except Exception as e:
            # Both Supabase and backup project_presets.json failed
            if db_error:
                return {}, f"Supabase failed ({db_error}) and backup_json/project_presets.json failed: {e}"
            return {}, f"Error loading backup_json/project_presets.json:\n{e}"
    
    def load_templates(self, templates_path: Optional[Path] = None) -> Tuple[Dict[str, Any], Optional[str]]:
        """Load templates from database or backup_json/templates.json.
        
        Args:
            templates_path: Optional custom path, uses default if None
            
        Returns:
            Tuple of (templates_dict, error_message)
        """
        # Try Supabase database first (templates table) unless custom path specified
        db_error = None
        if not templates_path and self.use_database and self.db_manager and self.db_manager.is_connected():
            try:
                templates_dict, error = self.db_manager.load_templates_from_templates_table()
                if not error and templates_dict:
                    return templates_dict, None
                # If Supabase templates table fails, capture error for reporting
                if error:
                    db_error = f"Supabase templates table failed: {error}"
                    print(f"{db_error}, falling back to backup_json/templates.json")
            except Exception as e:
                db_error = f"Supabase database error: {e}"
                print(f"{db_error}, falling back to backup_json/templates.json")
        
        # Fallback to backup JSON file
        path = templates_path or self.templates_default
        
        if not path.exists():
            return {}, f"Backup templates.json file not found:\n{path}"
        
        try:
            templates = json.loads(path.read_text(encoding="utf-8"))
            # If we successfully loaded from backup templates.json but Supabase failed, report it
            warning = f"Using backup_json/templates.json fallback - {db_error}" if db_error else None
            return templates, warning
        except Exception as e:
            # Both Supabase and backup templates.json failed
            if db_error:
                return {}, f"Supabase failed ({db_error}) and backup_json/templates.json failed: {e}"
            return {}, f"Could not load templates:\n{e}"
    
    def load_rules(self, rules_path: Optional[Path] = None) -> Tuple[Dict[str, Rule], Optional[str]]:
        """Load rules from database or backup_json/dwg_filename_rules.json.
        
        Args:
            rules_path: Optional custom path, uses default if None
            
        Returns:
            Tuple of (rules_dict, error_message)
        """
        # Try Supabase database first (dwg_filename_rules table)
        db_error = None
        if self.use_database and self.db_manager and self.db_manager.is_connected():
            try:
                dwg_filename_rules_dict, error = self.db_manager.load_dwg_filename_rules_from_dwg_filename_rules_table()
                if not error and dwg_filename_rules_dict:
                    return dwg_filename_rules_dict, None
                # If Supabase dwg_filename_rules table fails, capture error for reporting
                if error:
                    db_error = f"Supabase dwg_filename_rules table failed: {error}"
                    print(f"{db_error}, falling back to backup_json/dwg_filename_rules.json")
            except Exception as e:
                db_error = f"Supabase database error: {e}"
                print(f"{db_error}, falling back to backup_json/dwg_filename_rules.json")
        
        # Fallback to backup JSON file
        path = rules_path or self.rules_default
        
        if not path.exists():
            return {}, f"Backup dwg_filename_rules.json file not found:\n{path}"
        
        try:
            rules = self._load_rules_json(path)
            # If we successfully loaded from backup dwg_filename_rules.json but Supabase failed, report it
            warning = f"Using backup_json/dwg_filename_rules.json fallback - {db_error}" if db_error else None
            return rules, warning
        except Exception as e:
            # Both Supabase and backup dwg_filename_rules.json failed
            if db_error:
                return {}, f"Supabase failed ({db_error}) and backup_json/dwg_filename_rules.json failed: {e}"
            return {}, f"Error loading rules:\n{e}"
    
    def _load_rules_json(self, p: Path) -> Dict[str, Rule]:
        """Load and parse rules JSON file (exact same logic as original function).
        
        Args:
            p: Path to the rules JSON file
            
        Returns:
            Dictionary mapping rule codes to Rule objects
        """
        # Import these constants from the original file context
        required_desc_codes = {"EXHIBIT", "OBJECT", "VEHICLE TRACKING", "BR-IMAGE"}
        
        data = json.loads(p.read_text(encoding="utf-8"))
        items = data if isinstance(data, list) else data.get("rules")
        if not isinstance(items, list):
            raise ValueError("Rules JSON must be a list or have a 'rules' key.")
        
        rules: Dict[str, Rule] = {}
        for row in items:
            code = str(row.get("File_Type_Code", "")).strip()
            if not code:
                continue
            
            folder_path = str(row.get("Folder_Path", ""))
            last_seg = _last_segment(folder_path)
            folder_short = "" if "[Subnumber]" in last_seg or "[ProjectNumber]" in last_seg else last_seg
            desc_req = bool(row.get("Description_Required", False))
            multi_ok = bool(row.get("Multi_Instance_Allowed", False))
            
            # Apply same business logic as original
            if code.split("-")[0].upper() in required_desc_codes:
                desc_req = True
                multi_ok = True
            
            rules[code] = Rule(
                code=code,
                folder_short=folder_short,
                Description_Required=desc_req,
                Multi_Instance_Allowed=multi_ok,
                filename_pattern=str(row.get("Filename_Pattern") or "")
            )
        
        return rules
    
    def load_preset_file(self, preset_file_path: Path) -> Tuple[Dict[str, Any], Optional[str]]:
        """Load a specific preset file (used by automation workers).
        
        Args:
            preset_file_path: Path to the preset file
            
        Returns:
            Tuple of (preset_data, error_message)
        """
        if not preset_file_path.exists():
            return {}, f"Preset file not found: {preset_file_path}"
        
        try:
            preset_data = json.loads(preset_file_path.read_text(encoding="utf-8"))
            return preset_data, None
        except Exception as e:
            return {}, f"Error loading preset file {preset_file_path}: {e}"
    
    def load_viewport_presets(self, viewport_path: Optional[Path] = None) -> Tuple[Dict[str, Any], Optional[str]]:
        """PLACEHOLDER: Load viewport presets from Viewport_Presets.json.
        
        WARNING: This is PLACEHOLDER functionality - viewport_presets table exists in Supabase
        but database loading function not yet implemented.
        
        Args:
            viewport_path: Optional custom path to Viewport_Presets.json
            
        Returns:
            Tuple of (viewport_presets_dict, error_message)
        """
        # PLACEHOLDER: Should try Supabase viewport_presets table first
        # TODO: Implement load_viewport_presets_from_viewport_presets_table() in DatabaseManager
        print("PLACEHOLDER WARNING: Loading viewport presets from JSON file - database integration pending")
        
        # Currently only using backup JSON fallback (no database integration yet)
        path = viewport_path or (self.app_dir / "backup_json" / "viewport_presets.json")
        
        if not path.exists():
            return {}, f"PLACEHOLDER: backup_json/viewport_presets.json not found:\n{path}"
        
        try:
            viewport_presets = json.loads(path.read_text(encoding="utf-8"))
            return viewport_presets, "PLACEHOLDER: Using backup_json/viewport_presets.json - database integration needed"
        except Exception as e:
            return {}, f"PLACEHOLDER: Error loading backup_json/viewport_presets.json:\n{e}"