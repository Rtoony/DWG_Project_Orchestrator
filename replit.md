# DWG Project Orchestrator

## Overview

The DWG Project Orchestrator is a standalone CAD automation tool designed for managing and processing AutoCAD DWG files. It provides a PyQt6-based desktop application with capabilities for file analysis, batch operations, project setup, and automated drawing creation. The system integrates with AutoCAD through COM automation and includes comprehensive DXF parsing capabilities for extracting structured data from CAD files.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture
- **Desktop Application Framework**: PyQt6-based GUI with tabbed interface for different functionality areas
- **Main Application Structure**: Single-file application (`dwg_project_orchestrator.py`) containing the complete UI and orchestration logic
- **UI Components**: Multi-tab layout supporting project setup, file analysis, batch operations, and automation recipe management
- **Enhanced Drawing Analysis**: Supports both direct DXF analysis and loading of existing JSON analysis files for review

### Backend Architecture
- **Configuration Management**: Centralized configuration system (`config_manager.py`) that handles JSON-based settings for standalone operation
- **CAD File Processing**: Dedicated DXF analysis engine (`dxf_analyzer.py`) using the ezdxf library for parsing CAD file structures
- **AutoCAD Integration**: COM automation interface using win32com for direct AutoCAD manipulation and control
- **Batch Processing**: AutoCAD Core Console integration for headless DWG to DXF conversion operations

### Data Storage Solutions
- **Primary Storage**: JSON-based configuration files stored in `backup_json/` directory
- **Configuration Files**:
  - `automation_recipes.json`: Categorized automation scripts and procedures including DWG conversion workflows
  - `project_presets.json`: Template definitions for different project types
  - `dwg_filename_rules.json`: File naming conventions and validation rules
  - `templates.json`: CAD template file mappings
  - `viewport_presets.json`: Layout and viewport configuration presets
  - `layer_standards.json`: Comprehensive CAD layer standards database (361 layers) with properties, colors, linetypes, and usage rules
- **Analysis Results**: JSON export files containing extracted CAD file data and metadata

### File Processing Pipeline
- **DWG to DXF Conversion**: Automated conversion using AutoCAD Core Console via dynamically generated scripts
- **DXF Analysis**: Comprehensive parsing of CAD file structures including entities, layers, blocks, and metadata
- **Batch Operations**: Multi-file processing capabilities with progress tracking and error handling
- **JSON Analysis Loading**: Capability to load and display previously generated analysis results

### Automation Framework
- **Recipe System**: Categorized automation scripts supporting multiple execution engines:
  - Core Console (headless AutoCAD) - primary conversion method
  - PyAutoCAD (full GUI automation)
  - Python Direct (native Python operations)
- **Script Management**: External script file integration with AutoLISP and AutoCAD script support
- **Project Templates**: Configurable project setups with drawing type presets and naming conventions

## External Dependencies

### CAD Software Integration
- **AutoCAD**: Primary CAD platform requiring both full GUI and Core Console installations
- **AutoCAD COM Interface**: win32com.client for programmatic AutoCAD control and automation

### Python Libraries
- **PyQt6**: Desktop application framework for the user interface
- **ezdxf**: DXF file parsing and manipulation library for CAD file analysis
- **pywin32**: Windows COM automation and system integration

### Development Tools
- **Subprocess**: External process management for AutoCAD Core Console operations
- **Pathlib**: Modern file path handling and manipulation
- **JSON**: Configuration and analysis result storage and management

### File System Requirements
- **Temporary File Management**: Automated cleanup of converted DXF files and processing artifacts
- **Script File Dependencies**: External AutoLISP and AutoCAD script files for automation recipes
- **JSON Analysis Files**: Storage and retrieval of CAD analysis results for later review