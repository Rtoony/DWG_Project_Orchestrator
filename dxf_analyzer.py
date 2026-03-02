"""
CAD File Analyzer Module
A comprehensive tool for parsing DXF files and extracting structured data.
DWG files should be converted to DXF using AutoCAD Core Console before analysis.
"""

import json
import ezdxf
import subprocess
import tempfile
import shutil
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
import traceback


class DXFAnalyzer:
    """Main class for analyzing DXF files and extracting structured data."""
    
    def __init__(self):
        self.current_file = None
        self.document = None
        self.analysis_results = {}
        self.temp_dir = None
        
    def analyze_file(self, file_path: Path) -> Dict[str, Any]:
        """
        Analyze a single DXF file and extract all relevant information.
        DWG files must be converted to DXF using AutoCAD Core Console first.
        
        Args:
            file_path: Path to the DXF file
            
        Returns:
            Dictionary containing all extracted data
        """
        try:
            self.current_file = file_path
            dxf_file_path = file_path
            was_converted = False
            
            # Check if this is a DWG file - skip conversion, provide message
            if file_path.suffix.lower() == '.dwg':
                return {
                    "error": "DWG files must be converted to DXF using AutoCAD Core Console before analysis",
                    "original_file": str(file_path),
                    "extraction_timestamp": datetime.now().isoformat(),
                    "suggestion": "Use the Batch Operations tab to convert DWG files to DXF first"
                }
            
            # Try to read the DXF file
            try:
                self.document = ezdxf.readfile(str(dxf_file_path))
            except ezdxf.DXFStructureError:
                # Try with recovery mode for corrupted files
                from ezdxf import recover
                self.document, auditor = recover.readfile(str(dxf_file_path))
                if auditor.errors:
                    print(f"Warning: Found {len(auditor.errors)} errors in {file_path.name}")
            
            # Extract all data
            self.analysis_results = {
                "file_info": self._extract_file_info(),
                "drawing_metadata": self._extract_drawing_metadata(),
                "layers": self._extract_layers(),
                "entities": self._extract_entities(),
                "blocks": self._extract_blocks(),
                "text_objects": self._extract_text_objects(),
                "dimensions": self._extract_dimensions(),
                "statistics": self._generate_statistics(),
                "conversion_info": {
                    "was_converted_from_dwg": False,
                    "original_format": "DXF",
                    "converter_used": None
                },
                "extraction_timestamp": datetime.now().isoformat()
            }
            
            return self.analysis_results
            
        except Exception as e:
            error_result = {
                "error": str(e),
                "traceback": traceback.format_exc(),
                "file_path": str(file_path),
                "extraction_timestamp": datetime.now().isoformat()
            }
            return error_result
        
        finally:
            # Clean up temporary files if any were created
            if self.temp_dir and self.temp_dir.exists():
                try:
                    import shutil
                    shutil.rmtree(self.temp_dir)
                    self.temp_dir = None
                except:
                    pass
    
    
    
    def _extract_file_info(self) -> Dict[str, Any]:
        """Extract basic file information."""
        return {
            "file_path": str(self.current_file),
            "file_name": self.current_file.name,
            "file_size_bytes": self.current_file.stat().st_size,
            "dxf_version": self.document.dxfversion,
            "creation_date": self.current_file.stat().st_ctime,
            "modification_date": self.current_file.stat().st_mtime
        }
    
    def _extract_drawing_metadata(self) -> Dict[str, Any]:
        """Extract drawing metadata and settings."""
        header = self.document.header
        metadata = {
            "units": self._get_drawing_units(),
            "limits": self._get_drawing_limits(),
            "variables": {}
        }
        
        # Extract key header variables
        important_vars = [
            'ACADVER', 'DWGCODEPAGE', 'INSBASE', 'EXTMIN', 'EXTMAX',
            'LIMMIN', 'LIMMAX', 'ORTHOMODE', 'REGENMODE', 'FILLMODE',
            'QTEXTMODE', 'MIRRTEXT', 'LTSCALE', 'ATTMODE', 'TEXTSIZE',
            'TRACEWID', 'TEXTSTYLE', 'CLAYER', 'CELTYPE', 'CECOLOR',
            'CELTSCALE', 'DISPSILH', 'DIMSCALE', 'DIMASZ', 'DIMEXO',
            'DIMDLI', 'DIMRND', 'DIMDLE', 'DIMEXE', 'DIMTP', 'DIMTM'
        ]
        
        for var in important_vars:
            try:
                metadata["variables"][var] = header.get(var, None)
            except:
                continue
                
        return metadata
    
    def _get_drawing_units(self) -> str:
        """Determine the drawing units."""
        try:
            insunits = self.document.header.get('INSUNITS', 0)
            units_map = {
                0: 'Unitless',
                1: 'Inches',
                2: 'Feet',
                3: 'Miles',
                4: 'Millimeters',
                5: 'Centimeters',
                6: 'Meters',
                7: 'Kilometers',
                8: 'Microinches',
                9: 'Mils',
                10: 'Yards',
                11: 'Angstroms',
                12: 'Nanometers',
                13: 'Microns',
                14: 'Decimeters',
                15: 'Decameters',
                16: 'Hectometers',
                17: 'Gigameters',
                18: 'Astronomical units',
                19: 'Light years',
                20: 'Parsecs'
            }
            return units_map.get(insunits, f'Unknown ({insunits})')
        except:
            return 'Unknown'
    
    def _get_drawing_limits(self) -> Dict[str, Any]:
        """Get drawing limits."""
        try:
            limmin = self.document.header.get('LIMMIN', (0, 0))
            limmax = self.document.header.get('LIMMAX', (0, 0))
            extmin = self.document.header.get('EXTMIN', (0, 0, 0))
            extmax = self.document.header.get('EXTMAX', (0, 0, 0))
            
            return {
                "limits_min": list(limmin),
                "limits_max": list(limmax),
                "extents_min": list(extmin),
                "extents_max": list(extmax)
            }
        except:
            return {"error": "Could not extract limits"}
    
    def _extract_layers(self) -> List[Dict[str, Any]]:
        """Extract all layer information."""
        layers = []
        
        for layer in self.document.layers:
            layer_info = {
                "name": layer.dxf.name,
                "color": layer.dxf.color,
                "linetype": layer.dxf.linetype,
                "lineweight": getattr(layer.dxf, 'lineweight', None),
                "plot": getattr(layer.dxf, 'plot', True),
                "frozen": layer.is_frozen(),
                "locked": layer.is_locked(),
                "on": layer.is_on()
            }
            layers.append(layer_info)
        
        return layers
    
    def _extract_entities(self) -> Dict[str, List[Dict[str, Any]]]:
        """Extract all entities from modelspace and paperspace."""
        entities = {
            "modelspace": [],
            "paperspace_layouts": {}
        }
        
        # Extract modelspace entities
        msp = self.document.modelspace()
        entities["modelspace"] = self._extract_entities_from_space(msp)
        
        # Extract paperspace entities
        for layout in self.document.layout_names_in_taborder():
            if layout.lower() != 'model':
                psp = self.document.paperspace(layout)
                entities["paperspace_layouts"][layout] = self._extract_entities_from_space(psp)
        
        return entities
    
    def _extract_entities_from_space(self, space) -> List[Dict[str, Any]]:
        """Extract entities from a specific space (modelspace or paperspace)."""
        entities = []
        
        for entity in space:
            entity_info = {
                "type": entity.dxftype(),
                "layer": entity.dxf.layer,
                "color": entity.dxf.color,
                "linetype": getattr(entity.dxf, 'linetype', 'BYLAYER'),
                "handle": entity.dxf.handle,
                "geometry": self._extract_entity_geometry(entity)
            }
            entities.append(entity_info)
        
        return entities
    
    def _extract_entity_geometry(self, entity) -> Dict[str, Any]:
        """Extract geometric data from an entity."""
        entity_type = entity.dxftype()
        
        try:
            if entity_type == 'LINE':
                return {
                    "start_point": list(entity.dxf.start),
                    "end_point": list(entity.dxf.end),
                    "length": (entity.dxf.end - entity.dxf.start).magnitude
                }
            
            elif entity_type == 'CIRCLE':
                return {
                    "center": list(entity.dxf.center),
                    "radius": entity.dxf.radius,
                    "circumference": 2 * 3.14159 * entity.dxf.radius,
                    "area": 3.14159 * entity.dxf.radius ** 2
                }
            
            elif entity_type == 'ARC':
                return {
                    "center": list(entity.dxf.center),
                    "radius": entity.dxf.radius,
                    "start_angle": entity.dxf.start_angle,
                    "end_angle": entity.dxf.end_angle
                }
            
            elif entity_type == 'LWPOLYLINE':
                points = [(point[0], point[1]) for point in entity.get_points()]
                return {
                    "points": points,
                    "closed": entity.closed,
                    "point_count": len(points)
                }
            
            elif entity_type == 'POLYLINE':
                points = [(vertex.dxf.location[0], vertex.dxf.location[1]) for vertex in entity.vertices]
                return {
                    "points": points,
                    "closed": entity.is_closed,
                    "point_count": len(points)
                }
            
            elif entity_type == 'TEXT':
                return {
                    "text": entity.dxf.text,
                    "insert_point": list(entity.dxf.insert),
                    "height": entity.dxf.height,
                    "rotation": getattr(entity.dxf, 'rotation', 0),
                    "style": getattr(entity.dxf, 'style', 'STANDARD')
                }
            
            elif entity_type == 'MTEXT':
                return {
                    "text": entity.plain_text(),
                    "insert_point": list(entity.dxf.insert),
                    "char_height": entity.dxf.char_height,
                    "width": getattr(entity.dxf, 'width', None),
                    "attachment_point": getattr(entity.dxf, 'attachment_point', 1)
                }
            
            elif entity_type == 'INSERT':
                return {
                    "block_name": entity.dxf.name,
                    "insert_point": list(entity.dxf.insert),
                    "scale_x": getattr(entity.dxf, 'xscale', 1),
                    "scale_y": getattr(entity.dxf, 'yscale', 1),
                    "scale_z": getattr(entity.dxf, 'zscale', 1),
                    "rotation": getattr(entity.dxf, 'rotation', 0)
                }
            
            elif entity_type == 'DIMENSION':
                return {
                    "dim_type": getattr(entity.dxf, 'dimtype', 'unknown'),
                    "measurement": getattr(entity, 'get_measurement', lambda: None)()
                }
            
            else:
                # For other entity types, return basic info
                return {
                    "entity_type": entity_type,
                    "properties": "Geometry extraction not implemented for this entity type"
                }
                
        except Exception as e:
            return {
                "error": f"Failed to extract geometry: {str(e)}",
                "entity_type": entity_type
            }
    
    def _extract_blocks(self) -> List[Dict[str, Any]]:
        """Extract block definitions."""
        blocks = []
        
        for block in self.document.blocks:
            if not block.name.startswith('*'):  # Skip anonymous blocks
                block_info = {
                    "name": block.name,
                    "base_point": list(block.block.dxf.base_point) if hasattr(block.block.dxf, 'base_point') else [0, 0, 0],
                    "entities": self._extract_entities_from_space(block),
                    "entity_count": len(list(block))
                }
                blocks.append(block_info)
        
        return blocks
    
    def _extract_text_objects(self) -> List[Dict[str, Any]]:
        """Extract all text objects for easier analysis."""
        text_objects = []
        
        # Search in modelspace
        msp = self.document.modelspace()
        for entity in msp.query('TEXT MTEXT'):
            text_info = {
                "type": entity.dxftype(),
                "text": entity.dxf.text if entity.dxftype() == 'TEXT' else entity.plain_text(),
                "layer": entity.dxf.layer,
                "location": list(entity.dxf.insert),
                "height": entity.dxf.height if entity.dxftype() == 'TEXT' else entity.dxf.char_height,
                "space": "modelspace"
            }
            text_objects.append(text_info)
        
        # Search in paperspace layouts
        for layout_name in self.document.layout_names_in_taborder():
            if layout_name.lower() != 'model':
                psp = self.document.paperspace(layout_name)
                for entity in psp.query('TEXT MTEXT'):
                    text_info = {
                        "type": entity.dxftype(),
                        "text": entity.dxf.text if entity.dxftype() == 'TEXT' else entity.plain_text(),
                        "layer": entity.dxf.layer,
                        "location": list(entity.dxf.insert),
                        "height": entity.dxf.height if entity.dxftype() == 'TEXT' else entity.dxf.char_height,
                        "space": f"paperspace_{layout_name}"
                    }
                    text_objects.append(text_info)
        
        return text_objects
    
    def _extract_dimensions(self) -> List[Dict[str, Any]]:
        """Extract dimension objects."""
        dimensions = []
        
        # Search in all spaces
        spaces = [("modelspace", self.document.modelspace())]
        for layout_name in self.document.layout_names_in_taborder():
            if layout_name.lower() != 'model':
                spaces.append((f"paperspace_{layout_name}", self.document.paperspace(layout_name)))
        
        for space_name, space in spaces:
            for entity in space.query('DIMENSION'):
                dim_info = {
                    "layer": entity.dxf.layer,
                    "dim_type": getattr(entity.dxf, 'dimtype', 'unknown'),
                    "measurement": getattr(entity, 'get_measurement', lambda: None)(),
                    "space": space_name
                }
                dimensions.append(dim_info)
        
        return dimensions
    
    def _generate_statistics(self) -> Dict[str, Any]:
        """Generate statistical summary of the drawing."""
        stats = {
            "entity_counts": {},
            "layer_count": len(self.document.layers),
            "block_count": len([b for b in self.document.blocks if not b.name.startswith('*')]),
            "layout_count": len(self.document.layout_names_in_taborder()) - 1,  # Exclude Model
            "text_object_count": 0,
            "dimension_count": 0
        }
        
        # Count entities by type
        all_spaces = [self.document.modelspace()]
        for layout_name in self.document.layout_names_in_taborder():
            if layout_name.lower() != 'model':
                all_spaces.append(self.document.paperspace(layout_name))
        
        for space in all_spaces:
            for entity in space:
                entity_type = entity.dxftype()
                stats["entity_counts"][entity_type] = stats["entity_counts"].get(entity_type, 0) + 1
                
                if entity_type in ['TEXT', 'MTEXT']:
                    stats["text_object_count"] += 1
                elif entity_type == 'DIMENSION':
                    stats["dimension_count"] += 1
        
        return stats
    
    def export_to_json(self, output_path: Path, pretty_print: bool = True) -> bool:
        """
        Export analysis results to JSON file.
        
        Args:
            output_path: Path where to save the JSON file
            pretty_print: Whether to format JSON for readability
            
        Returns:
            True if successful, False otherwise
        """
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                if pretty_print:
                    json.dump(self.analysis_results, f, indent=2, ensure_ascii=False)
                else:
                    json.dump(self.analysis_results, f, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error exporting to JSON: {e}")
            return False
    
    def batch_analyze(self, input_folder: Path, output_folder: Path, file_pattern: str = "*.dxf") -> Dict[str, Any]:
        """
        Analyze multiple DXF and DWG files in a folder.
        
        Args:
            input_folder: Folder containing DXF/DWG files
            output_folder: Folder to save JSON results
            file_pattern: File pattern to match (default: "*.dxf"). Use "*.*" for both DXF and DWG
            
        Returns:
            Summary of batch processing results
        """
        input_folder = Path(input_folder)
        output_folder = Path(output_folder)
        
        # Create output folder if it doesn't exist
        output_folder.mkdir(parents=True, exist_ok=True)
        
        # Find CAD files
        if file_pattern == "*.*":
            # Look for both DXF and DWG files
            cad_files = list(input_folder.glob("*.dxf")) + list(input_folder.glob("*.dwg"))
        else:
            cad_files = list(input_folder.glob(file_pattern))
        
        results = {
            "total_files": len(cad_files),
            "processed_files": 0,
            "failed_files": 0,
            "failed_file_names": [],
            "processing_summary": {},
            "dwg_files_processed": 0,
            "dxf_files_processed": 0
        }
        
        for cad_file in cad_files:
            try:
                print(f"Processing: {cad_file.name}")
                
                # Analyze the file
                analysis_result = self.analyze_file(cad_file)
                
                # Generate output filename
                json_filename = cad_file.stem + "_analysis.json"
                json_path = output_folder / json_filename
                
                # Export to JSON
                if self.export_to_json(json_path):
                    results["processed_files"] += 1
                    
                    # Track file type statistics
                    if cad_file.suffix.lower() == '.dwg':
                        results["dwg_files_processed"] += 1
                    else:
                        results["dxf_files_processed"] += 1
                    
                    results["processing_summary"][cad_file.name] = {
                        "status": "success",
                        "output_file": str(json_path),
                        "original_format": cad_file.suffix.upper(),
                        "entity_count": sum(analysis_result.get("statistics", {}).get("entity_counts", {}).values()),
                        "was_converted": analysis_result.get("conversion_info", {}).get("was_converted_from_dwg", False)
                    }
                else:
                    results["failed_files"] += 1
                    results["failed_file_names"].append(cad_file.name)
                    results["processing_summary"][cad_file.name] = {
                        "status": "export_failed",
                        "error": "Failed to export JSON"
                    }
                    
            except Exception as e:
                results["failed_files"] += 1
                results["failed_file_names"].append(cad_file.name)
                results["processing_summary"][cad_file.name] = {
                    "status": "analysis_failed",
                    "error": str(e)
                }
        
        return results


def test_analyzer(test_file_path: Optional[str] = None):
    """Test function to demonstrate the DXF analyzer."""
    analyzer = DXFAnalyzer()
    
    if test_file_path and Path(test_file_path).exists():
        print(f"Analyzing: {test_file_path}")
        result = analyzer.analyze_file(Path(test_file_path))
        
        # Print summary
        if "error" not in result:
            print("\n=== Analysis Summary ===")
            print(f"DXF Version: {result['file_info']['dxf_version']}")
            print(f"Units: {result['drawing_metadata']['units']}")
            print(f"Layers: {result['statistics']['layer_count']}")
            print(f"Total Entities: {sum(result['statistics']['entity_counts'].values())}")
            print(f"Text Objects: {result['statistics']['text_object_count']}")
            
            # Export to JSON
            json_path = Path(test_file_path).with_suffix('.json')
            if analyzer.export_to_json(json_path):
                print(f"Results exported to: {json_path}")
        else:
            print(f"Error: {result['error']}")
    else:
        print("No test file provided or file doesn't exist.")
        print("Usage: test_analyzer('/path/to/test.dxf')")


if __name__ == "__main__":
    # Example usage
    test_analyzer()