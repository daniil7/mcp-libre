"""
LibreOffice MCP Extension - MCP Server Module

This module implements an embedded MCP server that integrates with LibreOffice
via the UNO API, providing real-time document manipulation capabilities.
"""

import asyncio
import json
import logging
from typing import Dict, Any, Optional, List
from uno_bridge import (
    UNOBridge, SLIDE_LAYOUTS, PARAGRAPH_ALIGNMENTS, VERTICAL_ALIGNMENTS,
    SHAPE_TYPES, IMPRESS_EXPORT_FILTERS, WRITER_EXPORT_FILTERS, CHART_TYPES,
    LIST_TYPES, NUMBER_FORMATS,
)

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class LibreOfficeMCPServer:
    """Embedded MCP server for LibreOffice plugin"""
    
    def __init__(self):
        """Initialize the MCP server"""
        self.uno_bridge = UNOBridge()
        self.tools = {}
        self._register_tools()
        logger.info("LibreOffice MCP Server initialized")
    
    def _register_tools(self):
        """Register all available MCP tools"""
        
        # Document creation tools
        self.tools["create_document_live"] = {
            "description": "Create a new document in LibreOffice",
            "parameters": {
                "type": "object",
                "properties": {
                    "doc_type": {
                        "type": "string",
                        "enum": ["writer", "calc", "impress", "draw"],
                        "description": "Type of document to create",
                        "default": "writer"
                    }
                }
            },
            "handler": self.create_document_live
        }
        
        # Text manipulation tools
        self.tools["insert_text_live"] = {
            "description": "Insert text into the currently active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "text": {
                        "type": "string",
                        "description": "Text to insert"
                    },
                    "position": {
                        "type": "integer",
                        "description": "Position to insert at (optional, defaults to cursor position)"
                    }
                },
                "required": ["text"]
            },
            "handler": self.insert_text_live
        }
        
        # Document info tools
        self.tools["get_document_info_live"] = {
            "description": "Get information about the currently active document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_document_info_live
        }
        
        # Text formatting tools
        self.tools["format_text_live"] = {
            "description": "Apply formatting to selected text in active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "bold": {
                        "type": "boolean",
                        "description": "Apply bold formatting"
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Apply italic formatting"
                    },
                    "underline": {
                        "type": "boolean",
                        "description": "Apply underline formatting"
                    },
                    "font_size": {
                        "type": "number",
                        "description": "Font size in points"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "Font family name"
                    }
                }
            },
            "handler": self.format_text_live
        }
        
        # Document saving tools
        self.tools["save_document_live"] = {
            "description": "Save the currently active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to save document to (optional, saves to current location if not specified)"
                    }
                }
            },
            "handler": self.save_document_live
        }
        
        # Document export tools
        self.tools["export_document_live"] = {
            "description": "Export the currently active document to a different format",
            "parameters": {
                "type": "object",
                "properties": {
                    "export_format": {
                        "type": "string",
                        "enum": sorted(set(WRITER_EXPORT_FILTERS) | set(IMPRESS_EXPORT_FILTERS)),
                        "description": "Format to export to (available formats depend on document type: Writer supports docx/doc/odt/txt/rtf/html/pdf, Impress supports pptx/ppt/odp/png/html/pdf)"
                    },
                    "file_path": {
                        "type": "string",
                        "description": "Path to export document to"
                    }
                },
                "required": ["export_format", "file_path"]
            },
            "handler": self.export_document_live
        }
        
        # Content reading tools
        self.tools["get_text_content_live"] = {
            "description": "Get the text content of the currently active document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_text_content_live
        }
        
        # Document list tools
        self.tools["list_open_documents"] = {
            "description": "List all currently open documents in LibreOffice",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.list_open_documents
        }

        # Impress slide tools
        self.tools["list_slides_live"] = {
            "description": "List all slides of the active Impress presentation, including each slide's layout and shapes (role, text, position, size)",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.list_slides_live
        }

        self.tools["add_slide_live"] = {
            "description": "Insert a new slide into the active Impress presentation",
            "parameters": {
                "type": "object",
                "properties": {
                    "layout": {
                        "type": "string",
                        "enum": list(SLIDE_LAYOUTS.keys()),
                        "description": "Slide layout (which placeholders it has)",
                        "default": "title_content"
                    },
                    "index": {
                        "type": "integer",
                        "description": "Position to insert the slide at (0-based, defaults to the end)"
                    }
                }
            },
            "handler": self.add_slide_live
        }

        self.tools["set_slide_layout_live"] = {
            "description": "Change which placeholders a slide has",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {
                        "type": "integer",
                        "description": "Index of the slide to modify (0-based)",
                        "default": 0
                    },
                    "layout": {
                        "type": "string",
                        "enum": list(SLIDE_LAYOUTS.keys()),
                        "description": "Slide layout to apply"
                    }
                },
                "required": ["layout"]
            },
            "handler": self.set_slide_layout_live
        }

        self.tools["set_slide_title_live"] = {
            "description": "Set a slide's title text, with optional text alignment and centering on the slide",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {
                        "type": "integer",
                        "description": "Index of the slide to modify (0-based)",
                        "default": 0
                    },
                    "text": {
                        "type": "string",
                        "description": "Title text"
                    },
                    "align": {
                        "type": "string",
                        "enum": list(PARAGRAPH_ALIGNMENTS.keys()),
                        "description": "Horizontal text alignment within the title box",
                        "default": "center"
                    },
                    "vertical_align": {
                        "type": "string",
                        "enum": list(VERTICAL_ALIGNMENTS.keys()),
                        "description": "Vertical text alignment within the title box (optional, leaves template default if omitted)"
                    },
                    "center_box": {
                        "type": "boolean",
                        "description": "Reposition the title box itself to the geometric center of the slide, instead of its template position",
                        "default": False
                    }
                },
                "required": ["text"]
            },
            "handler": self.set_slide_title_live
        }

        self.tools["set_slide_text_live"] = {
            "description": "Set the text of an arbitrary shape on a slide (e.g. a content/outline placeholder); use list_slides_live to find shape_index",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {
                        "type": "integer",
                        "description": "Index of the slide to modify (0-based)",
                        "default": 0
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "Index of the shape on the slide (0-based, see list_slides_live)",
                        "default": 1
                    },
                    "text": {
                        "type": "string",
                        "description": "Text to set"
                    },
                    "align": {
                        "type": "string",
                        "enum": list(PARAGRAPH_ALIGNMENTS.keys()),
                        "description": "Horizontal text alignment within the shape (optional)"
                    },
                    "vertical_align": {
                        "type": "string",
                        "enum": list(VERTICAL_ALIGNMENTS.keys()),
                        "description": "Vertical text alignment within the shape (optional)"
                    }
                },
                "required": ["text"]
            },
            "handler": self.set_slide_text_live
        }

        self.tools["remove_slide_live"] = {
            "description": "Remove a slide from the presentation",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {
                        "type": "integer",
                        "description": "Index of the slide to remove (0-based)",
                        "default": 0
                    }
                }
            },
            "handler": self.remove_slide_live
        }

        self.tools["duplicate_slide_live"] = {
            "description": "Duplicate a slide; the copy is inserted immediately after the original",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {
                        "type": "integer",
                        "description": "Index of the slide to duplicate (0-based)",
                        "default": 0
                    }
                }
            },
            "handler": self.duplicate_slide_live
        }

        self.tools["set_slide_background_live"] = {
            "description": "Set the solid background fill color of a slide, or all slides",
            "parameters": {
                "type": "object",
                "properties": {
                    "color": {
                        "type": "string",
                        "description": "Hex color, e.g. '#3366CC' or '3366CC'"
                    },
                    "slide_index": {
                        "type": "integer",
                        "description": "Index of the slide to modify (0-based, ignored if apply_to_all is true)",
                        "default": 0
                    },
                    "apply_to_all": {
                        "type": "boolean",
                        "description": "Apply to every slide in the presentation",
                        "default": False
                    }
                },
                "required": ["color"]
            },
            "handler": self.set_slide_background_live
        }

        self.tools["add_shape_live"] = {
            "description": "Add a rectangle/ellipse/line/text shape to a slide, with optional fill/line color and text",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {
                        "type": "integer",
                        "description": "Index of the slide to add the shape to (0-based)",
                        "default": 0
                    },
                    "shape_type": {
                        "type": "string",
                        "enum": list(SHAPE_TYPES.keys()),
                        "description": "Kind of shape to create",
                        "default": "rectangle"
                    },
                    "x": {"type": "integer", "description": "X position in 1/100 mm", "default": 1000},
                    "y": {"type": "integer", "description": "Y position in 1/100 mm", "default": 1000},
                    "width": {"type": "integer", "description": "Width in 1/100 mm", "default": 5000},
                    "height": {"type": "integer", "description": "Height in 1/100 mm", "default": 3000},
                    "text": {"type": "string", "description": "Text to place inside the shape (not supported for 'line')"},
                    "fill_color": {"type": "string", "description": "Hex fill color (e.g. '#3366CC'), or 'none' for no fill"},
                    "line_color": {"type": "string", "description": "Hex line/border color (e.g. '#000000'), or 'none' for no line"},
                    "shadow": {"type": "boolean", "description": "Enable a drop shadow on the shape"},
                    "shadow_color": {"type": "string", "description": "Hex shadow color (e.g. '#808080'); only applies when shadow is true"}
                }
            },
            "handler": self.add_shape_live
        }

        self.tools["remove_shape_live"] = {
            "description": ("Remove a shape from a slide. Also useful for visual verification of charts: since "
                           "chart objects don't render in exported PNG/PDF previews, add a labeled rectangle "
                           "(e.g. text='{DIAGRAM}') at the chart's position/size as a temporary marker before "
                           "exporting to check the slide, then call this to remove it. Add the marker last so "
                           "removing it doesn't shift other shapes' indices."),
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the shape to remove (0-based)", "default": 0}
                }
            },
            "handler": self.remove_shape_live
        }

        self.tools["set_shape_style_live"] = {
            "description": "Change the fill color, line color, and/or drop shadow of an existing shape on a slide (use list_slides_live to find shape_index)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the shape on the slide (0-based)", "default": 0},
                    "fill_color": {"type": "string", "description": "Hex fill color (e.g. '#3366CC'), or 'none' for no fill"},
                    "line_color": {"type": "string", "description": "Hex line/border color (e.g. '#000000'), or 'none' for no line"},
                    "shadow": {"type": "boolean", "description": "Enable/disable a drop shadow on the shape"},
                    "shadow_color": {"type": "string", "description": "Hex shadow color (e.g. '#808080'); only applies when shadow is true"}
                }
            },
            "handler": self.set_shape_style_live
        }

        self.tools["set_shape_position_live"] = {
            "description": "Move and/or resize an existing shape on a slide (use list_slides_live to find shape_index)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the shape on the slide (0-based)", "default": 0},
                    "x": {"type": "integer", "description": "New X position in 1/100 mm (omit to keep current)"},
                    "y": {"type": "integer", "description": "New Y position in 1/100 mm (omit to keep current)"},
                    "width": {"type": "integer", "description": "New width in 1/100 mm (omit to keep current)"},
                    "height": {"type": "integer", "description": "New height in 1/100 mm (omit to keep current)"}
                }
            },
            "handler": self.set_shape_position_live
        }

        self.tools["set_shape_hyperlink_live"] = {
            "description": ("Set (or clear, with url='') a shape's hyperlink - the whole shape becomes "
                           "clickable during a slideshow. Impress only supports hyperlinking a whole shape, "
                           "not a specific word/run of text within it like Writer does."),
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the shape on the slide (0-based)", "default": 0},
                    "url": {"type": "string", "description": "URL to link to; empty string clears the hyperlink"}
                },
                "required": ["url"]
            },
            "handler": self.set_shape_hyperlink_live
        }

        self.tools["set_list_formatting_live"] = {
            "description": "Turn every paragraph of a slide shape into a bulleted, numbered, or plain (no marker) list",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the shape on the slide (0-based)", "default": 0},
                    "list_type": {"type": "string", "enum": list(LIST_TYPES), "description": "Kind of list", "default": "bullet"},
                    "level": {"type": "integer", "description": "Outline depth for every paragraph (0 = top level); omit to keep each paragraph's current depth"},
                    "bullet_char": {"type": "string", "description": "Custom bullet character (list_type='bullet' only); omit to keep the master slide's default bullet"},
                    "number_format": {"type": "string", "enum": list(NUMBER_FORMATS.keys()), "description": "Marker style for list_type='number': '1'/'a'/'A'/'i'/'I'", "default": "1"},
                    "start_at": {"type": "integer", "description": "Starting number for list_type='number' (e.g. 5 to start counting at 5 instead of 1)"}
                }
            },
            "handler": self.set_list_formatting_live
        }

        self.tools["insert_image_live"] = {
            "description": "Insert an image onto a slide from a local file path; defaults to the image's natural size, centered on the slide",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "image_path": {"type": "string", "description": "Local filesystem path to the image file"},
                    "x": {"type": "integer", "description": "X position in 1/100 mm (omit to center)"},
                    "y": {"type": "integer", "description": "Y position in 1/100 mm (omit to center)"},
                    "width": {"type": "integer", "description": "Width in 1/100 mm (omit for the image's natural size)"},
                    "height": {"type": "integer", "description": "Height in 1/100 mm (omit for the image's natural size)"}
                },
                "required": ["image_path"]
            },
            "handler": self.insert_image_live
        }

        self.tools["replace_image_live"] = {
            "description": "Swap the picture of an existing image shape, keeping its position (and by default its size)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the image shape on the slide (0-based)", "default": 0},
                    "image_path": {"type": "string", "description": "Local filesystem path to the new image file"},
                    "keep_size": {"type": "boolean", "description": "Keep the shape's current size instead of resizing to the new image's natural size", "default": True}
                },
                "required": ["image_path"]
            },
            "handler": self.replace_image_live
        }

        self.tools["format_slide_text_live"] = {
            "description": "Apply character formatting (bold/italic/underline/strikeout/size/font/color/highlight/shadow/outline) to all text in a slide shape",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the shape on the slide (0-based)", "default": 0},
                    "bold": {"type": "boolean", "description": "Apply bold formatting"},
                    "italic": {"type": "boolean", "description": "Apply italic formatting"},
                    "underline": {"type": "boolean", "description": "Apply underline formatting"},
                    "strikeout": {"type": "boolean", "description": "Apply strikethrough formatting"},
                    "font_size": {"type": "number", "description": "Font size in points"},
                    "font_name": {"type": "string", "description": "Font family name"},
                    "color": {"type": "string", "description": "Hex text color, e.g. '#FFFFFF'"},
                    "highlight_color": {"type": "string", "description": "Hex highlight/background color behind the text (e.g. '#FFFF00'), or 'none' to clear it"},
                    "shadow": {"type": "boolean", "description": "Apply a drop shadow to the character glyphs (crude effect; for a proper shadow use set_shape_style_live's shadow on the shape itself)"},
                    "outline": {"type": "boolean", "description": "Render the text as hollow/outlined (contour) letters"}
                }
            },
            "handler": self.format_slide_text_live
        }

        self.tools["set_slide_footer_live"] = {
            "description": "Set footer text, slide number, and date visibility on one slide or all slides (Impress slides support a footer/date/number, not a header)",
            "parameters": {
                "type": "object",
                "properties": {
                    "footer_text": {"type": "string", "description": "Footer text"},
                    "show_footer": {"type": "boolean", "description": "Show/hide the footer"},
                    "show_slide_number": {"type": "boolean", "description": "Show/hide the slide number"},
                    "show_date": {"type": "boolean", "description": "Show/hide the date"},
                    "date_text": {"type": "string", "description": "Fixed date text to display (sets the date to a fixed value rather than auto-updating)"},
                    "slide_index": {"type": "integer", "description": "Index of the slide to modify (0-based); required if apply_to_all is false"},
                    "apply_to_all": {"type": "boolean", "description": "Apply to every slide in the presentation", "default": True}
                }
            },
            "handler": self.set_slide_footer_live
        }

        self.tools["insert_table_live"] = {
            "description": "Insert a table onto a slide, optionally pre-filled with cell text",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "rows": {"type": "integer", "description": "Number of rows", "default": 2},
                    "cols": {"type": "integer", "description": "Number of columns", "default": 2},
                    "x": {"type": "integer", "description": "X position in 1/100 mm", "default": 1000},
                    "y": {"type": "integer", "description": "Y position in 1/100 mm", "default": 1000},
                    "width": {"type": "integer", "description": "Width in 1/100 mm", "default": 20000},
                    "height": {"type": "integer", "description": "Height in 1/100 mm", "default": 8000},
                    "data": {
                        "type": "array",
                        "items": {"type": "array", "items": {"type": "string"}},
                        "description": "Row-major cell text, e.g. [[\"Header A\",\"Header B\"],[\"1\",\"2\"]]; extra rows/cols beyond the table size are ignored"
                    }
                }
            },
            "handler": self.insert_table_live
        }

        self.tools["set_table_cell_live"] = {
            "description": "Set the text of a single table cell (use list_slides_live to find shape_index)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the table shape on the slide (0-based)", "default": 0},
                    "row": {"type": "integer", "description": "Row index (0-based)", "default": 0},
                    "col": {"type": "integer", "description": "Column index (0-based)", "default": 0},
                    "text": {"type": "string", "description": "Cell text"}
                },
                "required": ["text"]
            },
            "handler": self.set_table_cell_live
        }

        self.tools["get_table_data_live"] = {
            "description": "Read back a table's full contents (dimensions and every cell's text)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the table shape on the slide (0-based)", "default": 0}
                }
            },
            "handler": self.get_table_data_live
        }

        self.tools["insert_table_row_live"] = {
            "description": "Insert a row into an existing table (appends at the end if index is omitted)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the table shape on the slide (0-based)", "default": 0},
                    "index": {"type": "integer", "description": "Row position to insert at (0-based); omit to append at the end"}
                }
            },
            "handler": self.insert_table_row_live
        }

        self.tools["remove_table_row_live"] = {
            "description": "Remove a row from an existing table",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the table shape on the slide (0-based)", "default": 0},
                    "index": {"type": "integer", "description": "Row index to remove (0-based)", "default": 0}
                }
            },
            "handler": self.remove_table_row_live
        }

        self.tools["insert_table_column_live"] = {
            "description": "Insert a column into an existing table (appends at the end if index is omitted)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the table shape on the slide (0-based)", "default": 0},
                    "index": {"type": "integer", "description": "Column position to insert at (0-based); omit to append at the end"}
                }
            },
            "handler": self.insert_table_column_live
        }

        self.tools["remove_table_column_live"] = {
            "description": "Remove a column from an existing table",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the table shape on the slide (0-based)", "default": 0},
                    "index": {"type": "integer", "description": "Column index to remove (0-based)", "default": 0}
                }
            },
            "handler": self.remove_table_column_live
        }

        self.tools["insert_chart_live"] = {
            "description": ("Insert a chart onto a slide, optionally with data. Pair categories[i] with series[name][i] - "
                           "e.g. categories=['Q1','Q2'], series={'Sales':[10,20],'Costs':[5,8]}. Note: the chart "
                           "won't appear in exported PNG/PDF previews used for visual verification (its data is "
                           "still saved correctly and shows up when opened in the real Impress UI) - see "
                           "remove_shape_live for a way to mark its location before a visual check."),
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "chart_type": {"type": "string", "enum": list(CHART_TYPES.keys()), "description": "Kind of chart", "default": "bar"},
                    "x": {"type": "integer", "description": "X position in 1/100 mm", "default": 1000},
                    "y": {"type": "integer", "description": "Y position in 1/100 mm", "default": 1000},
                    "width": {"type": "integer", "description": "Width in 1/100 mm", "default": 16000},
                    "height": {"type": "integer", "description": "Height in 1/100 mm", "default": 10000},
                    "title": {"type": "string", "description": "Chart title"},
                    "categories": {"type": "array", "items": {"type": "string"}, "description": "Category labels (one per data point)"},
                    "series": {
                        "type": "object",
                        "additionalProperties": {"type": "array", "items": {"type": "number"}},
                        "description": "Map of series name -> values, one value per category"
                    },
                    "legend": {"type": "boolean", "description": "Show/hide the legend"}
                }
            },
            "handler": self.insert_chart_live
        }

        self.tools["set_chart_data_live"] = {
            "description": "Update an existing chart's title, legend, and/or data (use list_slides_live to find shape_index)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the chart shape on the slide (0-based)", "default": 0},
                    "title": {"type": "string", "description": "Chart title"},
                    "categories": {"type": "array", "items": {"type": "string"}, "description": "Category labels (one per data point); replaces existing categories"},
                    "series": {
                        "type": "object",
                        "additionalProperties": {"type": "array", "items": {"type": "number"}},
                        "description": "Map of series name -> values, one value per category; replaces existing data"
                    },
                    "legend": {"type": "boolean", "description": "Show/hide the legend"}
                }
            },
            "handler": self.set_chart_data_live
        }

        self.tools["get_chart_data_live"] = {
            "description": "Read back an existing chart's title, legend state, and data (use list_slides_live to find shape_index)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "Index of the slide (0-based)", "default": 0},
                    "shape_index": {"type": "integer", "description": "Index of the chart shape on the slide (0-based)", "default": 0}
                }
            },
            "handler": self.get_chart_data_live
        }

        logger.info(f"Registered {len(self.tools)} MCP tools")
    
    async def execute_tool(self, tool_name: str, parameters: Dict[str, Any]) -> Dict[str, Any]:
        """
        Execute an MCP tool
        
        Args:
            tool_name: Name of the tool to execute
            parameters: Parameters for the tool
            
        Returns:
            Result dictionary
        """
        try:
            if tool_name not in self.tools:
                return {
                    "success": False,
                    "error": f"Unknown tool: {tool_name}",
                    "available_tools": list(self.tools.keys())
                }
            
            tool = self.tools[tool_name]
            handler = tool["handler"]
            
            # Execute the tool handler
            result = handler(**parameters)
            
            logger.info(f"Executed tool '{tool_name}' successfully")
            return result
            
        except Exception as e:
            logger.error(f"Error executing tool '{tool_name}': {e}")
            return {
                "success": False,
                "error": str(e),
                "tool": tool_name,
                "parameters": parameters
            }
    
    def get_tool_list(self) -> List[Dict[str, Any]]:
        """Get list of available tools with their descriptions"""
        return [
            {
                "name": name,
                "description": tool["description"],
                "parameters": tool["parameters"]
            }
            for name, tool in self.tools.items()
        ]
    
    # Tool handler methods
    
    def create_document_live(self, doc_type: str = "writer") -> Dict[str, Any]:
        """Create a new document in LibreOffice"""
        try:
            doc = self.uno_bridge.create_document(doc_type)
            doc_info = self.uno_bridge.get_document_info(doc)
            
            return {
                "success": True,
                "message": f"Created new {doc_type} document",
                "document_info": doc_info
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def insert_text_live(self, text: str, position: Optional[int] = None) -> Dict[str, Any]:
        """Insert text into the currently active document"""
        return self.uno_bridge.insert_text(text, position)
    
    def get_document_info_live(self) -> Dict[str, Any]:
        """Get information about the currently active document"""
        doc_info = self.uno_bridge.get_document_info()
        if "error" in doc_info:
            return {"success": False, **doc_info}
        else:
            return {"success": True, "document_info": doc_info}
    
    def format_text_live(self, **formatting) -> Dict[str, Any]:
        """Apply formatting to selected text"""
        return self.uno_bridge.format_text(formatting)
    
    def save_document_live(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Save the currently active document"""
        return self.uno_bridge.save_document(file_path=file_path)
    
    def export_document_live(self, export_format: str, file_path: str) -> Dict[str, Any]:
        """Export the currently active document"""
        return self.uno_bridge.export_document(export_format, file_path)
    
    def get_text_content_live(self) -> Dict[str, Any]:
        """Get text content of the currently active document"""
        return self.uno_bridge.get_text_content()
    
    def list_open_documents(self) -> Dict[str, Any]:
        """List all open documents in LibreOffice"""
        try:
            desktop = self.uno_bridge.desktop
            documents = []
            
            # Get all open documents
            frames = desktop.getFrames()
            for i in range(frames.getCount()):
                frame = frames.getByIndex(i)
                controller = frame.getController()
                if controller:
                    doc = controller.getModel()
                    if doc:
                        doc_info = self.uno_bridge.get_document_info(doc)
                        documents.append(doc_info)
            
            return {
                "success": True,
                "documents": documents,
                "count": len(documents)
            }
            
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_slides_live(self) -> Dict[str, Any]:
        """List all slides of the active Impress presentation"""
        return self.uno_bridge.list_slides()

    def add_slide_live(self, layout: str = "title_content", index: Optional[int] = None) -> Dict[str, Any]:
        """Insert a new slide into the active Impress presentation"""
        return self.uno_bridge.add_slide(layout=layout, index=index)

    def set_slide_layout_live(self, layout: str, slide_index: int = 0) -> Dict[str, Any]:
        """Change which placeholders a slide has"""
        return self.uno_bridge.set_slide_layout(slide_index=slide_index, layout=layout)

    def set_slide_title_live(self, text: str, slide_index: int = 0, align: Optional[str] = "center",
                              vertical_align: Optional[str] = None, center_box: bool = False) -> Dict[str, Any]:
        """Set a slide's title text"""
        return self.uno_bridge.set_slide_title(
            slide_index=slide_index, text=text, align=align,
            vertical_align=vertical_align, center_box=center_box)

    def set_slide_text_live(self, text: str, slide_index: int = 0, shape_index: int = 1,
                             align: Optional[str] = None, vertical_align: Optional[str] = None) -> Dict[str, Any]:
        """Set the text of an arbitrary shape on a slide"""
        return self.uno_bridge.set_slide_text(
            slide_index=slide_index, shape_index=shape_index, text=text,
            align=align, vertical_align=vertical_align)

    def remove_slide_live(self, slide_index: int = 0) -> Dict[str, Any]:
        """Remove a slide from the presentation"""
        return self.uno_bridge.remove_slide(slide_index=slide_index)

    def duplicate_slide_live(self, slide_index: int = 0) -> Dict[str, Any]:
        """Duplicate a slide"""
        return self.uno_bridge.duplicate_slide(slide_index=slide_index)

    def set_slide_background_live(self, color: str, slide_index: int = 0, apply_to_all: bool = False) -> Dict[str, Any]:
        """Set a slide's (or all slides') background color"""
        return self.uno_bridge.set_slide_background(color=color, slide_index=slide_index, apply_to_all=apply_to_all)

    def add_shape_live(self, slide_index: int = 0, shape_type: str = "rectangle", x: int = 1000, y: int = 1000,
                        width: int = 5000, height: int = 3000, text: Optional[str] = None,
                        fill_color: Optional[str] = None, line_color: Optional[str] = None,
                        shadow: Optional[bool] = None, shadow_color: Optional[str] = None) -> Dict[str, Any]:
        """Add a shape to a slide"""
        return self.uno_bridge.add_shape(
            slide_index=slide_index, shape_type=shape_type, x=x, y=y, width=width, height=height,
            text=text, fill_color=fill_color, line_color=line_color, shadow=shadow, shadow_color=shadow_color)

    def remove_shape_live(self, slide_index: int = 0, shape_index: int = 0) -> Dict[str, Any]:
        """Remove a shape from a slide"""
        return self.uno_bridge.remove_shape(slide_index=slide_index, shape_index=shape_index)

    def set_shape_style_live(self, slide_index: int = 0, shape_index: int = 0,
                              fill_color: Optional[str] = None, line_color: Optional[str] = None,
                              shadow: Optional[bool] = None, shadow_color: Optional[str] = None) -> Dict[str, Any]:
        """Change an existing shape's fill/line color and/or drop shadow"""
        return self.uno_bridge.set_shape_style(
            slide_index=slide_index, shape_index=shape_index, fill_color=fill_color, line_color=line_color,
            shadow=shadow, shadow_color=shadow_color)

    def set_shape_position_live(self, slide_index: int = 0, shape_index: int = 0, x: Optional[int] = None,
                                 y: Optional[int] = None, width: Optional[int] = None,
                                 height: Optional[int] = None) -> Dict[str, Any]:
        """Move/resize an existing shape"""
        return self.uno_bridge.set_shape_position(
            slide_index=slide_index, shape_index=shape_index, x=x, y=y, width=width, height=height)

    def set_shape_hyperlink_live(self, url: str, slide_index: int = 0, shape_index: int = 0) -> Dict[str, Any]:
        """Set/clear a shape's hyperlink"""
        return self.uno_bridge.set_shape_hyperlink(slide_index=slide_index, shape_index=shape_index, url=url)

    def set_list_formatting_live(self, slide_index: int = 0, shape_index: int = 0, list_type: str = "bullet",
                                  level: Optional[int] = None, bullet_char: Optional[str] = None,
                                  number_format: str = "1", start_at: Optional[int] = None) -> Dict[str, Any]:
        """Turn a shape's paragraphs into a bulleted/numbered/plain list"""
        return self.uno_bridge.set_list_formatting(
            slide_index=slide_index, shape_index=shape_index, list_type=list_type, level=level,
            bullet_char=bullet_char, number_format=number_format, start_at=start_at)

    def insert_image_live(self, image_path: str, slide_index: int = 0, x: Optional[int] = None,
                           y: Optional[int] = None, width: Optional[int] = None,
                           height: Optional[int] = None) -> Dict[str, Any]:
        """Insert an image onto a slide"""
        return self.uno_bridge.insert_image(
            slide_index=slide_index, image_path=image_path, x=x, y=y, width=width, height=height)

    def replace_image_live(self, image_path: str, slide_index: int = 0, shape_index: int = 0,
                            keep_size: bool = True) -> Dict[str, Any]:
        """Swap an existing image shape's picture"""
        return self.uno_bridge.replace_image(
            slide_index=slide_index, shape_index=shape_index, image_path=image_path, keep_size=keep_size)

    def format_slide_text_live(self, slide_index: int = 0, shape_index: int = 0, bold: Optional[bool] = None,
                                italic: Optional[bool] = None, underline: Optional[bool] = None,
                                strikeout: Optional[bool] = None, font_size: Optional[float] = None,
                                font_name: Optional[str] = None, color: Optional[str] = None,
                                highlight_color: Optional[str] = None, shadow: Optional[bool] = None,
                                outline: Optional[bool] = None) -> Dict[str, Any]:
        """Format text in a slide shape"""
        return self.uno_bridge.format_slide_text(
            slide_index=slide_index, shape_index=shape_index, bold=bold, italic=italic, underline=underline,
            strikeout=strikeout, font_size=font_size, font_name=font_name, color=color,
            highlight_color=highlight_color, shadow=shadow, outline=outline)

    def set_slide_footer_live(self, footer_text: Optional[str] = None, show_footer: Optional[bool] = None,
                               show_slide_number: Optional[bool] = None, show_date: Optional[bool] = None,
                               date_text: Optional[str] = None, slide_index: Optional[int] = None,
                               apply_to_all: bool = True) -> Dict[str, Any]:
        """Set footer/date/slide-number settings"""
        return self.uno_bridge.set_slide_footer(
            footer_text=footer_text, show_footer=show_footer, show_slide_number=show_slide_number,
            show_date=show_date, date_text=date_text, slide_index=slide_index, apply_to_all=apply_to_all)

    def insert_table_live(self, slide_index: int = 0, rows: int = 2, cols: int = 2, x: int = 1000, y: int = 1000,
                           width: int = 20000, height: int = 8000,
                           data: Optional[List[List[str]]] = None) -> Dict[str, Any]:
        """Insert a table onto a slide"""
        return self.uno_bridge.insert_table(
            slide_index=slide_index, rows=rows, cols=cols, x=x, y=y, width=width, height=height, data=data)

    def set_table_cell_live(self, text: str, slide_index: int = 0, shape_index: int = 0,
                             row: int = 0, col: int = 0) -> Dict[str, Any]:
        """Set a single table cell's text"""
        return self.uno_bridge.set_table_cell(
            slide_index=slide_index, shape_index=shape_index, row=row, col=col, text=text)

    def get_table_data_live(self, slide_index: int = 0, shape_index: int = 0) -> Dict[str, Any]:
        """Read back a table's full contents"""
        return self.uno_bridge.get_table_data(slide_index=slide_index, shape_index=shape_index)

    def insert_table_row_live(self, slide_index: int = 0, shape_index: int = 0,
                               index: Optional[int] = None) -> Dict[str, Any]:
        """Insert a row into an existing table"""
        return self.uno_bridge.insert_table_row(slide_index=slide_index, shape_index=shape_index, index=index)

    def remove_table_row_live(self, slide_index: int = 0, shape_index: int = 0, index: int = 0) -> Dict[str, Any]:
        """Remove a row from an existing table"""
        return self.uno_bridge.remove_table_row(slide_index=slide_index, shape_index=shape_index, index=index)

    def insert_table_column_live(self, slide_index: int = 0, shape_index: int = 0,
                                  index: Optional[int] = None) -> Dict[str, Any]:
        """Insert a column into an existing table"""
        return self.uno_bridge.insert_table_column(slide_index=slide_index, shape_index=shape_index, index=index)

    def remove_table_column_live(self, slide_index: int = 0, shape_index: int = 0, index: int = 0) -> Dict[str, Any]:
        """Remove a column from an existing table"""
        return self.uno_bridge.remove_table_column(slide_index=slide_index, shape_index=shape_index, index=index)

    def insert_chart_live(self, slide_index: int = 0, chart_type: str = "bar", x: int = 1000, y: int = 1000,
                           width: int = 16000, height: int = 10000, title: Optional[str] = None,
                           categories: Optional[List[str]] = None, series: Optional[Dict[str, List[float]]] = None,
                           legend: Optional[bool] = None) -> Dict[str, Any]:
        """Insert a chart onto a slide"""
        return self.uno_bridge.insert_chart(
            slide_index=slide_index, chart_type=chart_type, x=x, y=y, width=width, height=height,
            title=title, categories=categories, series=series, legend=legend)

    def set_chart_data_live(self, slide_index: int = 0, shape_index: int = 0, title: Optional[str] = None,
                             categories: Optional[List[str]] = None, series: Optional[Dict[str, List[float]]] = None,
                             legend: Optional[bool] = None) -> Dict[str, Any]:
        """Update an existing chart's title/legend/data"""
        return self.uno_bridge.set_chart_data(
            slide_index=slide_index, shape_index=shape_index, title=title,
            categories=categories, series=series, legend=legend)

    def get_chart_data_live(self, slide_index: int = 0, shape_index: int = 0) -> Dict[str, Any]:
        """Read back an existing chart's title/legend/data"""
        return self.uno_bridge.get_chart_data(slide_index=slide_index, shape_index=shape_index)


# Global instance
mcp_server = None

def get_mcp_server() -> LibreOfficeMCPServer:
    """Get or create the global MCP server instance"""
    global mcp_server
    if mcp_server is None:
        mcp_server = LibreOfficeMCPServer()
    return mcp_server
