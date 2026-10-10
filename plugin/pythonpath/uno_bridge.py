"""
LibreOffice MCP Extension - UNO Bridge Module

This module provides a bridge between MCP operations and LibreOffice UNO API,
enabling direct manipulation of LibreOffice documents.
"""

import os
import uno
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.document import XDocumentEventListener
from com.sun.star.awt import XActionListener, Point, Size
from com.sun.star.style.ParagraphAdjust import (
    LEFT as PARA_ADJUST_LEFT,
    CENTER as PARA_ADJUST_CENTER,
    RIGHT as PARA_ADJUST_RIGHT,
    BLOCK as PARA_ADJUST_BLOCK,
)
from com.sun.star.drawing.TextVerticalAdjust import (
    TOP as VERT_ADJUST_TOP,
    CENTER as VERT_ADJUST_CENTER,
    BOTTOM as VERT_ADJUST_BOTTOM,
)
from com.sun.star.drawing.FillStyle import SOLID as FILL_SOLID, NONE as FILL_NONE
from com.sun.star.drawing.LineStyle import SOLID as LINE_SOLID, NONE as LINE_NONE
from com.sun.star.style.NumberingType import (
    CHAR_SPECIAL as NUM_CHAR_SPECIAL,
    ARABIC as NUM_ARABIC,
    ROMAN_LOWER as NUM_ROMAN_LOWER,
    ROMAN_UPPER as NUM_ROMAN_UPPER,
    CHARS_LOWER_LETTER as NUM_CHARS_LOWER_LETTER,
    CHARS_UPPER_LETTER as NUM_CHARS_UPPER_LETTER,
)
from typing import Any, Optional, Dict, List
import logging
import traceback

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Friendly names for the legacy AutoLayout codes used by DrawPage.Layout,
# verified empirically against LibreOffice 26.8 (the numbering is undocumented
# and inconsistent across versions, so only the ones actually tested are exposed).
SLIDE_LAYOUTS = {
    "blank": 20,
    "title_only": 19,
    "title_content": 1,
    "title_two_content": 3,
}

PARAGRAPH_ALIGNMENTS = {
    "left": PARA_ADJUST_LEFT,
    "center": PARA_ADJUST_CENTER,
    "right": PARA_ADJUST_RIGHT,
    "justify": PARA_ADJUST_BLOCK,
}

VERTICAL_ALIGNMENTS = {
    "top": VERT_ADJUST_TOP,
    "center": VERT_ADJUST_CENTER,
    "bottom": VERT_ADJUST_BOTTOM,
}

# Shape types that can be created via add_shape_live
SHAPE_TYPES = {
    "rectangle": "com.sun.star.drawing.RectangleShape",
    "ellipse": "com.sun.star.drawing.EllipseShape",
    "line": "com.sun.star.drawing.LineShape",
    "text": "com.sun.star.drawing.TextShape",
}

LIST_TYPES = ("bullet", "number", "none")

# Numbered-list marker style -> com.sun.star.style.NumberingType code, verified against this LibreOffice
NUMBER_FORMATS = {
    "1": NUM_ARABIC,
    "a": NUM_CHARS_LOWER_LETTER,
    "A": NUM_CHARS_UPPER_LETTER,
    "i": NUM_ROMAN_LOWER,
    "I": NUM_ROMAN_UPPER,
}

# Chart type -> legacy com.sun.star.chart diagram service, verified against this LibreOffice
CHART_TYPES = {
    "bar": "com.sun.star.chart.BarDiagram",
    "column": "com.sun.star.chart.BarDiagram",
    "line": "com.sun.star.chart.LineDiagram",
    "pie": "com.sun.star.chart.PieDiagram",
    "area": "com.sun.star.chart.AreaDiagram",
}

# CLSID that tells an OLE2Shape to embed a chart object
CHART_CLSID = "12dcae26-281f-416f-a234-c3086127382e"

# Export filters verified against this LibreOffice's FilterFactory
WRITER_EXPORT_FILTERS = {
    'pdf': 'writer_pdf_Export',
    'docx': 'MS Word 2007 XML',
    'doc': 'MS Word 97',
    'odt': 'writer8',
    'txt': 'Text',
    'rtf': 'Rich Text Format',
    'html': 'HTML (StarWriter)',
}

IMPRESS_EXPORT_FILTERS = {
    'pdf': 'impress_pdf_Export',
    'pptx': 'Impress MS PowerPoint 2007 XML',
    'ppt': 'MS PowerPoint 97',
    'odp': 'impress8',
    'png': 'impress_png_Export',
    'html': 'impress_html_Export',
}


class UNOBridge:
    """Bridge between MCP operations and LibreOffice UNO API"""
    
    def __init__(self):
        """Initialize the UNO bridge"""
        try:
            self.ctx = uno.getComponentContext()
            self.smgr = self.ctx.ServiceManager
            self.desktop = self.smgr.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)
            logger.info("UNO Bridge initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize UNO Bridge: {e}")
            raise
    
    def create_document(self, doc_type: str = "writer") -> Any:
        """
        Create new document using UNO API
        
        Args:
            doc_type: Type of document ('writer', 'calc', 'impress', 'draw')
            
        Returns:
            Document object
        """
        try:
            url_map = {
                "writer": "private:factory/swriter",
                "calc": "private:factory/scalc", 
                "impress": "private:factory/simpress",
                "draw": "private:factory/sdraw"
            }
            
            url = url_map.get(doc_type, "private:factory/swriter")
            doc = self.desktop.loadComponentFromURL(url, "_blank", 0, ())
            logger.info(f"Created new {doc_type} document")
            return doc
            
        except Exception as e:
            logger.error(f"Failed to create document: {e}")
            raise
    
    def get_active_document(self) -> Optional[Any]:
        """Get currently active document.

        getCurrentComponent() relies on window-activation events that may never
        fire when there's no window manager dispatching focus (e.g. LibreOffice
        started headlessly/via script without any UI interaction) - fall back to
        the last open frame's document in that case.
        """
        try:
            doc = self.desktop.getCurrentComponent()
            if doc is None:
                frames = self.desktop.getFrames()
                if frames.getCount() > 0:
                    controller = frames.getByIndex(frames.getCount() - 1).getController()
                    doc = controller.getModel() if controller else None
            if doc:
                logger.info("Retrieved active document")
            return doc
        except Exception as e:
            logger.error(f"Failed to get active document: {e}")
            return None
    
    def get_document_info(self, doc: Any = None) -> Dict[str, Any]:
        """Get information about a document"""
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"error": "No document available"}
            
            info = {
                "title": getattr(doc, 'Title', 'Unknown') if hasattr(doc, 'Title') else "Unknown",
                "url": doc.getURL() if hasattr(doc, 'getURL') else "",
                "modified": doc.isModified() if hasattr(doc, 'isModified') else False,
                "type": self._get_document_type(doc),
                "has_selection": self._has_selection(doc)
            }
            
            # Add document-specific information
            if self._is_writer_document(doc):
                text = doc.getText()
                info["word_count"] = len(text.getString().split())
                info["character_count"] = len(text.getString())
            elif self._is_calc_document(doc):
                sheets = doc.getSheets()
                info["sheet_count"] = sheets.getCount()
                info["sheet_names"] = [sheets.getByIndex(i).getName() 
                                     for i in range(sheets.getCount())]
            
            return info
            
        except Exception as e:
            logger.error(f"Failed to get document info: {e}")
            return {"error": str(e)}
    
    def insert_text(self, text: str, position: Optional[int] = None, doc: Any = None) -> Dict[str, Any]:
        """
        Insert text into a document
        
        Args:
            text: Text to insert
            position: Position to insert at (None for current cursor position)
            doc: Document to insert into (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No active document"}
            
            # Handle Writer documents
            if self._is_writer_document(doc):
                text_obj = doc.getText()
                
                if position is None:
                    # Insert at current cursor position
                    cursor = doc.getCurrentController().getViewCursor()
                else:
                    # Insert at specific position
                    cursor = text_obj.createTextCursor()
                    cursor.gotoStart(False)
                    cursor.goRight(position, False)
                
                text_obj.insertString(cursor, text, False)
                logger.info(f"Inserted {len(text)} characters into Writer document")
                return {"success": True, "message": f"Inserted {len(text)} characters"}
            
            # Handle other document types
            else:
                return {"success": False, "error": f"Text insertion not supported for {self._get_document_type(doc)}"}
                
        except Exception as e:
            logger.error(f"Failed to insert text: {e}")
            return {"success": False, "error": str(e)}
    
    def format_text(self, formatting: Dict[str, Any], doc: Any = None) -> Dict[str, Any]:
        """
        Apply formatting to selected text
        
        Args:
            formatting: Dictionary of formatting options
            doc: Document to format (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc or not self._is_writer_document(doc):
                return {"success": False, "error": "No Writer document available"}
            
            # Get current selection
            selection = doc.getCurrentController().getSelection()
            if selection.getCount() == 0:
                return {"success": False, "error": "No text selected"}
            
            # Apply formatting to selection
            text_range = selection.getByIndex(0)
            
            # Apply various formatting options
            if "bold" in formatting:
                text_range.CharWeight = 150.0 if formatting["bold"] else 100.0
            
            if "italic" in formatting:
                text_range.CharPosture = 2 if formatting["italic"] else 0
            
            if "underline" in formatting:
                text_range.CharUnderline = 1 if formatting["underline"] else 0
            
            if "font_size" in formatting:
                text_range.CharHeight = formatting["font_size"]
            
            if "font_name" in formatting:
                text_range.CharFontName = formatting["font_name"]
            
            logger.info("Applied formatting to selected text")
            return {"success": True, "message": "Formatting applied successfully"}
            
        except Exception as e:
            logger.error(f"Failed to format text: {e}")
            return {"success": False, "error": str(e)}
    
    def save_document(self, doc: Any = None, file_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Save a document
        
        Args:
            doc: Document to save (None for active document)
            file_path: Path to save to (None to save to current location)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document to save"}
            
            if file_path:
                # Save as new file
                url = uno.systemPathToFileUrl(file_path)
                doc.storeAsURL(url, ())
                logger.info(f"Saved document to {file_path}")
                return {"success": True, "message": f"Document saved to {file_path}"}
            else:
                # Save to current location
                if doc.hasLocation():
                    doc.store()
                    logger.info("Saved document to current location")
                    return {"success": True, "message": "Document saved"}
                else:
                    return {"success": False, "error": "Document has no location, specify file_path"}
                    
        except Exception as e:
            logger.error(f"Failed to save document: {e}")
            return {"success": False, "error": str(e)}
    
    def export_document(self, export_format: str, file_path: str, doc: Any = None) -> Dict[str, Any]:
        """
        Export document to different format
        
        Args:
            export_format: Target format ('pdf', 'docx', 'odt', 'txt', etc.)
            file_path: Path to export to
            doc: Document to export (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document to export"}
            
            # Pick the filter table for this document's type
            filter_map = IMPRESS_EXPORT_FILTERS if self._is_impress_document(doc) else WRITER_EXPORT_FILTERS

            filter_name = filter_map.get(export_format.lower())
            if not filter_name:
                return {"success": False, "error": f"Unsupported export format: {export_format}"}
            
            # Prepare export properties
            properties = (
                PropertyValue("FilterName", 0, filter_name, 0),
                PropertyValue("Overwrite", 0, True, 0),
            )
            
            # Export document
            url = uno.systemPathToFileUrl(file_path)
            doc.storeToURL(url, properties)
            
            logger.info(f"Exported document to {file_path} as {export_format}")
            return {"success": True, "message": f"Document exported to {file_path}"}
            
        except Exception as e:
            logger.error(f"Failed to export document: {e}")
            return {"success": False, "error": str(e)}
    
    def get_text_content(self, doc: Any = None) -> Dict[str, Any]:
        """Get text content from a document"""
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document available"}
            
            if self._is_writer_document(doc):
                text = doc.getText().getString()
                return {"success": True, "content": text, "length": len(text)}
            else:
                return {"success": False, "error": f"Text extraction not supported for {self._get_document_type(doc)}"}
                
        except Exception as e:
            logger.error(f"Failed to get text content: {e}")
            return {"success": False, "error": str(e)}
    
    def _is_writer_document(self, doc: Any) -> bool:
        return hasattr(doc, 'supportsService') and doc.supportsService("com.sun.star.text.TextDocument")

    def _is_calc_document(self, doc: Any) -> bool:
        return hasattr(doc, 'supportsService') and doc.supportsService("com.sun.star.sheet.SpreadsheetDocument")

    def _is_impress_document(self, doc: Any) -> bool:
        return hasattr(doc, 'supportsService') and doc.supportsService("com.sun.star.presentation.PresentationDocument")

    def _get_document_type(self, doc: Any) -> str:
        """Determine document type"""
        if self._is_writer_document(doc):
            return "writer"
        elif self._is_calc_document(doc):
            return "calc"
        elif self._is_impress_document(doc):
            return "impress"
        else:
            return "unknown"
    
    def _has_selection(self, doc: Any) -> bool:
        """Check if document has selected content"""
        try:
            if hasattr(doc, 'getCurrentController'):
                controller = doc.getCurrentController()
                if hasattr(controller, 'getSelection'):
                    selection = controller.getSelection()
                    return selection.getCount() > 0
        except:
            pass
        return False

    # --- Impress slide operations -----------------------------------------

    def _get_impress_document(self, doc: Any = None) -> Optional[Any]:
        """Resolve doc to the active document and verify it's an Impress presentation"""
        if doc is None:
            doc = self.get_active_document()
        if not doc or not self._is_impress_document(doc):
            return None
        return doc

    def _get_slide(self, doc: Any, slide_index: int) -> Any:
        """Get a slide (DrawPage) by index, raising IndexError with a clear message"""
        pages = doc.getDrawPages()
        if slide_index < 0 or slide_index >= pages.getCount():
            raise IndexError(
                f"Slide index {slide_index} out of range (document has {pages.getCount()} slide(s))")
        return pages.getByIndex(slide_index)

    def _describe_shape(self, shape: Any, index: int) -> Dict[str, Any]:
        """Summarize a slide shape: its placeholder role, text, position and size"""
        if shape.supportsService("com.sun.star.presentation.TitleTextShape"):
            role = "title"
        elif shape.supportsService("com.sun.star.presentation.OutlinerShape"):
            role = "content"
        elif shape.supportsService("com.sun.star.presentation.SubtitleTextShape"):
            role = "subtitle"
        elif shape.supportsService("com.sun.star.presentation.NotesTextShape"):
            role = "notes"
        elif shape.supportsService("com.sun.star.drawing.OLE2Shape"):
            role = "chart"
        elif shape.supportsService("com.sun.star.drawing.GraphicObjectShape"):
            role = "image"
        elif hasattr(shape, 'Model'):
            role = "table"
        else:
            role = "other"
        try:
            text = shape.getString()
        except Exception:
            text = ""
        result = {
            "shape_index": index,
            "role": role,
            "text": text,
            "position": {"x": shape.Position.X, "y": shape.Position.Y},
            "size": {"width": shape.Size.Width, "height": shape.Size.Height},
        }
        if role == "table":
            try:
                model = shape.Model
                result["rows"] = model.RowCount
                result["cols"] = model.ColumnCount
            except Exception:
                pass
        try:
            if shape.Hyperlink:
                result["hyperlink"] = shape.Hyperlink
        except Exception:
            pass
        return result

    def _describe_slide(self, page: Any, index: int) -> Dict[str, Any]:
        return {
            "slide_index": index,
            "layout": page.Layout,
            "shape_count": page.getCount(),
            "shapes": [self._describe_shape(page.getByIndex(i), i) for i in range(page.getCount())],
        }

    def _find_shape_by_role(self, page: Any, service_name: str) -> Optional[Any]:
        for i in range(page.getCount()):
            shape = page.getByIndex(i)
            if shape.supportsService(service_name):
                return shape
        return None

    def _apply_text_to_shape(self, shape: Any, page: Any, text: str, align: Optional[str],
                              vertical_align: Optional[str], center_box: bool) -> None:
        shape.setString(text)

        if align is not None:
            adjust = PARAGRAPH_ALIGNMENTS.get(align)
            if adjust is None:
                raise ValueError(f"Unknown align '{align}'. Valid values: {list(PARAGRAPH_ALIGNMENTS)}")
            paragraphs = shape.getText().createEnumeration()
            while paragraphs.hasMoreElements():
                paragraphs.nextElement().ParaAdjust = adjust

        if vertical_align is not None:
            vadjust = VERTICAL_ALIGNMENTS.get(vertical_align)
            if vadjust is None:
                raise ValueError(f"Unknown vertical_align '{vertical_align}'. Valid values: {list(VERTICAL_ALIGNMENTS)}")
            shape.TextVerticalAdjust = vadjust

        if center_box:
            box_w, box_h = shape.Size.Width, shape.Size.Height
            shape.Position = Point(int((page.Width - box_w) / 2), int((page.Height - box_h) / 2))

    def list_slides(self, doc: Any = None) -> Dict[str, Any]:
        """List all slides of the active Impress document with their shapes"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        pages = impress_doc.getDrawPages()
        slides = [self._describe_slide(pages.getByIndex(i), i) for i in range(pages.getCount())]
        return {"success": True, "slide_count": pages.getCount(), "slides": slides}

    def add_slide(self, layout: str = "title_content", index: Optional[int] = None, doc: Any = None) -> Dict[str, Any]:
        """Insert a new slide with the given layout"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        layout_code = SLIDE_LAYOUTS.get(layout)
        if layout_code is None:
            return {"success": False, "error": f"Unknown layout '{layout}'. Valid values: {list(SLIDE_LAYOUTS)}"}

        pages = impress_doc.getDrawPages()
        insert_at = pages.getCount() if index is None else index
        pages.insertNewByIndex(insert_at)
        page = pages.getByIndex(insert_at)
        page.Layout = layout_code

        logger.info(f"Inserted slide at index {insert_at} with layout '{layout}'")
        return {"success": True, **self._describe_slide(page, insert_at)}

    def set_slide_layout(self, slide_index: int = 0, layout: str = "title_only", doc: Any = None) -> Dict[str, Any]:
        """Change a slide's layout (which placeholders it has)"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        layout_code = SLIDE_LAYOUTS.get(layout)
        if layout_code is None:
            return {"success": False, "error": f"Unknown layout '{layout}'. Valid values: {list(SLIDE_LAYOUTS)}"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        page.Layout = layout_code
        logger.info(f"Set slide {slide_index} layout to '{layout}'")
        return {"success": True, **self._describe_slide(page, slide_index)}

    def set_slide_title(self, slide_index: int = 0, text: str = "", align: Optional[str] = "center",
                         vertical_align: Optional[str] = None, center_box: bool = False,
                         doc: Any = None) -> Dict[str, Any]:
        """Set a slide's title text, with optional alignment and centering on the slide"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        title_shape = self._find_shape_by_role(page, "com.sun.star.presentation.TitleTextShape")
        if title_shape is None:
            return {
                "success": False,
                "error": (f"Slide {slide_index} has no title placeholder; "
                          "call set_slide_layout_live first with a layout that includes a title "
                          "(e.g. 'title_only' or 'title_content')")
            }

        try:
            self._apply_text_to_shape(title_shape, page, text, align, vertical_align, center_box)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        logger.info(f"Set slide {slide_index} title to {text!r}")
        return {"success": True, "slide_index": slide_index, "text": text}

    def set_slide_text(self, slide_index: int = 0, shape_index: int = 1, text: str = "",
                        align: Optional[str] = None, vertical_align: Optional[str] = None,
                        doc: Any = None) -> Dict[str, Any]:
        """Set the text of an arbitrary shape on a slide (e.g. a content/outline placeholder)"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        if shape_index < 0 or shape_index >= page.getCount():
            return {"success": False,
                    "error": f"Shape index {shape_index} out of range (slide {slide_index} has {page.getCount()} shape(s))"}

        shape = page.getByIndex(shape_index)
        try:
            self._apply_text_to_shape(shape, page, text, align, vertical_align, False)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        logger.info(f"Set slide {slide_index} shape {shape_index} text to {text!r}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index, "text": text}

    def _parse_color(self, color: Any) -> int:
        """Parse '#RRGGBB', 'RRGGBB', or an int into a 0xRRGGBB int"""
        if isinstance(color, int):
            return color
        s = str(color).strip().lstrip('#')
        try:
            return int(s, 16)
        except ValueError:
            raise ValueError(f"Invalid color '{color}'; expected hex like '#3366CC' or '3366CC'")

    def _apply_fill_color(self, shape: Any, color: str) -> None:
        if str(color).lower() == "none":
            shape.FillStyle = FILL_NONE
            return
        shape.FillStyle = FILL_SOLID
        shape.FillColor = self._parse_color(color)

    def _apply_line_color(self, shape: Any, color: str) -> None:
        if str(color).lower() == "none":
            shape.LineStyle = LINE_NONE
            return
        shape.LineStyle = LINE_SOLID
        shape.LineColor = self._parse_color(color)

    def set_slide_background(self, color: str, slide_index: int = 0, apply_to_all: bool = False,
                              doc: Any = None) -> Dict[str, Any]:
        """Set the solid background fill color of one slide, or all slides"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            rgb = self._parse_color(color)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        pages = impress_doc.getDrawPages()
        if apply_to_all:
            indices = list(range(pages.getCount()))
        else:
            if slide_index < 0 or slide_index >= pages.getCount():
                return {"success": False,
                        "error": f"Slide index {slide_index} out of range (document has {pages.getCount()} slide(s))"}
            indices = [slide_index]

        for i in indices:
            page = pages.getByIndex(i)
            background = impress_doc.createInstance("com.sun.star.drawing.Background")
            background.FillStyle = FILL_SOLID
            background.FillColor = rgb
            page.Background = background

        logger.info(f"Set background color {color} on slide(s) {indices}")
        return {"success": True, "slides_updated": indices, "color": color}

    def _apply_shadow(self, shape: Any, shadow: bool, shadow_color: Optional[str]) -> None:
        shape.Shadow = shadow
        if shadow and shadow_color is not None:
            shape.ShadowColor = self._parse_color(shadow_color)

    def add_shape(self, slide_index: int = 0, shape_type: str = "rectangle", x: int = 1000, y: int = 1000,
                  width: int = 5000, height: int = 3000, text: Optional[str] = None,
                  fill_color: Optional[str] = None, line_color: Optional[str] = None,
                  shadow: Optional[bool] = None, shadow_color: Optional[str] = None,
                  doc: Any = None) -> Dict[str, Any]:
        """Add a rectangle/ellipse/line/text shape to a slide"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        service_name = SHAPE_TYPES.get(shape_type)
        if service_name is None:
            return {"success": False, "error": f"Unknown shape_type '{shape_type}'. Valid values: {list(SHAPE_TYPES)}"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        shape = impress_doc.createInstance(service_name)
        page.add(shape)
        shape.Position = Point(int(x), int(y))
        shape.Size = Size(int(width), int(height))

        try:
            if fill_color is not None:
                self._apply_fill_color(shape, fill_color)
            if line_color is not None:
                self._apply_line_color(shape, line_color)
            if shadow is not None:
                self._apply_shadow(shape, shadow, shadow_color)
        except ValueError as e:
            page.remove(shape)
            return {"success": False, "error": str(e)}

        if text and hasattr(shape, 'setString'):
            shape.setString(text)

        shape_index = page.getCount() - 1
        logger.info(f"Added {shape_type} shape to slide {slide_index} at index {shape_index}")
        return {"success": True, "slide_index": slide_index, **self._describe_shape(shape, shape_index)}

    def remove_shape(self, slide_index: int = 0, shape_index: int = 0, doc: Any = None) -> Dict[str, Any]:
        """Remove a shape from a slide (e.g. a temporary marker added for visual verification)"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        if shape_index < 0 or shape_index >= page.getCount():
            return {"success": False,
                    "error": f"Shape index {shape_index} out of range (slide {slide_index} has {page.getCount()} shape(s))"}

        page.remove(page.getByIndex(shape_index))
        logger.info(f"Removed shape {shape_index} from slide {slide_index}")
        return {"success": True, "slide_index": slide_index, "shape_count": page.getCount()}

    def set_shape_style(self, slide_index: int = 0, shape_index: int = 0, fill_color: Optional[str] = None,
                         line_color: Optional[str] = None, shadow: Optional[bool] = None,
                         shadow_color: Optional[str] = None, doc: Any = None) -> Dict[str, Any]:
        """Change the fill color, line color, and/or drop shadow of an existing shape on a slide"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        if shape_index < 0 or shape_index >= page.getCount():
            return {"success": False,
                    "error": f"Shape index {shape_index} out of range (slide {slide_index} has {page.getCount()} shape(s))"}

        shape = page.getByIndex(shape_index)
        try:
            if fill_color is not None:
                self._apply_fill_color(shape, fill_color)
            if line_color is not None:
                self._apply_line_color(shape, line_color)
            if shadow is not None:
                self._apply_shadow(shape, shadow, shadow_color)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        logger.info(f"Set slide {slide_index} shape {shape_index} style fill={fill_color} line={line_color} shadow={shadow}")
        return {"success": True, "slide_index": slide_index, **self._describe_shape(shape, shape_index)}

    def set_shape_position(self, slide_index: int = 0, shape_index: int = 0, x: Optional[int] = None,
                            y: Optional[int] = None, width: Optional[int] = None, height: Optional[int] = None,
                            doc: Any = None) -> Dict[str, Any]:
        """Move and/or resize an existing shape on a slide (position/size in 1/100 mm)"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        if shape_index < 0 or shape_index >= page.getCount():
            return {"success": False,
                    "error": f"Shape index {shape_index} out of range (slide {slide_index} has {page.getCount()} shape(s))"}

        shape = page.getByIndex(shape_index)
        pos, size = shape.Position, shape.Size
        shape.Position = Point(pos.X if x is None else int(x), pos.Y if y is None else int(y))
        shape.Size = Size(size.Width if width is None else int(width), size.Height if height is None else int(height))

        logger.info(f"Repositioned slide {slide_index} shape {shape_index}")
        return {"success": True, "slide_index": slide_index, **self._describe_shape(shape, shape_index)}

    def _load_graphic(self, image_path: str) -> Any:
        """Load a local image file into a UNO Graphic object"""
        graphic_provider = self.smgr.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", self.ctx)
        url_prop = PropertyValue("URL", 0, uno.systemPathToFileUrl(image_path), 0)
        return graphic_provider.queryGraphic((url_prop,))

    def insert_image(self, slide_index: int = 0, image_path: str = "", x: Optional[int] = None,
                      y: Optional[int] = None, width: Optional[int] = None, height: Optional[int] = None,
                      doc: Any = None) -> Dict[str, Any]:
        """Insert an image onto a slide from a local file path; defaults to the image's natural size, centered"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        if not os.path.isfile(image_path):
            return {"success": False, "error": f"Image file not found: {image_path}"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        try:
            graphic = self._load_graphic(image_path)
        except Exception as e:
            return {"success": False, "error": f"Failed to load image: {e}"}

        shape = impress_doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
        page.add(shape)
        shape.Graphic = graphic

        natural = graphic.Size100thMM
        final_w = natural.Width if width is None else int(width)
        final_h = natural.Height if height is None else int(height)
        shape.Size = Size(final_w, final_h)
        shape.Position = Point(
            int((page.Width - final_w) / 2) if x is None else int(x),
            int((page.Height - final_h) / 2) if y is None else int(y))

        shape_index = page.getCount() - 1
        logger.info(f"Inserted image {image_path} on slide {slide_index} at shape index {shape_index}")
        return {"success": True, "slide_index": slide_index, **self._describe_shape(shape, shape_index)}

    def replace_image(self, slide_index: int = 0, shape_index: int = 0, image_path: str = "",
                       keep_size: bool = True, doc: Any = None) -> Dict[str, Any]:
        """Swap the picture of an existing image shape, keeping its position (and by default its size)"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        if not os.path.isfile(image_path):
            return {"success": False, "error": f"Image file not found: {image_path}"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        if shape_index < 0 or shape_index >= page.getCount():
            return {"success": False,
                    "error": f"Shape index {shape_index} out of range (slide {slide_index} has {page.getCount()} shape(s))"}

        shape = page.getByIndex(shape_index)
        if not shape.supportsService("com.sun.star.drawing.GraphicObjectShape"):
            return {"success": False, "error": f"Shape {shape_index} is not an image"}

        try:
            graphic = self._load_graphic(image_path)
        except Exception as e:
            return {"success": False, "error": f"Failed to load image: {e}"}

        previous_size = shape.Size
        shape.Graphic = graphic
        if keep_size:
            shape.Size = previous_size

        logger.info(f"Replaced image on slide {slide_index} shape {shape_index} with {image_path}")
        return {"success": True, "slide_index": slide_index, **self._describe_shape(shape, shape_index)}

    def format_slide_text(self, slide_index: int = 0, shape_index: int = 0, bold: Optional[bool] = None,
                           italic: Optional[bool] = None, underline: Optional[bool] = None,
                           font_size: Optional[float] = None, font_name: Optional[str] = None,
                           color: Optional[str] = None, highlight_color: Optional[str] = None,
                           shadow: Optional[bool] = None, outline: Optional[bool] = None,
                           strikeout: Optional[bool] = None, doc: Any = None) -> Dict[str, Any]:
        """Apply character formatting to all text in a slide shape.

        highlight_color is CharBackColor - what the Impress/Draw UI calls
        "Character Highlighting Color"; Writer's separate CharHighlight property
        doesn't exist on Draw/Impress text, so this is the real equivalent there.
        outline is CharContoured - hollow/contour letters.
        """
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        if shape_index < 0 or shape_index >= page.getCount():
            return {"success": False,
                    "error": f"Shape index {shape_index} out of range (slide {slide_index} has {page.getCount()} shape(s))"}

        try:
            rgb = self._parse_color(color) if color is not None else None
            highlight_rgb = (None if highlight_color is None or highlight_color.lower() == "none"
                              else self._parse_color(highlight_color))
        except ValueError as e:
            return {"success": False, "error": str(e)}

        shape = page.getByIndex(shape_index)
        paragraphs = shape.getText().createEnumeration()
        while paragraphs.hasMoreElements():
            portions = paragraphs.nextElement().createEnumeration()
            while portions.hasMoreElements():
                portion = portions.nextElement()
                if bold is not None:
                    portion.CharWeight = 150.0 if bold else 100.0
                if italic is not None:
                    portion.CharPosture = 2 if italic else 0
                if underline is not None:
                    portion.CharUnderline = 1 if underline else 0
                if font_size is not None:
                    portion.CharHeight = float(font_size)
                if font_name is not None:
                    portion.CharFontName = font_name
                if rgb is not None:
                    portion.CharColor = rgb
                if highlight_color is not None:
                    if highlight_rgb is None:
                        portion.CharBackTransparent = True
                    else:
                        portion.CharBackColor = highlight_rgb
                        portion.CharBackTransparent = False
                if shadow is not None:
                    portion.CharShadowed = shadow
                if outline is not None:
                    portion.CharContoured = outline
                if strikeout is not None:
                    portion.CharStrikeout = 1 if strikeout else 0

        logger.info(f"Formatted text on slide {slide_index} shape {shape_index}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index}

    def remove_slide(self, slide_index: int = 0, doc: Any = None) -> Dict[str, Any]:
        """Remove a slide from the presentation"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        pages = impress_doc.getDrawPages()
        if pages.getCount() <= 1:
            return {"success": False, "error": "Cannot remove the only remaining slide"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        pages.remove(page)
        logger.info(f"Removed slide {slide_index}")
        return {"success": True, "slide_count": pages.getCount()}

    def duplicate_slide(self, slide_index: int = 0, doc: Any = None) -> Dict[str, Any]:
        """Duplicate a slide; the copy is inserted immediately after the original"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        impress_doc.duplicate(page)
        new_index = slide_index + 1
        logger.info(f"Duplicated slide {slide_index} -> new slide at {new_index}")
        pages = impress_doc.getDrawPages()
        return {"success": True, **self._describe_slide(pages.getByIndex(new_index), new_index)}

    def set_slide_footer(self, footer_text: Optional[str] = None, show_footer: Optional[bool] = None,
                          show_slide_number: Optional[bool] = None, show_date: Optional[bool] = None,
                          date_text: Optional[str] = None, slide_index: Optional[int] = None,
                          apply_to_all: bool = True, doc: Any = None) -> Dict[str, Any]:
        """Set footer text, slide number, and date visibility (Impress only supports these per-slide, not a header)"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        pages = impress_doc.getDrawPages()
        if apply_to_all:
            indices = list(range(pages.getCount()))
        else:
            if slide_index is None:
                return {"success": False, "error": "slide_index is required when apply_to_all is False"}
            if slide_index < 0 or slide_index >= pages.getCount():
                return {"success": False,
                        "error": f"Slide index {slide_index} out of range (document has {pages.getCount()} slide(s))"}
            indices = [slide_index]

        for i in indices:
            page = pages.getByIndex(i)
            if footer_text is not None:
                page.FooterText = footer_text
            if show_footer is not None:
                page.IsFooterVisible = show_footer
            if show_slide_number is not None:
                page.IsPageNumberVisible = show_slide_number
            if show_date is not None:
                page.IsDateTimeVisible = show_date
            if date_text is not None:
                page.DateTimeText = date_text
                page.IsDateTimeFixed = True

        logger.info(f"Updated footer/date/page-number settings on slide(s) {indices}")
        return {"success": True, "slides_updated": indices}

    def _get_table_shape(self, page: Any, shape_index: int) -> Any:
        """Look up a shape and verify it's a table; returns the shape or raises ValueError"""
        if shape_index < 0 or shape_index >= page.getCount():
            raise ValueError(f"Shape index {shape_index} out of range (slide has {page.getCount()} shape(s))")
        shape = page.getByIndex(shape_index)
        if shape.supportsService("com.sun.star.drawing.OLE2Shape") or not hasattr(shape, 'Model'):
            raise ValueError(f"Shape {shape_index} is not a table")
        return shape

    def _redistribute_table_dimensions(self, table_shape: Any) -> None:
        """Newly inserted/removed rows and columns leave uneven widths/heights - even them out
        so they sum back to the shape's own size."""
        model = table_shape.Model
        col_width = table_shape.Size.Width // model.ColumnCount
        for i in range(model.ColumnCount):
            model.Columns.getByIndex(i).Width = col_width
        row_height = table_shape.Size.Height // model.RowCount
        for i in range(model.RowCount):
            model.Rows.getByIndex(i).Height = row_height

    def insert_table(self, slide_index: int = 0, rows: int = 2, cols: int = 2, x: int = 1000, y: int = 1000,
                      width: int = 20000, height: int = 8000, data: Optional[List[List[str]]] = None,
                      doc: Any = None) -> Dict[str, Any]:
        """Insert a table onto a slide, optionally pre-filled with cell text (data[row][col])"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        if rows < 1 or cols < 1:
            return {"success": False, "error": "rows and cols must each be at least 1"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        table_shape = impress_doc.createInstance("com.sun.star.presentation.TableShape")
        page.add(table_shape)
        table_shape.Position = Point(int(x), int(y))
        table_shape.Size = Size(int(width), int(height))

        model = table_shape.Model
        if rows > 1:
            model.Rows.insertByIndex(0, rows - 1)
        if cols > 1:
            model.Columns.insertByIndex(0, cols - 1)
        self._redistribute_table_dimensions(table_shape)

        if data:
            for r, row_values in enumerate(data):
                if r >= model.RowCount:
                    break
                for c, value in enumerate(row_values):
                    if c >= model.ColumnCount:
                        break
                    model.getCellByPosition(c, r).setString(str(value))

        shape_index = page.getCount() - 1
        logger.info(f"Inserted {model.RowCount}x{model.ColumnCount} table on slide {slide_index} at shape index {shape_index}")
        return {"success": True, "slide_index": slide_index, "rows": model.RowCount, "cols": model.ColumnCount,
                **self._describe_shape(table_shape, shape_index)}

    def set_table_cell(self, slide_index: int = 0, shape_index: int = 0, row: int = 0, col: int = 0,
                        text: str = "", doc: Any = None) -> Dict[str, Any]:
        """Set the text of a single table cell"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        try:
            shape = self._get_table_shape(page, shape_index)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        model = shape.Model
        if row < 0 or row >= model.RowCount or col < 0 or col >= model.ColumnCount:
            return {"success": False,
                    "error": f"Cell ({row},{col}) out of range (table is {model.RowCount}x{model.ColumnCount})"}

        model.getCellByPosition(col, row).setString(text)
        logger.info(f"Set table cell ({row},{col}) on slide {slide_index} shape {shape_index}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index, "row": row, "col": col, "text": text}

    def get_table_data(self, slide_index: int = 0, shape_index: int = 0, doc: Any = None) -> Dict[str, Any]:
        """Read back a table's full contents"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        try:
            shape = self._get_table_shape(page, shape_index)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        model = shape.Model
        data = [[model.getCellByPosition(c, r).getString() for c in range(model.ColumnCount)]
                for r in range(model.RowCount)]
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index,
                "rows": model.RowCount, "cols": model.ColumnCount, "data": data}

    def insert_table_row(self, slide_index: int = 0, shape_index: int = 0, index: Optional[int] = None,
                          doc: Any = None) -> Dict[str, Any]:
        """Insert a row into an existing table (appends at the end if index is omitted)"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        try:
            shape = self._get_table_shape(page, shape_index)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        model = shape.Model
        insert_at = model.RowCount if index is None else index
        if insert_at < 0 or insert_at > model.RowCount:
            return {"success": False, "error": f"index {insert_at} out of range (table has {model.RowCount} row(s))"}

        model.Rows.insertByIndex(insert_at, 1)
        self._redistribute_table_dimensions(shape)
        logger.info(f"Inserted row at {insert_at} in slide {slide_index} shape {shape_index}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index,
                "rows": model.RowCount, "cols": model.ColumnCount}

    def remove_table_row(self, slide_index: int = 0, shape_index: int = 0, index: int = 0,
                          doc: Any = None) -> Dict[str, Any]:
        """Remove a row from an existing table"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        try:
            shape = self._get_table_shape(page, shape_index)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        model = shape.Model
        if model.RowCount <= 1:
            return {"success": False, "error": "Cannot remove the only remaining row"}
        if index < 0 or index >= model.RowCount:
            return {"success": False, "error": f"index {index} out of range (table has {model.RowCount} row(s))"}

        model.Rows.removeByIndex(index, 1)
        self._redistribute_table_dimensions(shape)
        logger.info(f"Removed row {index} from slide {slide_index} shape {shape_index}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index,
                "rows": model.RowCount, "cols": model.ColumnCount}

    def insert_table_column(self, slide_index: int = 0, shape_index: int = 0, index: Optional[int] = None,
                             doc: Any = None) -> Dict[str, Any]:
        """Insert a column into an existing table (appends at the end if index is omitted)"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        try:
            shape = self._get_table_shape(page, shape_index)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        model = shape.Model
        insert_at = model.ColumnCount if index is None else index
        if insert_at < 0 or insert_at > model.ColumnCount:
            return {"success": False, "error": f"index {insert_at} out of range (table has {model.ColumnCount} column(s))"}

        model.Columns.insertByIndex(insert_at, 1)
        self._redistribute_table_dimensions(shape)
        logger.info(f"Inserted column at {insert_at} in slide {slide_index} shape {shape_index}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index,
                "rows": model.RowCount, "cols": model.ColumnCount}

    def remove_table_column(self, slide_index: int = 0, shape_index: int = 0, index: int = 0,
                             doc: Any = None) -> Dict[str, Any]:
        """Remove a column from an existing table"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        try:
            shape = self._get_table_shape(page, shape_index)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        model = shape.Model
        if model.ColumnCount <= 1:
            return {"success": False, "error": "Cannot remove the only remaining column"}
        if index < 0 or index >= model.ColumnCount:
            return {"success": False, "error": f"index {index} out of range (table has {model.ColumnCount} column(s))"}

        model.Columns.removeByIndex(index, 1)
        self._redistribute_table_dimensions(shape)
        logger.info(f"Removed column {index} from slide {slide_index} shape {shape_index}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index,
                "rows": model.RowCount, "cols": model.ColumnCount}

    def _get_chart_shape(self, page: Any, shape_index: int) -> Any:
        """Look up a shape and verify it's a chart; returns the shape or raises ValueError"""
        if shape_index < 0 or shape_index >= page.getCount():
            raise ValueError(f"Shape index {shape_index} out of range (slide has {page.getCount()} shape(s))")
        shape = page.getByIndex(shape_index)
        if not shape.supportsService("com.sun.star.drawing.OLE2Shape"):
            raise ValueError(f"Shape {shape_index} is not a chart")
        return shape

    def _apply_chart_data(self, chart_doc: Any, categories: Optional[List[str]],
                           series: Optional[Dict[str, List[float]]], title: Optional[str],
                           legend: Optional[bool]) -> None:
        if title is not None:
            chart_doc.getTitle().String = title
        if legend is not None:
            chart_doc.setPropertyValue("HasLegend", legend)
        if categories is not None or series is not None:
            chart_data = chart_doc.getData()
            if categories is not None:
                chart_data.setRowDescriptions(tuple(categories))
            if series is not None:
                chart_data.setColumnDescriptions(tuple(series.keys()))
                values_lists = list(series.values())
                num_rows = len(categories) if categories is not None else (len(values_lists[0]) if values_lists else 0)
                rows_data = tuple(
                    tuple(float(vals[i]) if i < len(vals) else 0.0 for vals in values_lists)
                    for i in range(num_rows)
                )
                chart_data.setData(rows_data)

    def insert_chart(self, slide_index: int = 0, chart_type: str = "bar", x: int = 1000, y: int = 1000,
                      width: int = 16000, height: int = 10000, title: Optional[str] = None,
                      categories: Optional[List[str]] = None, series: Optional[Dict[str, List[float]]] = None,
                      legend: Optional[bool] = None, doc: Any = None) -> Dict[str, Any]:
        """Insert a chart onto a slide, optionally with data (categories[i] paired with series[name][i])"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        diagram_service = CHART_TYPES.get(chart_type)
        if diagram_service is None:
            return {"success": False, "error": f"Unknown chart_type '{chart_type}'. Valid values: {list(CHART_TYPES)}"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        shape = impress_doc.createInstance("com.sun.star.presentation.OLE2Shape")
        page.add(shape)
        shape.Position = Point(int(x), int(y))
        shape.Size = Size(int(width), int(height))
        shape.CLSID = CHART_CLSID

        chart_doc = shape.Model
        chart_doc.setDiagram(chart_doc.createInstance(diagram_service))
        self._apply_chart_data(chart_doc, categories, series, title, legend)

        shape_index = page.getCount() - 1
        logger.info(f"Inserted {chart_type} chart on slide {slide_index} at shape index {shape_index}")
        return {"success": True, "slide_index": slide_index, "chart_type": chart_type,
                **self._describe_shape(shape, shape_index)}

    def set_chart_data(self, slide_index: int = 0, shape_index: int = 0, categories: Optional[List[str]] = None,
                        series: Optional[Dict[str, List[float]]] = None, title: Optional[str] = None,
                        legend: Optional[bool] = None, doc: Any = None) -> Dict[str, Any]:
        """Update an existing chart's title, legend, and/or data"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        try:
            shape = self._get_chart_shape(page, shape_index)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        self._apply_chart_data(shape.Model, categories, series, title, legend)

        logger.info(f"Updated chart data on slide {slide_index} shape {shape_index}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index}

    def get_chart_data(self, slide_index: int = 0, shape_index: int = 0, doc: Any = None) -> Dict[str, Any]:
        """Read back an existing chart's title, legend, and data"""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        try:
            shape = self._get_chart_shape(page, shape_index)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        chart_doc = shape.Model
        chart_data = chart_doc.getData()
        categories = list(chart_data.getRowDescriptions())
        series_names = list(chart_data.getColumnDescriptions())
        raw = chart_data.getData()
        series = {name: [row[i] for row in raw] for i, name in enumerate(series_names)}

        try:
            title = chart_doc.getTitle().String
        except Exception:
            title = ""
        try:
            legend = bool(chart_doc.getPropertyValue("HasLegend"))
        except Exception:
            legend = False

        return {"success": True, "slide_index": slide_index, "shape_index": shape_index,
                "title": title, "categories": categories, "series": series, "legend": legend}

    def set_shape_hyperlink(self, slide_index: int = 0, shape_index: int = 0, url: str = "",
                             doc: Any = None) -> Dict[str, Any]:
        """Set (or clear, with url='') a shape's hyperlink. Impress only supports hyperlinking a whole
        shape (clickable during a slideshow) - there's no per-character/per-run hyperlink like in Writer."""
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        if shape_index < 0 or shape_index >= page.getCount():
            return {"success": False,
                    "error": f"Shape index {shape_index} out of range (slide {slide_index} has {page.getCount()} shape(s))"}

        shape = page.getByIndex(shape_index)
        shape.Hyperlink = url
        logger.info(f"Set hyperlink on slide {slide_index} shape {shape_index} to {url!r}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index, "url": url}

    def _patch_numbering_level(self, impress_doc: Any, patch: Dict[str, Any]) -> Any:
        """Build a standalone NumberingRules object with level 0 patched.

        A paragraph's own NumberingRules is a read-only snapshot - replaceByIndex on it fails
        silently (IllegalArgumentException with no message); you have to build a fresh
        com.sun.star.text.NumberingRules instance and assign the whole thing back to the paragraph.
        replaceByIndex's second argument is typed 'any' in the IDL, so pyuno can't infer the
        sequence type from a plain Python tuple - it must be wrapped via uno.Any(...) explicitly,
        or every call raises the same unhelpful IllegalArgumentException.
        """
        rules = impress_doc.createInstance("com.sun.star.text.NumberingRules")
        level0 = rules.getByIndex(0)
        new_level = []
        for pv in level0:
            if pv.Name in patch:
                new_level.append(PropertyValue(pv.Name, 0, patch[pv.Name], 0))
            else:
                new_level.append(pv)
        rules.replaceByIndex(0, uno.Any("[]com.sun.star.beans.PropertyValue", tuple(new_level)))
        return rules

    def set_list_formatting(self, slide_index: int = 0, shape_index: int = 0, list_type: str = "bullet",
                             level: Optional[int] = None, bullet_char: Optional[str] = None,
                             number_format: str = "1", start_at: Optional[int] = None,
                             doc: Any = None) -> Dict[str, Any]:
        """Turn every paragraph of a slide shape into a bulleted/numbered/plain list.

        bullet_char is only used for list_type='bullet' (omit to keep the master slide's default
        bullet style). number_format/start_at are only used for list_type='number'. level sets the
        outline depth for every paragraph (0 = top level); omit to keep each paragraph's current depth.
        """
        impress_doc = self._get_impress_document(doc)
        if impress_doc is None:
            return {"success": False, "error": "No active Impress document"}

        if list_type not in LIST_TYPES:
            return {"success": False, "error": f"Unknown list_type '{list_type}'. Valid values: {list(LIST_TYPES)}"}

        try:
            page = self._get_slide(impress_doc, slide_index)
        except IndexError as e:
            return {"success": False, "error": str(e)}

        if shape_index < 0 or shape_index >= page.getCount():
            return {"success": False,
                    "error": f"Shape index {shape_index} out of range (slide {slide_index} has {page.getCount()} shape(s))"}

        number_type_code = None
        if list_type == "number":
            number_type_code = NUMBER_FORMATS.get(number_format)
            if number_type_code is None:
                return {"success": False,
                        "error": f"Unknown number_format '{number_format}'. Valid values: {list(NUMBER_FORMATS)}"}

        shape = page.getByIndex(shape_index)
        paragraphs = shape.getText().createEnumeration()
        count = 0
        while paragraphs.hasMoreElements():
            para = paragraphs.nextElement()
            if list_type == "none":
                para.NumberingIsNumber = False
            else:
                para.NumberingIsNumber = True
                if list_type == "bullet" and bullet_char is not None:
                    para.NumberingRules = self._patch_numbering_level(
                        impress_doc, {"NumberingType": NUM_CHAR_SPECIAL, "BulletChar": bullet_char})
                elif list_type == "number":
                    patch = {"NumberingType": number_type_code, "Suffix": ".", "BulletRelSize": 100}
                    if start_at is not None:
                        patch["StartWith"] = int(start_at)
                    para.NumberingRules = self._patch_numbering_level(impress_doc, patch)
            # Re-assigning NumberingLevel - even to its own current value - is what makes the
            # Impress text engine actually apply/render the numbering; without this write the
            # paragraph keeps NumberingIsNumber=True and valid NumberingRules but renders no marker.
            # An untouched paragraph's NumberingLevel getter returns None (UNO void), not 0 - feeding
            # that straight back into the setter raises CannotConvertException, so fall back to 0.
            if level is not None:
                target_level = int(level)
            else:
                current_level = para.NumberingLevel
                target_level = current_level if current_level is not None else 0
            para.NumberingLevel = target_level
            count += 1

        logger.info(f"Set list formatting ({list_type}) on {count} paragraph(s), slide {slide_index} shape {shape_index}")
        return {"success": True, "slide_index": slide_index, "shape_index": shape_index,
                "list_type": list_type, "paragraphs": count}
