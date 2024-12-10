# src/style_manager.py

import logging
from typing import Dict, Optional

class StyleManager:
    """Manages Word document styles for markdown conversion"""
    
    def __init__(self):
        self.current_level = 0
        self.style_types = {
            'heading': 'Heading',
            'body': 'Body',
            'bullet': 'Bullet',
            'numbered': 'Numbered',
            'blockquote': 'Blockquote',
            'code': 'Code'
        }
        logging.debug("StyleManager initialized")

    def get_style_name(self, element_type: str, level: Optional[int] = None) -> str:
        """
        Get the appropriate style name based on element type and level
        Args:
            element_type: Type of element (heading, body, etc.)
            level: Optional level override (if not provided, uses current_level)
        Returns:
            str: The full style name to be applied
        """
        use_level = level if level is not None else self.current_level
        try:
            base_style = self.style_types[element_type.lower()]
            style_name = f"{base_style} {use_level}"
            logging.debug(f"Generated style name: {style_name} for {element_type} level {use_level}")
            return style_name
        except KeyError:
            fallback = f"Body {use_level}"
            logging.warning(f"Unknown element type {element_type}, falling back to {fallback}")
            return fallback

    def verify_style_exists(self, word_doc, style_name: str) -> bool:
        """
        Check if a style exists in the document
        Args:
            word_doc: Word document object
            style_name: Name of the style to check
        Returns:
            bool: True if style exists, False otherwise
        """
        try:
            _ = word_doc.Styles(style_name)
            return True
        except Exception as e:
            logging.debug(f"Style {style_name} not found: {str(e)}")
            return False

    def apply_style(self, word_doc, paragraph, style_type: str, level: Optional[int] = None) -> None:
        """
        Apply style to paragraph with fallback handling
        Args:
            word_doc: Word document object
            paragraph: Paragraph object to style
            style_type: Type of style to apply
            level: Optional level override
        """
        try:
            # Get the appropriate style name
            style_name = self.get_style_name(style_type, level)
            
            # Try to apply the specific level style
            if self.verify_style_exists(word_doc, style_name):
                paragraph.Range.Style = style_name
                logging.info(f"Applied style: {style_name}")
                return
            
            # Try base style (level 0)
            base_style = f"{self.style_types.get(style_type.lower(), 'Body')} 0"
            if self.verify_style_exists(word_doc, base_style):
                paragraph.Range.Style = base_style
                logging.info(f"Applied base style: {base_style}")
                return
            
            # Final fallback to Normal
            logging.warning(f"Neither {style_name} nor {base_style} found, falling back to Normal")
            paragraph.Range.Style = "Normal"
            
        except Exception as e:
            logging.error(f"Failed to apply style {style_type}: {str(e)}")
            try:
                paragraph.Range.Style = "Normal"
            except:
                logging.error("Failed to apply Normal style as fallback")

    def set_heading_level(self, level: int) -> None:
        """
        Set the current heading level
        Args:
            level: Heading level (1-9)
        """
        self.current_level = min(max(level, 0), 9)  # Ensure level is between 0 and 9
        logging.debug(f"Set heading level to: {self.current_level}")

    def get_current_level(self) -> int:
        """Get current heading level"""
        return self.current_level