# src/formatters.py

import re
import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .style_manager import StyleManager

class TextFormatter:
    """Handles text formatting and markdown processing"""
    
    def __init__(self):
        self.patterns = [
            (r'^(#{1,9})\s+(.+?)(?:\s*>>.*)?$', 'heading'),    # Headers with optional metadata
            (r'^\s*[-*+]\s+(.+)$', 'bullet'),                  # Bullet points
            (r'^\s*\d+\.\s+(.+)$', 'numbered'),                # Numbered lists
            (r'^>\s*(.+)$', 'blockquote'),                     # Blockquotes
            (r'^```.*\n([\s\S]*?)\n```$', 'code'),            # Code blocks
            (r'^(.+)$', 'body')                                # Default body text
        ]

    def add_paragraph(self, doc, text: str, style_type: str, style_manager: 'StyleManager'):
        """Add a new paragraph with proper formatting"""
        try:
            # Get the selection object
            selection = doc.Application.Selection
            
            # Move to the end of the document
            selection.EndKey(Unit=6)  # 6 = wdStory
            
            # Add a new paragraph
            selection.TypeParagraph()
            
            # Get the newly created paragraph
            paragraph = selection.Paragraphs.Item(selection.Paragraphs.Count)
            
            # Set the text
            paragraph.Range.Text = text.strip()
            
            # Let StyleManager handle all style applications
            style_manager.apply_style(doc, paragraph, style_type)
            
            # Apply character formatting for special text (bold, italic, etc.)
            self.apply_character_formatting(paragraph.Range, text)
            
            logging.debug(f"Added paragraph with style {style_type}: {text[:50]}...")
            
        except Exception as e:
            logging.error(f"Failed to add paragraph: {str(e)}")
            raise

    def process_content(self, doc, content: str, style_manager: 'StyleManager'):
        """Process markdown content and apply formatting"""
        try:
            # Clear any existing content
            doc.Content.Delete()
            
            # Split content into lines
            lines = content.split('\n')
            in_code_block = False
            code_block_content = []
            
            for line in lines:
                if line.strip().startswith('```'):
                    if in_code_block:
                        # End code block
                        if code_block_content:
                            self.add_paragraph(doc, '\n'.join(code_block_content), 'code', style_manager)
                        code_block_content = []
                    in_code_block = not in_code_block
                    continue

                if in_code_block:
                    code_block_content.append(line)
                    continue

                if line.strip():  # Only process non-empty lines
                    self.process_line(doc, line, style_manager)

        except Exception as e:
            logging.error(f"Failed to process content: {str(e)}")
            raise

    def process_line(self, doc, line: str, style_manager: 'StyleManager'):
        """Process a single line of markdown"""
        try:
            line = line.strip()
            if not line:
                return

            for pattern, style_type in self.patterns:
                match = re.match(pattern, line)
                if match:
                    if style_type == 'heading':
                        # Update heading level and get text
                        level = len(match.group(1))
                        style_manager.current_level = min(level, 9)
                        text = match.group(2)
                    elif style_type == 'blockquote':
                        # Extract blockquote text without the '>' prefix
                        text = match.group(1).strip()
                    else:
                        text = match.group(1) if len(match.groups()) > 0 else line

                    # Add the paragraph with proper formatting
                    self.add_paragraph(doc, text, style_type, style_manager)
                    break

        except Exception as e:
            logging.error(f"Failed to process line: {str(e)}")
            raise

def apply_character_formatting(self, range_object, text: str):
    """Apply character-level formatting with proper order and nesting"""
    try:
        # Define formatting patterns in order of specificity with more strict matching
        format_patterns = [
            (r'\*{3}([^*]+?)\*{3}', 'bold-italic'),   # Exactly three asterisks for bold-italic
            (r'\*{2}([^*]+?)\*{2}', 'bold'),          # Exactly two asterisks for bold
            (r'\*{1}([^*]+?)\*{1}', 'italic'),        # Exactly one asterisk for italic
            (r'`([^`]+?)`', 'code')                    # Code remains the same
        ]

        # First set the plain text
        clean_text = text
        for pattern, _ in format_patterns:
            clean_text = re.sub(pattern, r'\1', clean_text)
        range_object.Text = clean_text.strip()

        # Then apply formatting with strict pattern matching
        for pattern, format_type in format_patterns:
            matches = list(re.finditer(pattern, text))
            for match in matches:
                try:
                    content = match.group(1)
                    # Find the content in the clean text
                    start_pos = range_object.Text.find(content)
                    if start_pos != -1:
                        # Create a range for this match
                        format_range = range_object.Document.Range(
                            range_object.Start + start_pos,
                            range_object.Start + start_pos + len(content)
                        )

                        # Apply formatting based on exact pattern match
                        if format_type == 'bold-italic' and text.count('*', match.start(), match.end()) == 6:  # Ensure exactly 6 asterisks
                            format_range.Font.Bold = -1
                            format_range.Font.Italic = -1
                            logging.debug(f"Applied bold-italic to: {content}")
                        elif format_type == 'bold' and text.count('*', match.start(), match.end()) == 4:  # Ensure exactly 4 asterisks
                            format_range.Font.Bold = -1
                            logging.debug(f"Applied bold to: {content}")
                        elif format_type == 'italic' and text.count('*', match.start(), match.end()) == 2:  # Ensure exactly 2 asterisks
                            format_range.Font.Italic = -1
                            logging.debug(f"Applied italic to: {content}")
                        elif format_type == 'code':
                            format_range.Font.Name = "Consolas"
                            format_range.Font.Size = 9
                            format_range.Font.Color = 0x505050  # Dark gray
                            format_range.Shading.BackgroundPatternColor = 0xF0F0F0  # Light gray
                            logging.debug(f"Applied code formatting to: {content}")

                except Exception as e:
                    logging.warning(f"Failed to apply {format_type} formatting to '{content}': {str(e)}")

        except Exception as e:
            logging.error(f"Failed to apply character formatting: {str(e)}")
            logging.debug(f"Original text: {text}")
            # Ensure the text is at least visible
            range_object.Text = text.replace('*', '').replace('`', '')