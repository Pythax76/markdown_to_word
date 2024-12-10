# src/converter.py

import os
import time
import logging
import win32com.client
from win32com.client import constants
from .style_manager import StyleManager
from .formatters import TextFormatter
from .utils import create_unique_filename, setup_logging

class MarkdownToWordConverter:
    def __init__(self):
        setup_logging()
        self.style_manager = StyleManager()
        self.text_formatter = TextFormatter()
        self.word_app = None
        self.doc = None

    def init_word(self):
        """Initialize Microsoft Word application"""
        try:
            # Kill any existing Word processes
            os.system("taskkill /f /im WINWORD.EXE 2>nul")
            time.sleep(1)
            
            # Create new Word instance
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = True
            
            # Wait for Word to be ready
            time.sleep(1)
            
            logging.info("Word application initialized successfully")
            return True
        except Exception as e:
            logging.error(f"Failed to initialize Word: {str(e)}")
            raise RuntimeError(f"Failed to initialize Microsoft Word: {str(e)}")

    def verify_template(self, template_path: str):
        """Verify template file exists and is valid"""
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
        
        if not template_path.lower().endswith(('.dotm', '.dotx', '.dot')):
            raise ValueError("Template file must be a Word template (.dotm, .dotx, or .dot)")
        
        # Check file size
        if os.path.getsize(template_path) == 0:
            raise ValueError("Template file is empty")
        
        logging.info(f"Template verified: {template_path}")
        return True

    def create_document(self, template_path: str):
        """Create new document from template"""
        try:
            logging.info(f"Creating document from template: {template_path}")
            
            # Verify template first
            self.verify_template(template_path)
            
            # Close any open documents
            if self.word_app.Documents.Count > 0:
                for doc in self.word_app.Documents:
                    try:
                        doc.Close(SaveChanges=False)
                    except:
                        pass
            
            try:
                # Try to create blank document first
                blank_doc = self.word_app.Documents.Add()
                if not blank_doc:
                    raise RuntimeError("Failed to create blank document")
                
                # Now try to attach template
                blank_doc.set_AttachedTemplate(template_path)
                self.doc = blank_doc
                
            except Exception as template_error:
                logging.warning(f"Failed to attach template, trying direct creation: {str(template_error)}")
                # If that fails, try direct creation
                self.doc = self.word_app.Documents.Add(Template=template_path)
            
            if not self.doc:
                raise RuntimeError("Document creation failed - no document object")
            
            # Verify document is accessible
            _ = self.doc.Content
            
            logging.info("Document created successfully")
            return True
            
        except Exception as e:
            logging.error(f"Failed to create document: {str(e)}")
            self.cleanup()
            raise RuntimeError(f"Failed to create document: {str(e)}")

    def convert(self, template_path: str, markdown_path: str, output_dir: str) -> str:
        """Convert markdown file to Word document"""
        try:
            # Initialize Word
            self.init_word()
            
            # Normalize paths
            template_path = os.path.abspath(template_path)
            markdown_path = os.path.abspath(markdown_path)
            output_dir = os.path.abspath(output_dir)
            
            logging.info(f"Converting {markdown_path} using template {template_path}")
            
            # Create output directory if needed
            os.makedirs(output_dir, exist_ok=True)
            
            # Create document from template
            self.create_document(template_path)
            
            if not self.doc:
                raise RuntimeError("No document object after creation")
            
            # Generate output filename
            output_path = create_unique_filename(markdown_path, output_dir)
            
            # Process markdown
            self.process_markdown_file(self.doc, markdown_path)
            
            # Save the document
            self.doc.SaveAs(output_path)
            logging.info(f"Document saved successfully to {output_path}")
            
            return output_path

        except Exception as e:
            logging.error(f"Conversion failed: {str(e)}")
            raise
        finally:
            self.cleanup()

    def process_markdown_file(self, doc, markdown_path: str):
        """Process the markdown file"""
        try:
            with open(markdown_path, 'r', encoding='utf-8') as file:
                content = file.read()
                if not content.strip():
                    raise ValueError("Markdown file is empty")
                self.text_formatter.process_content(doc, content, self.style_manager)
        except Exception as e:
            logging.error(f"Failed to process markdown file: {str(e)}")
            raise

    def cleanup(self):
        """Clean up Word resources"""
        try:
            if self.doc:
                try:
                    self.doc.Close(SaveChanges=False)
                except:
                    pass
                self.doc = None
            
            if self.word_app:
                try:
                    self.word_app.Quit()
                except:
                    pass
                self.word_app = None
                
            time.sleep(1)
            
        except Exception as e:
            logging.error(f"Error during cleanup: {str(e)}")