# main.py

import os
import logging
from tkinter import Tk, filedialog, messagebox

from src.converter import MarkdownToWordConverter

def get_template_file():
    """Open file chooser dialog for Word template selection"""
    root = Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select Word Template",
        filetypes=[
            ("Word Templates", "*.dotm;*.dotx;*.dot"),
            ("All files", "*.*")
        ],
        initialdir=os.path.expanduser("~/Documents")
    )
    
    return file_path if file_path else None

def get_markdown_file():
    """Open file chooser dialog for markdown file selection"""
    root = Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select Markdown File",
        filetypes=[
            ("Markdown files", "*.md"),
            ("Text files", "*.txt"),
            ("All files", "*.*")
        ],
        initialdir=os.path.expanduser("~/Documents")
    )
    
    return file_path if file_path else None

def main():
    try:
        # Select template file
        template_path = get_template_file()
        if not template_path:
            print("No template selected. Exiting...")
            return

        # Select markdown file
        markdown_path = get_markdown_file()
        if not markdown_path:
            print("No markdown file selected. Exiting...")
            return

        # Generate suggested output filename
        base_name = os.path.splitext(os.path.basename(markdown_path))[0]
        suggested_name = f"{base_name}.docx"
        
        # Get save location
        root = Tk()
        root.withdraw()
        output_path = filedialog.asksaveasfilename(
            title="Save Word Document As",
            filetypes=[("Word Document", "*.docx")],
            initialdir=os.path.expanduser("~/Documents"),
            initialfile=suggested_name,
            defaultextension=".docx"
        )
        
        if not output_path:
            print("No save location selected. Exiting...")
            return
            
        # Get output directory from selected path
        output_dir = os.path.dirname(output_path)
        
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        # Initialize converter
        converter = MarkdownToWordConverter()
        
        print(f"Converting {os.path.basename(markdown_path)}...")
        print("Please wait...")
        
        # Perform conversion
        output_file = converter.convert(template_path, markdown_path, output_dir)
        
        print(f"\nConversion completed successfully!")
        print(f"Output file: {output_file}")
        
        # Ask if user wants to open the file
        if messagebox.askyesno("Success", "Would you like to open the converted file?"):
            os.startfile(output_file)

    except Exception as e:
        error_message = f"Error during conversion: {str(e)}"
        print(f"\n{error_message}")
        messagebox.showerror("Error", error_message)
        logging.error(error_message, exc_info=True)

if __name__ == "__main__":
    main()