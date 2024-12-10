# tests/test_environment.py

import os
import sys
import win32com.client
import yaml
import logging

def test_environment():
    tests_passed = True
    
    print("Testing Markdown to Word Environment Setup...")
    print("-" * 50)

    # Add project root to Python path
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    if project_root not in sys.path:
        sys.path.insert(0, project_root)

    # 1. Test project structure
    print("\nChecking project structure:")
    required_paths = [
        "src",
        "tests",
        "config",
        "src/__init__.py",
        "src/converter.py",
        "src/style_manager.py",
        "src/formatters.py",
        "src/utils.py",
        "config/logging_config.yaml"
    ]
    
    for path in required_paths:
        exists = os.path.exists(path)
        print(f"{'✓' if exists else '✗'} {path}")
        if not exists:
            tests_passed = False

    # 2. Test imports and class instantiation
    print("\nTesting imports and class instantiation:")
    try:
        from src.style_manager import StyleManager
        from src.formatters import TextFormatter
        from src.converter import MarkdownToWordConverter
        from src.utils import create_unique_filename, setup_logging
        
        # Create test instances
        style_manager = StyleManager()
        text_formatter = TextFormatter()
        converter = MarkdownToWordConverter()
        
        print("✓ All modules imported and initialized successfully")
        
    except Exception as e:
        print(f"✗ Error during import testing: {str(e)}")
        tests_passed = False

    # 3. Test Word COM interface
    print("\nTesting Word COM interface:")
    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        print("✓ Word COM interface working")
        word_app.Quit()
    except Exception as e:
        print(f"✗ Word COM interface failed: {str(e)}")
        tests_passed = False

    # 4. Test logging configuration
    print("\nTesting logging configuration:")
    try:
        with open("config/logging_config.yaml", 'r') as f:
            config = yaml.safe_load(f)
        print("✓ Logging config file is valid YAML")
        
        logging.basicConfig(
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler("markdown_to_word_debug.log"),
            ],
        )
        print("✓ Logging configured successfully")
        logging.debug("Test debug message")
        logging.info("Test info message")
        print("✓ Logging messages written successfully")
        
    except Exception as e:
        print(f"✗ Logging configuration failed: {str(e)}")
        tests_passed = False

    # 5. Test file paths
    print("\nChecking file paths:")
    paths_to_check = {
        "Template": r"C:\Users\jlawrence\OneDrive - Photronics\Documents\TemplateCity\Photronics_Governance_Template_Python.dotm",
        "Output directory": r"C:\Users\jlawrence\OneDrive - Photronics\Documents\IT Securtiy\Policy Management\Output"
    }
    
    for name, path in paths_to_check.items():
        exists = os.path.exists(path)
        print(f"{'✓' if exists else '✗'} {name}: {path}")
        if not exists:
            tests_passed = False

    # Final summary
    print("\n" + "=" * 50)
    if tests_passed:
        print("✓ All environment tests passed!")
    else:
        print("✗ Some tests failed. Please review the output above.")
    
    return tests_passed

if __name__ == "__main__":
    test_environment()