# setup.py
from setuptools import setup, find_packages

setup(
    name="markdown_to_word",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        'pywin32',
        'pyyaml',
        'pytest'
    ]
)