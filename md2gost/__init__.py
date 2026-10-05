"""md2gost — Markdown в DOCX по ГОСТ 7.32-2017"""

from .converter import DocumentSettings, MarkdownToDocxConverter

__version__ = "1.0.0"
__all__ = ["DocumentSettings", "MarkdownToDocxConverter", "__version__"]
