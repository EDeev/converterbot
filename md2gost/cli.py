import argparse
import sys
from pathlib import Path

from . import __version__
from .converter import DocumentSettings, MarkdownToDocxConverter


def build_parser():
    parser = argparse.ArgumentParser(
        prog="md2gost",
        description="Конвертирует Markdown в DOCX, оформленный по ГОСТ 7.32-2017: поля, шрифт, интервалы, "
                    "нумерация страниц и заголовков, таблицы, списки, сноски, список литературы.",
    )
    parser.add_argument("input", type=Path, help="файл .md")
    parser.add_argument("-o", "--output", type=Path, help="куда сохранить .docx (по умолчанию — рядом с .md)")
    parser.add_argument("--font", default="Times New Roman", help="шрифт (по умолчанию Times New Roman)")
    parser.add_argument("--size", type=int, default=14, help="размер основного текста, пт (по умолчанию 14)")
    parser.add_argument("--spacing", type=float, default=1.5, help="межстрочный интервал (по умолчанию 1.5)")
    parser.add_argument("--no-heading-numbers", action="store_true", help="не нумеровать заголовки")
    parser.add_argument("--no-page-numbers", action="store_true", help="не нумеровать страницы")
    parser.add_argument("--number-title-page", action="store_true", help="ставить номер и на первой странице")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    if not args.input.is_file():
        print(f"md2gost: файл не найден: {args.input}", file=sys.stderr)
        return 1

    settings = DocumentSettings()
    settings.font_name = args.font
    settings.font_size = args.size
    settings.line_spacing = args.spacing
    settings.auto_numbering_headings = not args.no_heading_numbers
    settings.page_numbering = not args.no_page_numbers
    settings.exclude_title_page_numbering = not args.number_title_page

    output = MarkdownToDocxConverter(settings).convert(str(args.input), str(args.output) if args.output else None)
    print(output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
