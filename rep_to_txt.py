import argparse
import os

IGNORE_PATTERNS = {
    '.git', '.svn', '.hg',  # Version control systems
    '__pycache__', '.pytest_cache',  # Python artifacts
    'node_modules', '.npm',  # Node.js dependencies
    'target', 'build', 'dist',  # Build outputs
    '.idea', '.vscode',  # IDE metadata
    '.DS_Store', 'Thumbs.db',  # OS metadata
    '.pro.user' # QT user config
}

BINARY_EXTENSIONS = {
    '.exe', '.dll', '.so', '.dylib', '.zip', '.tar', '.gz', '.rar', '.7z',
    '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.ico', '.svg', '.webp',
    '.mp3', '.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv',
    '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',
    '.bin', '.dat', '.db', '.sqlite', '.mdb'
}


def scan_directory(path, prefix=""):
    """Рекурсивное сканирование с форматированием дерева"""
    items = []
    try:
        entries = sorted(os.listdir(path))
        # Critical filtering layer для performance optimization
        dirs = [e for e in entries if os.path.isdir(os.path.join(path, e)) and e not in IGNORE_PATTERNS]
        files = [e for e in entries if os.path.isfile(os.path.join(path, e)) and e not in IGNORE_PATTERNS]
        all_items = dirs + files

        for i, item in enumerate(all_items):
            item_path = os.path.join(path, item)
            is_last_item = (i == len(all_items) - 1)

            if is_last_item:
                current_prefix = prefix + "└── "
                next_prefix = prefix + "    "
            else:
                current_prefix = prefix + "├── "
                next_prefix = prefix + "│   "

            items.append(current_prefix + item)

            if os.path.isdir(item_path):
                items.extend(scan_directory(item_path, next_prefix))

    except PermissionError:
        items.append(prefix + "└── [Access Denied]")

    return items


def generate_complete_project_structure(root_path):
    """Генератор проектной документации корпоративного уровня"""
    if not os.path.exists(root_path):
        return f"Error: Path {root_path} does not exist"

    result = []

    # Этап 1: Создание древовидной структуры
    root_name = os.path.basename(root_path) or root_path
    result.append(root_name)
    result.extend(scan_directory(root_path))

    # Этап 2: Полное извлечение содержимого файла
    result.append("\n")  # Separator между разделами дерева и содержимым
    result.extend(extract_all_file_contents(root_path))

    return "\n".join(result)


def extract_all_file_contents(root_path):
    """Механизм извлечения контента с обработкой файлов"""
    content_lines = []

    for root, dirs, files in os.walk(root_path):
        dirs[:] = [d for d in dirs if d not in IGNORE_PATTERNS]

        for file in sorted(files):
            if file in IGNORE_PATTERNS:
                continue

            file_path = os.path.join(root, file)
            relative_path = os.path.relpath(file_path, root_path)

            content_lines.extend(process_single_file(relative_path, file_path))

    return content_lines


def process_single_file(relative_path, file_path):
    """Обработка файлов"""
    content_lines = []

    # Раздел заголовка
    content_lines.append("\n" + "-" * 80)
    content_lines.append(f"{relative_path}:")
    content_lines.append("-" * 80)

    file_ext = os.path.splitext(relative_path)[1].lower()

    # Обнаружение двоичных файлов (вместо изображений раньше выводилась ссылка-заготовка
    # https://raw.githubusercontent.com/.../ — она никуда не вела)
    if file_ext in BINARY_EXTENSIONS or is_likely_binary(file_path):
        if file_ext in {'.png', '.jpg', '.jpeg', '.gif', '.svg', '.ico', '.webp', '.bmp'}:
            content_lines.append("[Image - content not displayed]")
        else:
            content_lines.append("[Binary file - content not displayed]")
    else:
        # Извлечение содержимого текстового файла
        content_lines.extend(extract_text_content(file_path))

    content_lines.append("")
    return content_lines


def extract_text_content(file_path):
    """Резервное извлечение с несколькими кодировками"""
    encodings_priority = ['utf-8', 'utf-8-sig', 'cp1251', 'latin1', 'cp1252']

    for encoding in encodings_priority:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                lines = f.readlines()
                return [f"{i:4} | {line.rstrip()}" for i, line in enumerate(lines, 1)]
        except (UnicodeDecodeError, UnicodeError):
            continue
        except Exception as e:
            return [f"ERROR: Не удается прочитать файл - {e}"]

    return ["WARNING: Кодировка файла, не поддерживаемая для извлечения текста"]


def is_likely_binary(file_path):
    """Эвристическое обнаружение двоичных файлов для крайних случаев"""
    try:
        with open(file_path, 'rb') as f:
            chunk = f.read(8192)
            # Обнаружение нулевого байта - надежный бинарный индикатор
            return b'\x00' in chunk
    except OSError:
        return True


def main(argv=None):
    parser = argparse.ArgumentParser(description="Дерево каталогов проекта и содержимое текстовых файлов в один .txt")
    parser.add_argument("path", help="папка проекта")
    parser.add_argument("-o", "--output", help="куда сохранить (по умолчанию <папка>_rep.txt)")
    args = parser.parse_args(argv)

    project_path = os.path.abspath(args.path)
    output = args.output or os.path.basename(project_path.rstrip(os.sep)) + "_rep.txt"
    with open(output, "w", encoding="utf-8") as f:
        f.write(generate_complete_project_structure(project_path))
    print(output)


if __name__ == "__main__":
    main()
