import asyncio
import logging
import os
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from aiogram import Bot, Dispatcher, F, Router
from aiogram.types import Message, BufferedInputFile
from aiogram.filters import Command
from aiogram.enums.parse_mode import ParseMode
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.client.bot import DefaultBotProperties

# Импорт наших конвертеров
from md2gost import MarkdownToDocxConverter, DocumentSettings
from rep_to_txt import generate_complete_project_structure

# Конфигурация
BOT_TOKEN = os.getenv("BOT_TOKEN", "XXXXXXXXXXXXXXXXXXXXXXXX")  # @my_convbot

MAX_FILE_SIZE = 20 * 1024 * 1024        # больше Telegram-боту не скачать
MAX_UNPACKED_SIZE = 200 * 1024 * 1024   # защита от zip-бомбы
MAX_FILES_IN_ARCHIVE = 5000

# Инициализация бота
bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())
router = Router()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@router.message(Command("start"))
async def start_handler(msg: Message) -> None:
    """Приветственное сообщение"""
    await msg.answer(
        "<b>Приветствую в своём Конвертере!</b>\n\n"
        "📋 <b>Возможности:</b>\n"
        "• Отправьте .md файл → получите .docx\n"
        "• Отправьте .zip архив → получите структуру проекта в .txt\n\n"
        "📝 <b>/help</b> - Для подробной информации"
    )

@router.message(Command("help"))
async def help_handler(msg: Message) -> None:
    """Подробная справка"""
    help_text = (
        "<b>📚 Подробное руководство</b>\n\n"
        "<b>1. Конвертация Markdown → DOCX:</b>\n"
        "• Отправьте .md файл\n"
        "• Получите DOCX с форматированием по ГОСТ\n\n"
        "<b>2. Анализ архива → TXT:</b>\n"
        "• Отправьте .zip архив\n"
        "• Получите полную структуру проекта в текстовом файле\n\n"
        "<b>⚡ Ограничения:</b>\n"
        "• Размер файла: до 20 МБ\n"
        "• Поддерживаемые форматы: .md, .zip"
    )
    await msg.answer(help_text)

@router.message(F.document)
async def handle_document(msg: Message) -> None:
    """Обработка загруженных документов"""
    document = msg.document
    # у документа может не быть имени; путь берём только из имени файла, без каталогов
    file_name = Path(document.file_name or "file").name
    file_size = document.file_size or 0

    if file_size > MAX_FILE_SIZE:
        await msg.answer("❌ Файл слишком большой! Максимум 20 МБ")
        return

    file_ext = Path(file_name).suffix.lower()
    if file_ext not in (".md", ".zip"):
        await msg.answer("❌ Неподдерживаемый формат файла! Пришлите .md или .zip")
        return

    status_msg = await msg.answer("⏳ Обрабатываю файл...")
    
    try:
        with TemporaryDirectory() as temp_dir:
            # Скачиваем файл
            file_info = await bot.get_file(document.file_id)
            input_path = os.path.join(temp_dir, file_name)
            await bot.download_file(file_info.file_path, input_path)
            
            # конвертация — в отдельном потоке, чтобы бот не замирал для остальных
            if file_ext == '.md':
                # Конвертация MD → DOCX
                output_path = await asyncio.to_thread(convert_md_to_docx, input_path, temp_dir)
                output_name = Path(file_name).stem + '.docx'
            else:
                # Анализ архива → TXT
                output_path = await asyncio.to_thread(analyze_archive, input_path, temp_dir)
                output_name = Path(file_name).stem + '_structure.txt'
            
            # Отправка результата
            with open(output_path, 'rb') as output_file:
                result_file = BufferedInputFile(
                    output_file.read(),
                    filename=output_name
                )
                await msg.answer_document(result_file)
            
            await status_msg.edit_text("✅ Конвертация завершена!")
            
    except ArchiveTooLarge as e:
        await status_msg.edit_text(f"❌ {e}")
    except zipfile.BadZipFile:
        await status_msg.edit_text("❌ Архив повреждён или это не .zip")
    except Exception as e:
        logger.exception("Ошибка обработки файла")
        await status_msg.edit_text(f"❌ Ошибка обработки: {str(e)}")


class ArchiveTooLarge(Exception):
    pass


def convert_md_to_docx(md_path: str, temp_dir: str) -> str:
    """Конвертация Markdown в DOCX"""
    output_path = os.path.join(temp_dir, "output.docx")
    
    # Настройки
    settings = DocumentSettings()
    settings.font_name = "Times New Roman"
    settings.font_size = 14
    settings.line_spacing = 1.5
    settings.margin_left = 3.0
    settings.auto_numbering_headings = True
    
    converter = MarkdownToDocxConverter(settings)
    converter.convert(md_path, output_path)
    
    return output_path

def check_archive(zip_ref: zipfile.ZipFile) -> None:
    """Архив на 20 МБ может распаковаться в гигабайты и забить диск — проверяем до распаковки"""
    infos = zip_ref.infolist()
    if len(infos) > MAX_FILES_IN_ARCHIVE:
        raise ArchiveTooLarge(f"В архиве больше {MAX_FILES_IN_ARCHIVE} файлов")
    if sum(info.file_size for info in infos) > MAX_UNPACKED_SIZE:
        raise ArchiveTooLarge(f"Распакованный архив больше {MAX_UNPACKED_SIZE // 1024 // 1024} МБ")


def analyze_archive(archive_path: str, temp_dir: str) -> str:
    """Анализ архива и создание структуры проекта"""
    extract_dir = os.path.join(temp_dir, "extracted")
    os.makedirs(extract_dir, exist_ok=True)

    # Извлечение архива
    with zipfile.ZipFile(archive_path, 'r') as zip_ref:
        check_archive(zip_ref)
        zip_ref.extractall(extract_dir)
    
    # Поиск основной папки проекта
    extracted_items = os.listdir(extract_dir)
    if len(extracted_items) == 1 and os.path.isdir(os.path.join(extract_dir, extracted_items[0])):
        project_root = os.path.join(extract_dir, extracted_items[0])
    else:
        project_root = extract_dir
    
    # Генерация структуры
    structure = generate_complete_project_structure(project_root)
    
    # Сохранение в файл
    output_path = os.path.join(temp_dir, "project_structure.txt")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(structure)
    
    return output_path

@router.message()
async def handle_other_messages(msg: Message) -> None:
    """Обработка остальных сообщений"""
    await msg.answer(
        "<b>Отправьте файл для конвертации:</b>\n"
        "• .md файл для конвертации в DOCX\n"
        "• .zip архив для анализа структуры\n\n"
        "Используйте /help для подробной справки!"
    )

async def main() -> None:
    """Запуск бота"""
    try:
        logger.info("Запуск File Converter Bot...")
        dp.include_router(router)
        await bot.delete_webhook(drop_pending_updates=True)
        logger.info("Bot started successfully")

        await dp.start_polling(bot)
        
    except Exception as e:
        logger.critical(f"Критическая ошибка: {e}")
    finally:
        await bot.session.close()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Bot stopped by user")
    except Exception as e:
        logger.critical(f"Unhandled exception: {e}")