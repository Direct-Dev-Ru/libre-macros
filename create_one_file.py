import os


def pack_macro_codebase(root_dir, output_file):
    # Игнорируем бинарный кэш компиляции питона
    ignore_dirs = {".git", "__pycache__", "venv", ".idea", ".vscode"}
    # Собираем только исходники и метаданные макросов
    allowed_extensions = {".py", ".xml", ".txt", ".md", ".json"}

    with open(output_file, "w", encoding="utf-8") as outfile:
        outfile.write("# ИСХОДНЫЙ КОД И СТРУКТУРА МАКРОСОВ LIBRE/ALTER OFFICE\n\n")

        # 1. Записываем чистую структуру проекта
        outfile.write("## 1. Структура проекта (без кэша)\n```text\n")
        for root, dirs, files in os.walk(root_dir):
            dirs[:] = [d for d in dirs if d not in ignore_dirs]
            level = root.replace(root_dir, "").count(os.sep)
            indent = " " * 4 * level
            outfile.write(f"{indent}{os.path.basename(root) or root}/\n")
            subindent = " " * 4 * (level + 1)
            for f in files:
                if os.path.splitext(f)[1] in allowed_extensions:
                    outfile.write(f"{subindent}{f}\n")
        outfile.write("```\n\n## 2. Содержимое исходных файлов\n\n")

        # 2. Собираем содержимое
        for root, dirs, files in os.walk(root_dir):
            dirs[:] = [d for d in dirs if d not in ignore_dirs]
            for file in files:
                ext = os.path.splitext(file)[1]
                if ext in allowed_extensions:
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, root_dir)

                    outfile.write(f"### Модуль: {rel_path}\n")
                    outfile.write(f"```{ext.strip('.')}\n")
                    try:
                        with open(full_path, "r", encoding="utf-8") as f:
                            outfile.write(f.read())
                    except Exception as e:
                        outfile.write(f"# Ошибка чтения файла: {str(e)}\n")
                    outfile.write("\n```\n\n")


if __name__ == "__main__":
    # Запускайте в корне папки macro-lib/
    pack_macro_codebase(".", "office_macros_context.txt")
    print("Контекст макросов успешно упакован в office_macros_context.md")
