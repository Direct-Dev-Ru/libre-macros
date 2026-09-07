#!/usr/bin/env bash
# Установка Python-макросов LibreOffice Calc из каталога macro-lib.
#
# Использование:
#   ./install.sh                          — скопировать все *.py из этой папки
#   ./install.sh --all                    — то же, без вопросов о перезаписи
#   ./install.sh collect_workbooks.py       — только указанные файлы
#   ./install.sh -v 3.6.0                   — установить с версией 3.6.0
#   ./install.sh --version 3.6.0 collect_workbooks.py
#
# Если файл уже есть в каталоге LibreOffice — сравнивается SHA-256; при совпадении
# копирование пропускается. Иначе запрашивается подтверждение перезаписи
# (ключ --all / -a — сразу перезаписать все отличающиеся, без вопросов).
# Ответ «a» / «all» / «v» / «в» / «все» — перезаписать все оставшиеся без дальнейших вопросов.
#
# Каталоги назначения (весь пакет ставится в каждый):
#   LibreOffice:  $LO_MACROS_DIR или ~/.config/libreoffice/4/user/Scripts/python/
#   AlterOffice:  $AO_MACROS_DIR или ~/.config/alteroffice/5/user/Scripts/python/
#
# Пример для другой версии LO:
#   LO_MACROS_DIR=~/.config/libreoffice/24/user/Scripts/python ./install.sh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LO_TARGET_DIR="${LO_MACROS_DIR:-${HOME}/.config/libreoffice/4/user/Scripts/python}"
AO_TARGET_DIR="${AO_MACROS_DIR:-${HOME}/.config/alteroffice/5/user/Scripts/python}"

# Параметры версии
VERSION=""
VERSION_FILE="$SCRIPT_DIR/version.txt"
MACRO_FILE="$SCRIPT_DIR/collect_workbooks.py"

# Вспомогательная функция: добавить .py если нет расширения
normalize_name() {
    local name="$1"
    if [[ "$name" == *.py ]]; then
        echo "$name"
    else
        echo "${name}.py"
    fi
}

usage() {
    cat <<'EOF'
Установка макросов LibreOffice из папки macro-lib.

  ./install.sh [опции] [файл ...]

Опции:
  -a, --all              Перезаписать все отличающиеся файлы без вопросов
  -v, --version ВЕРСИЯ   Задать версию макроса (сохраняет в version.txt
                         и обновляет MACRO_VERSION в collect_workbooks.py)
  -h, --help             Показать эту справку

Аргументы:
  файл ...               Имена макросов для установки (без .py)
                         Без аргументов — копируются все *.py

Примеры:
  ./install.sh                          — все файлы
  ./install.sh --all                    — все файлы, перезапись без вопросов
  ./install.sh collect_workbooks        — только collect_workbooks.py
  ./install.sh -v 3.6.0                 — все файлы, версия 3.6.0
  ./install.sh --version 3.6.0 functions_pp.py

Если макрос уже установлен — сравнивается SHA-256 с копией в LO; при совпадении
файл не трогается. Иначе скрипт спросит, перезаписывать ли файл.
Ключ --all / -a — сразу перезаписать все отличающиеся, без вопросов.
Ответ a / all / v / в / все — перезаписать все оставшиеся без повторных вопросов.
В отчёте отдельно: новые, перезаписанные, идентичные (хэш), пропущенные (отказ).

Переменные окружения:
  LO_MACROS_DIR   каталог Scripts/python LibreOffice (по умолчанию см. выше)
  AO_MACROS_DIR   каталог Scripts/python AlterOffice
                  (по умолчанию ~/.config/alteroffice/5/user/Scripts/python)

После установки перезапустите Calc / AlterOffice или обновите список макросов.
EOF
}

# Обновление версии в файле макроса
update_version_in_macro() {
    local version="$1"
    local macro_file="$2"

    if [[ ! -f "$macro_file" ]]; then
        echo "  Предупреждение: файл макроса не найден: $macro_file" >&2
        return 1
    fi

    # Заменяем MACRO_VERSION = "X.Y.Z" на новую версию
    if sed -i "s/^MACRO_VERSION = \"[^\"]*\"/MACRO_VERSION = \"$version\"/" "$macro_file" 2>/dev/null; then
        echo "  Версия обновлена: $version (в $macro_file)"
        return 0
    else
        echo "  Предупреждение: не удалось обновить версию в $macro_file" >&2
        return 1
    fi
}

# Сохранение версии в файл
save_version() {
    local version="$1"
    echo "$version" > "$VERSION_FILE"
    echo "  Версия сохранена: $version (в version.txt)"
}

OVERWRITE_ALL=false

file_sha256() {
    local path="$1"
    if command -v sha256sum >/dev/null 2>&1; then
        sha256sum "$path" | awk '{print $1}'
    elif command -v shasum >/dev/null 2>&1; then
        shasum -a 256 "$path" | awk '{print $1}'
    else
        echo "Ошибка: нужен sha256sum или shasum" >&2
        return 1
    fi
}

# Установить src → dest; label — для сообщений (имя файла в логе).
# Результат в INSTALL_ONE_RESULT: new | overwritten | identical | skipped | error
install_one_file() {
    local src="$1"
    local dest="$2"
    local label="$3"

    INSTALL_ONE_RESULT="error"

    if [[ ! -f "$src" ]]; then
        echo "  пропуск — не найден в macro-lib: $label" >&2
        return
    fi

    if [[ ! -f "$dest" ]]; then
        cp -f "$src" "$dest"
        echo "  установлен (новый): $label"
        INSTALL_ONE_RESULT="new"
        return
    fi

    local src_hash dest_hash
    src_hash="$(file_sha256 "$src")" || return
    dest_hash="$(file_sha256 "$dest")" || return

    if [[ "$src_hash" == "$dest_hash" ]]; then
        echo "  пропуск — идентичен (sha256 ${src_hash:0:12}…): $label"
        INSTALL_ONE_RESULT="identical"
        return
    fi

    if ask_overwrite "$label"; then
        cp -f "$src" "$dest"
        echo "  перезаписан: $label"
        INSTALL_ONE_RESULT="overwritten"
    else
        echo "  оставлен без изменений: $label"
        INSTALL_ONE_RESULT="skipped"
    fi
}

ask_overwrite() {
    local file="$1"
    local reply

    if [[ "$OVERWRITE_ALL" == true ]]; then
        return 0
    fi

    while true; do
        read -r -p "  Уже установлен: $file — перезаписать? [д/а/N] " reply
        case "$reply" in
            [aA]|[aA][lL][lL]|[vV]|в|В|все|Все|ВСЕ)
                OVERWRITE_ALL=true
                return 0
                ;;
            [yY]|[yY][eE][sS]|[dD]|[dD][aA]|д|Д|да|Да|ДА)
                return 0
                ;;
            [nN]|[nN][oO]|н|Н|нет|Нет|НЕТ|"")
                return 1
                ;;
            *)
                echo "  Введите д (да), а (все), н (нет); также y / n / all."
                ;;
        esac
    done
}

# Разбор аргументов
files_to_install=()

while [[ $# -gt 0 ]]; do
    case "$1" in
        -h|--help)
            usage
            exit 0
            ;;
        -a|--all)
            OVERWRITE_ALL=true
            shift
            ;;
        -v|--version)
            if [[ -n "${2:-}" && "$2" != -* ]]; then
                VERSION="$2"
                shift 2
            else
                echo "Ошибка: после $1 требуется указать версию" >&2
                exit 1
            fi
            ;;
        -*)
            echo "Ошибка: неизвестная опция: $1" >&2
            echo "Используйте --help для справки" >&2
            exit 1
            ;;
        *)
            files_to_install+=("$(normalize_name "$1")")
            shift
            ;;
    esac
done

# Обработка версии если задана
if [[ -n "$VERSION" ]]; then
    echo "Установка версии: $VERSION"
    save_version "$VERSION"
    if [[ -f "$MACRO_FILE" ]]; then
        update_version_in_macro "$VERSION" "$MACRO_FILE"
    fi
    echo
fi

if [[ ! -d "$SCRIPT_DIR" ]]; then
    echo "Ошибка: каталог макросов не найден: $SCRIPT_DIR" >&2
    exit 1
fi

# AlterOffice 2026: препроцессор Temp/AlterOfficeScripts ломает многострочные
# сигнатуры def (следующий def «вклеивается» в параметры → SyntaxError).
# Перед установкой сворачиваем def/async def в одну строку (идемпотентно).
if command -v python3 >/dev/null 2>&1 && [[ -f "$SCRIPT_DIR/aoffice_flatten_defs.py" ]]; then
    echo "Совместимость AlterOffice: свёртка многострочных def…"
    python3 "$SCRIPT_DIR/aoffice_flatten_defs.py" --in-place --tree "$SCRIPT_DIR" || exit 1
    echo
fi
if command -v python3 >/dev/null 2>&1 && [[ -f "$SCRIPT_DIR/aoffice_compat_check.py" ]]; then
    echo "Совместимость AlterOffice: проверка фильтра pythonscript…"
    PYTHONPATH="$SCRIPT_DIR/pythonpath${PYTHONPATH:+:$PYTHONPATH}" \
        python3 "$SCRIPT_DIR/aoffice_compat_check.py" --tree "$SCRIPT_DIR" || exit 1
    echo
fi

# Актуальная копия без docstring — из collect_workbooks.py (меню Calc часто использует её)
if [[ -f "$SCRIPT_DIR/collect_workbooks.py" && -f "$SCRIPT_DIR/generate_no_docstrings.py" ]]; then
    need_nds=0
    if [[ ${#files_to_install[@]} -eq 0 ]]; then
        need_nds=1
    else
        for f in "${files_to_install[@]}"; do
            if [[ "$f" == "collect_workbooks_no_docstrings.py" ]]; then
                need_nds=1
                break
            fi
        done
    fi
    if [[ "$need_nds" -eq 1 ]]; then
        echo "Сборка collect_workbooks_no_docstrings.py из collect_workbooks.py…"
        python3 "$SCRIPT_DIR/generate_no_docstrings.py" || exit 1
        echo
    fi
fi

# Если файлы не указаны явно — берём все *.py
if [[ ${#files_to_install[@]} -eq 0 ]]; then
    shopt -s nullglob
    for path in "$SCRIPT_DIR"/*.py; do
        base="$(basename "$path")"
        # Служебный генератор и автособранная копия — только по явному имени
        if [[ "$base" == "generate_no_docstrings.py" || "$base" == "generate_py2_version.py" || "$base" == "collect_workbooks_no_docstrings.py" ]]; then
            continue
        fi
        files_to_install+=("$base")
    done
    shopt -u nullglob

    if [[ ${#files_to_install[@]} -eq 0 ]]; then
        echo "В $SCRIPT_DIR нет файлов *.py для установки." >&2
        exit 1
    fi
fi

# Шаблоны пользовательских функций — нужны collect_workbooks и визарду
ALWAYS_WITH_COLLECT=(functions_pp.py functions_final.py)
needs_helpers=false
for f in "${files_to_install[@]}"; do
    case "$f" in
        collect_workbooks.py|collect_workbooks_no_docstrings.py|param_wizard.py)
            needs_helpers=true
            break
            ;;
    esac
done
if [[ "$needs_helpers" == true ]]; then
    for helper in "${ALWAYS_WITH_COLLECT[@]}"; do
        found=false
        for f in "${files_to_install[@]}"; do
            if [[ "$f" == "$helper" ]]; then
                found=true
                break
            fi
        done
        if [[ "$found" == false && -f "$SCRIPT_DIR/$helper" ]]; then
            files_to_install+=("$helper")
        fi
    done
fi

installed_new=0
overwritten=0
identical=0
skipped=0
errors=0
INSTALL_ONE_RESULT=""

# Установка полного пакета (entry *.py + pythonpath/) в один каталог Scripts/python.
install_package_to_dir() {
    local target_dir="$1"
    local target_label="$2"

    mkdir -p "$target_dir"

    echo "────────────────────────────────────────"
    echo "Назначение ($target_label): $target_dir"
    echo

    local file src dest
    for file in "${files_to_install[@]}"; do
        src="$SCRIPT_DIR/$file"
        dest="$target_dir/$file"

        install_one_file "$src" "$dest" "$file"
        case "$INSTALL_ONE_RESULT" in
            new) installed_new=$((installed_new + 1)) ;;
            overwritten) overwritten=$((overwritten + 1)) ;;
            identical) identical=$((identical + 1)) ;;
            skipped) skipped=$((skipped + 1)) ;;
            error) errors=$((errors + 1)) ;;
        esac
    done

    # pythonpath/ — общие библиотеки (libre_macros_lib.py и др.)
    local pp_src="$SCRIPT_DIR/pythonpath"
    local pp_dest="$target_dir/pythonpath"
    if [[ -d "$pp_src" ]]; then
        mkdir -p "$pp_dest"
        shopt -s nullglob
        local path base
        for path in "$pp_src"/*.py; do
            base="$(basename "$path")"
            dest="$pp_dest/$base"
            install_one_file "$path" "$dest" "pythonpath/$base"
            case "$INSTALL_ONE_RESULT" in
                new) installed_new=$((installed_new + 1)) ;;
                overwritten) overwritten=$((overwritten + 1)) ;;
                identical) identical=$((identical + 1)) ;;
                skipped) skipped=$((skipped + 1)) ;;
                error) errors=$((errors + 1)) ;;
            esac
        done
        shopt -u nullglob
    fi
    echo
}

echo "Источник: $SCRIPT_DIR"
if [[ "$OVERWRITE_ALL" == true ]]; then
    echo "Режим: перезапись всех отличающихся файлов без вопросов (--all)"
fi
echo

install_package_to_dir "$LO_TARGET_DIR" "LibreOffice"
install_package_to_dir "$AO_TARGET_DIR" "AlterOffice"

echo "Отчёт (оба назначения):"
echo "  новых:          $installed_new"
echo "  перезаписано:   $overwritten"
echo "  идентичны:      $identical"
echo "  без изменений:  $skipped"

if [[ $errors -gt 0 ]]; then
    echo "  не найдено:     $errors"
    echo
    echo "Готово с ошибками."
    exit 1
fi

echo
echo "Готово."
