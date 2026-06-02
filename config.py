"""
Конфигурация проекта: пути, параметры алгоритма, настройки логирования.
"""

import os
from pathlib import Path

# =============================================================================
# Пути к данным
# =============================================================================

PROJECT_ROOT = Path(__file__).parent.resolve()
DATA_DIR = PROJECT_ROOT / "data"

# Вспомогательные файлы
STOPWORDS_PATH = DATA_DIR / "russian_stopwords.txt"
REPLACEMENTS_PATH = DATA_DIR / "replacements.txt"

# !!!!! УКАЖИТЕ ЗДЕСЬ ПУТЬ К ВАШЕМУ СЛОВАРЮ ТЕРМИНОВ !!!!!
DICT_PATH = Path(r"D:\term-analyzer\data\terms_dictionary.xlsx")

# !!!!! УКАЖИТЕ ЗДЕСЬ ДИРЕКТОРИЮ С ВАШИМИ УЧЕБНИКАМИ !!!!!
TEXTBOOKS_DIR = Path(r"D:\term-analyzer\textbooks")

# Результаты
OUTPUT_DIR = PROJECT_ROOT / "output"
OUT_XLSX = OUTPUT_DIR / "term_dynamics.xlsx"
OUT_BOOKS_DIR = OUTPUT_DIR / "annotated_textbooks"
LOG_DIR = OUTPUT_DIR / "context_logs"

# =============================================================================
# Пути к учебникам по классам
# =============================================================================

GRADES = (5, 6, 7, 8, 9, 10, 11)

GRADE_DOCX = {
    5:  TEXTBOOKS_DIR / "Биология 5 база.docx",
    6:  TEXTBOOKS_DIR / "Биология 6 база.docx",
    7:  TEXTBOOKS_DIR / "Биология 7 база.docx",
    8:  TEXTBOOKS_DIR / "Биология 8 база.docx",
    9:  TEXTBOOKS_DIR / "Биология 9 база.docx",
    10: TEXTBOOKS_DIR / "Биология 10 база.docx",
    11: TEXTBOOKS_DIR / "Биология 11 база.docx",
}

FALLBACK_TOTALS = {
    5: 35834,
    6: 33572,
    7: 37901,
    8: 67138,
    9: 73704,
    10: 56202,
    11: 65891
}

# =============================================================================
# Параметры алгоритма
# =============================================================================

ALLOW_PUNCT_BETWEEN = False
DISALLOW_HYPHEN_ADJACENT_FOR_UNIGRAMS = False

# =============================================================================
# Параметры логирования
# =============================================================================

DEBUG_ALL_TERMS = True
DEBUG_TERMS = []
MAX_CONTEXTS_PER_TERM = 20
CONTEXT_WINDOW = 6

# =============================================================================
# Параметры нормировки
# =============================================================================

NORM_PER = 1_000_000

# =============================================================================
# Валидация
# =============================================================================

def validate_config():
    warnings = []
    
    if not STOPWORDS_PATH.exists():
        warnings.append(f"Файл стоп-слов не найден: {STOPWORDS_PATH}")
    if not REPLACEMENTS_PATH.exists():
        warnings.append(f"Файл замен не найден: {REPLACEMENTS_PATH}")
    if not DICT_PATH.exists():
        warnings.append(f"Словарь терминов не найден: {DICT_PATH}")
    if not TEXTBOOKS_DIR.exists():
        warnings.append(f"Директория учебников не найдена: {TEXTBOOKS_DIR}")
    else:
        for grade, path in GRADE_DOCX.items():
            if not path.exists():
                warnings.append(f"Учебник {grade} класса не найден: {path}")
    
    return warnings

def print_config_summary():
    print("=" * 60)
    print("КОНФИГУРАЦИЯ ПРОЕКТА")
    print("=" * 60)
    print(f"Корень проекта:     {PROJECT_ROOT}")
    print(f"Словарь терминов:   {DICT_PATH}")
    print(f"Учебники:           {TEXTBOOKS_DIR}")
    print(f"Выходная папка:     {OUTPUT_DIR}")
    print(f"Классы:             {GRADES}")
    print()
    
    issues = validate_config()
    if issues:
        print("ПРЕДУПРЕЖДЕНИЯ:")
        for w in issues:
            print(f"  ⚠  {w}")
    else:
        print("✓ Все файлы и директории найдены.")
    print("=" * 60)

if __name__ == "__main__":
    print_config_summary()