"""
Предобработка текста: очистка, нормализация, лемматизация.
...
"""

from src.patch_pkg_resources import apply as _patch_pkg
_patch_pkg()

import re
import string
import unicodedata
from abc import ABC, abstractmethod
from pathlib import Path

# =============================================================================
# ПАТЧ: pkg_resources нужен pymorphy2 внутри natasha
# Должен быть ДО любого импорта natasha или pymorphy
# =============================================================================
import importlib
import sys

try:
    import pkg_resources
except ImportError:
    # pkg_resources живёт внутри setuptools — принудительно подгружаем
    import setuptools
    # Перезагружаем setuptools чтобы pkg_resources стал доступен
    if 'pkg_resources' not in sys.modules:
        from importlib.metadata import packages_distributions
        # Финальный fallback — создаём минимальный модуль-заглушку
        import types
        pkg_resources = types.ModuleType('pkg_resources')

        def _iter_entry_points(group, name=None):
            """Замена pkg_resources.iter_entry_points через importlib.metadata."""
            try:
                from importlib.metadata import entry_points
                eps = entry_points()
                if hasattr(eps, 'select'):
                    selected = eps.select(group=group)
                else:
                    selected = eps.get(group, [])
                if name:
                    selected = [ep for ep in selected if ep.name == name]
                return selected
            except Exception:
                return []

        pkg_resources.iter_entry_points = _iter_entry_points
        sys.modules['pkg_resources'] = pkg_resources

# =============================================================================
# Теперь безопасно импортируем natasha и pymorphy
# =============================================================================
import pymorphy3
import nltk
from nltk.corpus import stopwords
from nltk import word_tokenize

from natasha import Segmenter, MorphVocab, NewsEmbedding, NewsMorphTagger, Doc

from src.utils import Utils


# Части речи, которые всегда исключаются из лемм
_BANNED_POS = {'PRON', 'NUM', 'ADP', 'SCONJ', 'CCONJ', 'PART', 'INTJ', 'PUNCT', 'SYM'}

# Специальные символы для полной очистки текста
_SPEC_CHARS_FULL = (
    string.punctuation
    + '।\n\xa0«»\t…•1234567890°ρ‒–————­———ρ''"£‹£∆∆∠∠„№£→↔"♦✓abc′″€±'
    + '।∙∙‚"÷≤⊃⊥⋂⋃■↓∙∙∙∙≠×××∈∉∥∩∪≈⊂□←§¬®éeöμχ±νκυβοςζαοβζαβοδυναστημιςαχβχοχ'
    + 'żženüóōáóēžžíáłżöé·®'
)

# Специальные символы для мягкой очистки (docx — сохраняем пунктуацию)
_SPEC_CHARS_SOFT = (
    '…•°''"£‹£∆∆∠∠„№£→↔"♦′″।∙∙⊥⋂⋃■↓∙∙∙∙≠×××∈∉∥∩∪⊂□←¬'
    + 'éeöμχ±νκυβοςζαοβζαβοδυναστημιςαχβχοχżženüóōáóēžžíáłżöé·'
)

# Паттерн для месяцев — фильтруется из лемм
_MONTHS_PATTERN = re.compile(
    r'\b(?:январ[ья]|феврал[ья]|март[а]?|апрел[ья]|ма[йя]'
    r'|июн[ья]|июл[ья]|август[а]?|сентябр[ья]|октябр[ья]'
    r'|ноябр[ья]|декабр[ья])\b',
    flags=re.IGNORECASE
)


class TextPreprocessor(ABC):
    """
    Абстрактный базовый класс предобработки текста.
    
    Реализует общую логику очистки и нормализации.
    Подклассы обязаны реализовать lemmatize() и filter_lemmas().
    """

    def __init__(
        self,
        stopwords_path: str | Path = "data/russian_stopwords.txt",
        replacements_path: str | Path = "data/replacements.txt",
    ):
        """
        Args:
            stopwords_path: путь к файлу с кастомными стоп-словами (одно слово на строку)
            replacements_path: путь к файлу замен (формат "ключ: значение" на строку)
        """
        self.russian_stopwords = self._load_stopwords(stopwords_path)
        self.replacements = self._load_replacements(replacements_path)
        self._init_nlp_tools()

        # Таблица замены похожих латинских букв на кириллицу и нормализации
        self.letter_replacements = {
            'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н',
            'K': 'К', 'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т',
            'X': 'Х', 'Y': 'У', 'a': 'а', 'b': 'в', 'c': 'с',
            'e': 'е', 'h': 'н', 'k': 'к', 'm': 'м', 'o': 'о',
            'p': 'р', 't': 'т', 'x': 'х', 'y': 'у',
            'ё': 'е', 'Ё': 'Е',
            r'\b3': 'З', '0': 'О', '6': 'б',
        }

        # Двухбуквенные слова, которые не удаляются при фильтрации
        self.important_2letter = {
            'во', 'за', 'из', 'ко', 'на', 'об', 'от', 'по', 'со',
            'уж', 'мы', 'ты', 'вы', 'он', 'ей', 'их', 'бы', 'да',
            'же', 'ил', 'но', 'ну', 'то', 'ай', 'ах', 'им', 'ли',
            'ми', 'ой', 'ре', 'си', 'ту', 'яд',
        }

    # -------------------------------------------------------------------------
    # Инициализация
    # -------------------------------------------------------------------------

    @staticmethod
    def _load_stopwords(path: str | Path) -> set:
        """Загружает стоп-слова из NLTK и кастомного файла."""
        nltk_sw = set(stopwords.words("russian"))
        try:
            with open(path, encoding='utf-8') as f:
                custom_sw = {line.strip() for line in f if line.strip()}
        except FileNotFoundError:
            custom_sw = set()
            print(f"[Предупреждение] Файл стоп-слов не найден: {path}. "
                  f"Используются только стандартные стоп-слова NLTK.")
        return nltk_sw | custom_sw

    @staticmethod
    def _load_replacements(path: str | Path) -> dict:
        """
        Загружает словарь замен из файла.
        Формат файла: одна замена на строку вида "ключ: значение".
        """
        replacements = {}
        try:
            with open(path, encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and ':' in line:
                        key, value = line.split(':', 1)
                        replacements[key.strip()] = value.strip()
        except FileNotFoundError:
            print(f"[Предупреждение] Файл замен не найден: {path}. "
                  f"Замены применяться не будут.")
        return replacements

    def _init_nlp_tools(self):
        """Инициализирует NLP-инструменты: Natasha и pymorphy3."""
        self.segmenter = Segmenter()
        self.emb = NewsEmbedding()
        self.morph_tagger = NewsMorphTagger(self.emb)
        self.morph_vocab = MorphVocab()
        self.morph = pymorphy3.MorphAnalyzer()

    # -------------------------------------------------------------------------
    # Очистка текста
    # -------------------------------------------------------------------------

    def remove_accents(self, text: str) -> str:
        """
        Удаляет акценты (диакритические знаки) из текста.
        Краткость буквы 'й' (unicode category Mn, код U+0306) сохраняется.
        """
        normalized = unicodedata.normalize('NFD', text)
        return unicodedata.normalize('NFC', ''.join(
            char for char in normalized
            if unicodedata.category(char) != 'Mn' or char == '\u0306'
        ))

    def replace_letters(self, text: str) -> str:
        """
        Заменяет визуально похожие латинские буквы на кириллицу,
        нормализует 'ё' -> 'е', заменяет цифры-омографы (0 -> О, 6 -> б).
        """
        for pattern, replacement in self.letter_replacements.items():
            text = re.sub(pattern, replacement, text)
        return text

    def remove_special_chars(self, text: str) -> str:
        """
        Полная очистка: удаляет спецсимволы, цифры и пунктуацию.
        Используется перед лемматизацией в анализе частотности.
        """
        return Utils.remove_chars_from_text(text, _SPEC_CHARS_FULL + string.digits)

    def remove_special_chars_soft(self, text: str) -> str:
        """
        Мягкая очистка: удаляет экзотические символы, сохраняет пунктуацию.
        Используется при обработке docx для сохранения структуры текста.
        """
        return Utils.remove_chars_from_text(text, _SPEC_CHARS_SOFT)

    # -------------------------------------------------------------------------
    # Лемматизация
    # -------------------------------------------------------------------------

    def correct_lemma(self, token_text: str, natasha_lemma: str, existing_lemmas: list) -> str:
        """
        Корректирует лемму, полученную от Natasha, с помощью pymorphy3.
        
        Natasha иногда возвращает краткие формы прилагательных/причастий
        вместо полных, или неверно определяет вид глагола.
        pymorphy3 используется как корректирующий слой.
        
        Args:
            token_text: оригинальный токен из текста
            natasha_lemma: лемма от Natasha
            existing_lemmas: уже обработанные леммы (для контекста, не используется напрямую)
            
        Returns:
            Скорректированная лемма с заменой ё->е
        """
        parsed = self.morph.parse(natasha_lemma)

        # Если pymorphy3 не уверен в лемме или Natasha вернула исходную форму
        if not parsed or parsed[0].score < 0 or natasha_lemma.lower() == token_text.lower():
            parsed_token = self.morph.parse(token_text)
            if parsed_token:
                lemma = self._process_parsed_form(parsed_token[0])
                return lemma.replace('ё', 'е').replace('Ё', 'Е')
            return token_text.replace('ё', 'е').replace('Ё', 'Е')

        best_parse = parsed[0]

        # Краткие прилагательные и причастия -> полная форма
        if 'ADJS' in best_parse.tag or 'PRTS' in best_parse.tag:
            lemma = self._convert_short_to_full(best_parse)
            return lemma.replace('ё', 'е').replace('Ё', 'Е')

        # Глаголы: предпочитаем несовершенный вид (impf)
        if 'VERB' in best_parse.tag:
            for parse in parsed:
                if 'VERB' in parse.tag and 'impf' in parse.tag:
                    return parse.normal_form.replace('ё', 'е').replace('Ё', 'Е')
            return best_parse.normal_form.replace('ё', 'е').replace('Ё', 'Е')

        return natasha_lemma.replace('ё', 'е').replace('Ё', 'Е')

    def _convert_short_to_full(self, parsed_form) -> str:
        """Преобразует краткое прилагательное или причастие в полную форму."""
        for parse in self.morph.parse(parsed_form.word):
            if (('ADJF' in parse.tag and 'ADJS' not in parse.tag)
                    or ('PRTF' in parse.tag and 'PRTS' not in parse.tag)):
                return parse.normal_form
        return parsed_form.normal_form

    def _process_parsed_form(self, parsed_form) -> str:
        """Получает корректную нормальную форму с учётом типа части речи."""
        if 'ADJS' in parsed_form.tag or 'PRTS' in parsed_form.tag:
            return self._convert_short_to_full(parsed_form)
        if 'VERB' in parsed_form.tag:
            for parse in self.morph.parse(parsed_form.word):
                if 'VERB' in parse.tag and 'impf' in parse.tag:
                    return parse.normal_form
        return parsed_form.normal_form

    # -------------------------------------------------------------------------
    # Публичный интерфейс
    # -------------------------------------------------------------------------

    def preprocess(self, text: str) -> list[str]:
        """
        Полный пайплайн предобработки текста.
        
        Шаги:
            1. Удаление акцентов
            2. Удаление невидимых управляющих символов (SHY, ZWSP и др.)
            3. Замена похожих латинских букв на кириллицу
            4. Удаление спецсимволов и цифр
            5. Лемматизация (реализуется в подклассе)
            6. Фильтрация лемм (реализуется в подклассе)
            
        Args:
            text: исходный текст
            
        Returns:
            Список отфильтрованных лемм
        """
        text = self.remove_accents(text)
        text = re.sub(r'[\u00AD\u200B\u200C\u200D\u2060]', '', text)
        text = self.replace_letters(text)
        text = self.remove_special_chars(text)
        lemmas = self.lemmatize(text)
        return self.filter_lemmas(lemmas)

    @abstractmethod
    def lemmatize(self, text: str) -> list[str]:
        """
        Лемматизирует текст и возвращает список лемм.
        Реализуется в подклассах.
        """
        raise NotImplementedError

    @abstractmethod
    def filter_lemmas(self, lemmas: list[str]) -> list[str]:
        """
        Фильтрует список лемм.
        Реализуется в подклассах.
        """
        raise NotImplementedError


# =============================================================================
# Подкласс 1: стандартная предобработка (для анализа частотности)
# =============================================================================

class StandardTextPreprocessor(TextPreprocessor):
    """
    Предобработка для анализа частотности терминов.
    
    Что делает дополнительно к базовому классу:
    - Пропускает собственные имена (имена, фамилии, географические названия)
    - Удаляет месяцы, одно- и двухбуквенные токены
    - Удаляет латинские токены и римские цифры
    - Фильтрует стоп-слова
    """

    def lemmatize(self, text: str, skip_proper_nouns: bool = True) -> list[str]:
        """
        Лемматизирует текст с помощью Natasha + pymorphy3.
        
        Args:
            text: очищенный текст
            skip_proper_nouns: если True — собственные имена пропускаются
            
        Returns:
            Список лемм (строчные или с заглавной для имён собственных)
        """
        doc = Doc(text)
        doc.segment(self.segmenter)
        doc.tag_morph(self.morph_tagger)

        lemmas = []
        for token in doc.tokens:
            token.lemmatize(self.morph_vocab)

            pos = token.pos
            lemma = (token.lemma or '').replace('ё', 'е').replace('Ё', 'Е')
            corrected_lemma = self.correct_lemma(token.text, lemma, lemmas)

            parsed_word = self.morph.parse(corrected_lemma)
            if not parsed_word:
                continue
            parsed_word = parsed_word[0]

            # Отсекаем служебные части речи и слова с нулевым score
            if parsed_word.score <= 0 or pos in _BANNED_POS:
                continue

            # Определяем собственное имя
            is_proper = (
                'Geox' in parsed_word.tag
                or 'Name' in parsed_word.tag
                or 'Surn' in parsed_word.tag
            )

            if is_proper and skip_proper_nouns:
                continue

            # Собственные имена — с заглавной, остальные — строчные
            lemma_final = corrected_lemma.title() if is_proper else corrected_lemma.lower()
            lemma_final = lemma_final.replace('ё', 'е').replace('Ё', 'Е')

            lemmas.append(self.replacements.get(lemma_final, lemma_final))

        return lemmas

    def filter_lemmas(self, lemmas: list[str]) -> list[str]:
        """
        Фильтрует леммы:
        - удаляет месяцы
        - удаляет латинские слова и римские цифры
        - удаляет одиночные буквы
        - удаляет двухбуквенные токены
        - удаляет стоп-слова
        
        Returns:
            Отфильтрованный список токенов
        """
        text = ' '.join(lemmas)
        text = re.sub(r'\b[a-zA-Z]+\b', ' ', text)        # латинские слова
        text = re.sub(r'\b[ivxlcdm]+\b', ' ', text)        # римские цифры
        text = re.sub(_MONTHS_PATTERN, ' ', text)           # месяцы
        text = re.sub(r'\b\w\b', ' ', text)                 # одиночные буквы
        text = re.sub(r'\b\w{2}\b|\b\w-\w\b', ' ', text)   # двухбуквенные
        text = re.sub(r'\b3', 'З', text)                    # цифра 3 в начале слова

        tokens = word_tokenize(text)
        return [t.strip() for t in tokens if t not in self.russian_stopwords]


# =============================================================================
# Подкласс 2: предобработка без стоп-слов (для поиска паттернов терминов)
# =============================================================================

class TextPreprocessorWithoutStopwords(TextPreprocessor):
    """
    Предобработка с минимальной фильтрацией.
    
    Используется при нормализации словаря терминов и поиске их паттернов
    в тексте: важно сохранить максимум словоформ, убрав только явно
    нежелательные (служебные части речи, собственные имена).
    
    В отличие от StandardTextPreprocessor:
    - НЕ удаляет стоп-слова
    - НЕ удаляет двухбуквенные токены
    - НЕ удаляет месяцы и латиницу
    """

    def lemmatize(self, text: str) -> list[str]:
        """
        Лемматизирует текст, исключая только служебные части речи
        и собственные имена.
        
        Args:
            text: очищенный текст
            
        Returns:
            Список лемм в нижнем регистре
        """
        doc = Doc(text)
        doc.segment(self.segmenter)
        doc.tag_morph(self.morph_tagger)

        lemmas = []
        for token in doc.tokens:
            token.lemmatize(self.morph_vocab)

            pos = token.pos
            lemma = (token.lemma or '').replace('ё', 'е').replace('Ё', 'Е')
            corrected_lemma = self.correct_lemma(token.text, lemma, lemmas)

            parsed_word = self.morph.parse(corrected_lemma)
            if not parsed_word:
                continue
            parsed_word = parsed_word[0]

            # Отсекаем служебные части речи, собственные имена, нулевой score
            if (parsed_word.score <= 0
                    or 'Geox' in parsed_word.tag
                    or 'Name' in parsed_word.tag
                    or 'Surn' in parsed_word.tag
                    or pos in _BANNED_POS):
                continue

            lemma_lower = corrected_lemma.lower().replace('ё', 'е').replace('Ё', 'Е')
            lemmas.append(self.replacements.get(lemma_lower, lemma_lower))

        return lemmas

    def filter_lemmas(self, lemmas: list[str]) -> list[str]:
        """
        Минимальная фильтрация: токенизация без удаления стоп-слов.
        
        Returns:
            Список токенов
        """
        return word_tokenize(' '.join(lemmas))