"""
Вспомогательные утилиты: загрузка файлов, работа с DataFrame.
"""

import os
import numpy as np
import pandas as pd


class Utils:
    @staticmethod
    def remove_chars_from_text(text: str, chars: str) -> str:
        """
        Удаляет указанные символы из текста, заменяя их пробелами.
        
        Args:
            text: исходный текст
            chars: строка символов для удаления
            
        Returns:
            Текст с заменёнными символами
        """
        return "".join([' ' if ch in chars else ch for ch in text])

    @staticmethod
    def load_dataframe(file_path: str) -> pd.DataFrame:
        """
        Загружает данные из файла Excel (.xlsx) или CSV (.csv).
        
        Args:
            file_path: путь к файлу
            
        Returns:
            DataFrame с данными
            
        Raises:
            ValueError: если формат файла не поддерживается
            FileNotFoundError: если файл не найден
        """
        path = str(file_path)
        if not os.path.exists(path):
            raise FileNotFoundError(f"Файл не найден: {path}")

        ext = os.path.splitext(path)[1].lower()
        if ext == '.xlsx':
            return pd.read_excel(path)
        elif ext == '.csv':
            return pd.read_csv(path)

        raise ValueError(f"Неподдерживаемый формат файла: {ext}. Ожидается .xlsx или .csv")

    @staticmethod
    def merge_dataframes(
        df1: pd.DataFrame,
        df2: pd.DataFrame,
        key_column1: str,
        key_column2: str
    ) -> pd.DataFrame:
        """
        Объединяет два DataFrame по ключевым столбцам (left join).
        Числовые пропуски заполняются нулями, строковые — пустой строкой.
        
        Args:
            df1: основной DataFrame
            df2: присоединяемый DataFrame
            key_column1: ключевой столбец в df1
            key_column2: ключевой столбец в df2
            
        Returns:
            Объединённый DataFrame
        """
        df2 = df2.rename(columns={key_column2: key_column1})
        merged_df = pd.merge(df1, df2, on=key_column1, how='left')

        numeric_cols = merged_df.select_dtypes(include=[np.number]).columns
        merged_df[numeric_cols] = merged_df[numeric_cols].fillna(0)

        string_cols = merged_df.select_dtypes(include=[object]).columns
        merged_df[string_cols] = merged_df[string_cols].fillna('')

        return merged_df