\# 📊 Biology Term Frequency Analyzer



Инструмент для морфологического анализа частотности терминов в текстовых корпусах (на примере учебников биологии 5–11 классов).



\## 🎯 Что делает проект



\- \*\*Морфологический поиск терминов\*\* — находит все словоформы (клетка → клетки, клеткой, клеток)

\- \*\*Поддержка мультиграмм\*\* — работает с фразами ("клеточная мембрана", "аппарат Гольджи")

\- \*\*Динамика по корпусам\*\* — показывает изменение частотности термина по годам/классам

\- \*\*Подсветка в документах\*\* — создаёт копии .docx с выделенными терминами

\- \*\*Логирование контекстов\*\* — сохраняет примеры употребления с окружением



\---



\## 🛠 Технологический стек



| Задача | Инструмент |

|--------|-----------|

| Морфологический анализ | \[Natasha](https://github.com/natasha/natasha) + \[pymorphy3](https://github.com/no-plagiarism/pymorphy3) |

| Сегментация и POS-tagging | Natasha (NewsEmbedding, NewsMorphTagger) |

| Обработка данных | pandas, numpy |

| Работа с .docx | python-docx |

| Работа с .xlsx | openpyxl |



\*\*Особенность реализации:\*\* двухслойная лемматизация (Natasha для скорости + pymorphy3 для коррекции кратких форм и видов глаголов).



\---



\## 📁 Структура проекта



biology-term-analyzer/

├── config.py # Конфигурация (пути, параметры)

├── main.py # Точка входа

├── requirements.txt

├── README.md

│

├── src/

│ ├── patch\_pkg\_resources.py # Патч совместимости pymorphy2

│ ├── utils.py # Вспомогательные функции

│ ├── preprocessing.py # Очистка и лемматизация

│ ├── analyzer.py # Поиск паттернов, подсчёт

│ └── highlighter.py # Подсветка в .docx

│

├── data/

│ ├── russian\_stopwords.txt # Стоп-слова (расширенный список)

│ └── replacements.txt # Нормализация синонимов

│

├── textbooks/ # .docx файлы корпуса (не в репозитории)

└── output/ # Результаты (генерируется автоматически)

├── term\_dynamics.xlsx # Сводная таблица

├── annotated\_textbooks/ # .docx с подсветкой

└── context\_logs/ # CSV с контекстами











\---



\## 🚀 Быстрый старт



\### 1. Установка



```bash

git clone https://github.com/ваш-username/biology-term-analyzer.git

cd biology-term-analyzer



\# Создание виртуального окружения

python -m venv venv

source venv/bin/activate  # Linux/Mac

\# venv\\Scripts\\activate   # Windows



\# Установка зависимостей

pip install -r requirements.txt



\# Загрузка данных NLTK

python -c "import nltk; nltk.download('stopwords'); nltk.download('punkt')"

