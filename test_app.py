"""
Полный тест word-processor
Запуск: python3 test_app.py
Требует запущенного сервера на http://localhost:5000
"""
import requests
import json
import os
import zipfile
import re
import time
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO

BASE = "http://localhost:5000"
PASS = "\033[92m✓\033[0m"
FAIL = "\033[91m✗\033[0m"
WARN = "\033[93m⚠\033[0m"

results = []

def check(name, ok, detail=""):
    icon = PASS if ok else FAIL
    print(f"  {icon} {name}" + (f": {detail}" if detail else ""))
    results.append((name, ok, detail))

def section(title):
    print(f"\n{'═'*55}")
    print(f"  {title}")
    print(f"{'═'*55}")

section("1. СЕРВЕР — доступность и статус")

try:
    r = requests.get(f"{BASE}/status", timeout=5)
    check("Сервер отвечает", r.status_code == 200)
    data = r.json()
    check("Ключ API задан", data.get("ok") == True, f"ok={data.get('ok')}")
    check("Модель указана", bool(data.get("model")), data.get("model"))
except Exception as e:
    check("Сервер отвечает", False, str(e))
    print("\n⛔ Сервер недоступен. Запусти: docker compose up")
    exit(1)

try:
    r = requests.get(f"{BASE}/")
    check("Главная страница отдаётся", r.status_code == 200)
    check("HTML содержит интерфейс", "word-processor" in r.text.lower() or "скилл" in r.text.lower())
except Exception as e:
    check("Главная страница", False, str(e))


section("2. SKILLS — CRUD операции")

# GET
try:
    r = requests.get(f"{BASE}/skills")
    check("GET /skills возвращает 200", r.status_code == 200)
    check("GET /skills возвращает список", isinstance(r.json(), list))
except Exception as e:
    check("GET /skills", False, str(e))

# CREATE
skill_id = None
try:
    payload = {
        "name": "Тест-скилл",
        "description": "Скилл для тестирования",
        "instructions": "Нумеровать формулы: (1), (2). Заголовки по центру."
    }
    r = requests.post(f"{BASE}/skills", json=payload)
    check("POST /skills создаёт скилл", r.status_code == 201)
    data = r.json()
    skill_id = data.get("id")
    check("Скилл получает id", bool(skill_id), skill_id)
    check("Имя сохранено корректно", data.get("name") == "Тест-скилл")
    check("Описание сохранено", data.get("description") == "Скилл для тестирования")
    check("Инструкции сохранены", bool(data.get("instructions")))
except Exception as e:
    check("POST /skills", False, str(e))

# READ после создания
try:
    r = requests.get(f"{BASE}/skills")
    skills = r.json()
    found = any(s.get("id") == skill_id for s in skills)
    check("Созданный скилл виден в списке", found)
except Exception as e:
    check("Чтение после создания", False, str(e))

# UPDATE
try:
    r = requests.put(f"{BASE}/skills/{skill_id}", json={"name": "Тест-скилл-обновлён", "instructions": "Новые правила"})
    check("PUT /skills/:id обновляет", r.status_code == 200)
    check("Имя обновилось", r.json().get("name") == "Тест-скилл-обновлён")
except Exception as e:
    check("PUT /skills/:id", False, str(e))

# DELETE несуществующего
try:
    r = requests.delete(f"{BASE}/skills/nonexistent_id_123")
    check("DELETE несуществующего → 404", r.status_code == 404)
except Exception as e:
    check("DELETE несуществующего", False, str(e))

section("3. PREVIEW — форматирование текста")

def preview(text, skill_id=None):
    payload = {"text": text}
    if skill_id:
        payload["skill_id"] = skill_id
    r = requests.post(f"{BASE}/preview", json=payload, timeout=60)
    return r

# Пустой текст
try:
    r = preview("")
    check("Пустой текст → ошибка 400", r.status_code == 400)
except Exception as e:
    check("Пустой текст", False, str(e))

# Обычный текст
try:
    r = preview("Это простой текст для проверки форматирования.")
    check("Простой текст → 200", r.status_code == 200)
    data = r.json()
    check("Возвращает markdown", "markdown" in data)
    check("Возвращает cache_id", "cache_id" in data)
    check("Markdown непустой", bool(data.get("markdown", "").strip()))
    cache_id_simple = data.get("cache_id")
except Exception as e:
    check("Простой текст", False, str(e))
    cache_id_simple = None

# Unicode мусор
try:
    dirty = "КПД\u202f=\u202f85\u202f%. Значение\u00a0x равно\u200b5."
    r = preview(dirty)
    check("Unicode-мусор → 200", r.status_code == 200)
    md = r.json().get("markdown", "")
    check("\\u202f убран из результата", "\u202f" not in md)
    check("\\u200b убран из результата", "\u200b" not in md)
    check("\\u00a0 убран из результата", "\u00a0" not in md)
except Exception as e:
    check("Unicode очистка", False, str(e))

# Формулы
try:
    formula_text = """
    Скорость вычисляется по формуле v = delta_s / delta_t.
    Уравнение Эйнштейна: E = m * c^2.
    КПД: eta = W_пол / W_затр умножить на 100%.
    Сумма: сумма от i=1 до n от x_i в квадрате.
    Дробь: a делить на b плюс корень из x.
    """
    r = preview(formula_text)
    check("Текст с формулами → 200", r.status_code == 200)
    md = r.json().get("markdown", "")
    check("Markdown содержит LaTeX $", "$" in md, f"найдено $ в markdown: {'да' if '$' in md else 'нет'}")
    check("Есть выключные формулы $$", "$$" in md)
    has_frac = "\\frac" in md or "frac" in md.lower()
    check("Дроби конвертированы", has_frac)
except Exception as e:
    check("Формулы", False, str(e))

# Таблица
try:
    table_text = """
    Результаты измерений:
    | Параметр | Значение | Единица |
    | Температура | 25 | °C |
    | Давление | 101.3 | кПа |
    | Влажность | 65 | % |
    """
    r = preview(table_text)
    check("Таблица → 200", r.status_code == 200)
    md = r.json().get("markdown", "")
    check("Markdown содержит таблицу (|)", "|" in md)
    check("Таблица имеет разделитель (---)", "---" in md or "---" in md)
except Exception as e:
    check("Таблица", False, str(e))

# Код
try:
    code_text = """
    Программа на Python:
    import math
    def f(x):
        return x**2 + math.exp(-x)
    result = f(3.14)
    print(result)
    """
    r = preview(code_text)
    check("Код → 200", r.status_code == 200)
    md = r.json().get("markdown", "")
    check("Код обёрнут в ```", "```" in md)
except Exception as e:
    check("Код", False, str(e))

# Со скиллом
try:
    if skill_id:
        r = preview("Текст для форматирования со скиллом.", skill_id=skill_id)
        check("Preview со скиллом → 200", r.status_code == 200)
        check("Скилл применяется (есть markdown)", bool(r.json().get("markdown")))
except Exception as e:
    check("Preview со скиллом", False, str(e))

# Think-теги от Qwen
try:
    r = preview("Простой текст")
    md = r.json().get("markdown", "")
    check("Нет <think> тегов в ответе", "<think>" not in md and "</think>" not in md)
    check("Нет ``` обёртки вокруг всего", not md.strip().startswith("```"))
except Exception as e:
    check("Think-теги", False, str(e))

section("4. DOCX — генерация Word документа")


def download_docx(text, skill_id=None, cache_id=None, font="Times New Roman", toc=False):
    payload = {
        "text": text if not cache_id else "",
        "format": "docx",
        "font": font,
        "toc": toc,
        "toc_depth": 3,
    }
    if skill_id:
        payload["skill_id"] = skill_id
    if cache_id:
        payload["cache_id"] = cache_id
    r = requests.post(f"{BASE}/process", json=payload, timeout=120)
    return r

# Базовый docx
try:
    r = download_docx("Введение\n\nЭто тестовый документ. Формула: E = mc^2.")
    check("Базовый DOCX → 200", r.status_code == 200)
    check("Content-Type = docx", "wordprocessingml" in r.headers.get("Content-Type", ""))
    check("Файл непустой", len(r.content) > 5000, f"{len(r.content)} байт")
    
    # Проверяем что это валидный ZIP (docx = zip)
    try:
        with zipfile.ZipFile(BytesIO(r.content)) as z:
            files = z.namelist()
            check("DOCX — валидный ZIP", True)
            check("DOCX содержит document.xml", "word/document.xml" in files)
            check("DOCX содержит settings.xml", "word/settings.xml" in files)
            
            # Проверяем Cambria Math
            settings = z.read("word/settings.xml").decode("utf-8")
            check("Cambria Math в settings.xml", "Cambria Math" in settings or "mathPr" in settings)
            
            # Проверяем формулы
            doc_xml = z.read("word/document.xml").decode("utf-8")
            check("OMML формулы присутствуют", "<m:oMath" in doc_xml or "oMath" in doc_xml)
    except Exception as e:
        check("DOCX структура", False, str(e))
except Exception as e:
    check("Базовый DOCX", False, str(e))

# Кеш работает (не вызывает LLM повторно)
try:
    if cache_id_simple:
        t0 = time.time()
        r = download_docx("", cache_id=cache_id_simple)
        t1 = time.time()
        check("DOCX из кеша → 200", r.status_code == 200)
        check("DOCX из кеша быстрее 15 сек", (t1 - t0) < 15, f"{t1-t0:.1f}с")
except Exception as e:
    check("DOCX кеш", False, str(e))

# Шрифты
for font in ["Times New Roman", "Arial", "Calibri"]:
    try:
        r = download_docx(f"Тест шрифта {font}.", font=font)
        check(f"DOCX с шрифтом {font} → 200", r.status_code == 200)
        if r.status_code == 200:
            with zipfile.ZipFile(BytesIO(r.content)) as z:
                # Проверяем что шрифт упомянут в document.xml или styles.xml
                doc = z.read("word/document.xml").decode("utf-8")
                styles = z.read("word/styles.xml").decode("utf-8") if "word/styles.xml" in z.namelist() else ""
                has_font = font in doc or font in styles
                check(f"  Шрифт {font} применён в документе", has_font)
    except Exception as e:
        check(f"DOCX шрифт {font}", False, str(e))

# TOC
try:
    r = download_docx("# Глава 1\n\nТекст.\n\n## Раздел 1.1\n\nПодтекст.", toc=True)
    check("DOCX с оглавлением → 200", r.status_code == 200)
    if r.status_code == 200:
        with zipfile.ZipFile(BytesIO(r.content)) as z:
            doc = z.read("word/document.xml").decode("utf-8")
            check("Оглавление содержит TOC поле", "TOC" in doc or "toc" in doc.lower())
except Exception as e:
    check("DOCX с TOC", False, str(e))

# Большой текст
try:
    big_text = """
    # Исследование ионосферных параметров
    
    ## 1. Введение
    
    Прогнозирование параметров ионосферы f0F2 является актуальной задачей.
    Критическая частота вычисляется по формуле: f = sqrt(81 * N).
    
    ## 2. Методология
    
    Коэффициент детерминации R^2 = 1 - (SS_res / SS_tot).
    
    Таблица результатов:
    | Модель | RMSE | MAE | R2 |
    | CatBoost | 0.42 | 0.31 | 0.87 |
    | RandomForest | 0.51 | 0.39 | 0.83 |
    
    ## 3. Код модели
    
    from catboost import CatBoostRegressor
    model = CatBoostRegressor(iterations=1000)
    model.fit(X_train, y_train)
    
    ## 4. Заключение
    
    Модель CatBoost показала лучшие результаты.
    """
    r = download_docx(big_text)
    check("Большой текст (формулы+таблица+код) → 200", r.status_code == 200)
    check("Файл > 10KB", len(r.content) > 10000, f"{len(r.content)} байт")
except Exception as e:
    check("Большой текст", False, str(e))

section("5. PDF — генерация")


try:
    payload = {
        "text": "# Тест PDF\n\nФормула: $E = mc^2$\n\nТекст документа.",
        "format": "pdf",
        "font": "DejaVu Serif",
        "toc": False,
        "toc_depth": 3,
    }
    r = requests.post(f"{BASE}/process", json=payload, timeout=180)
    check("PDF → 200", r.status_code == 200, f"status={r.status_code}")
    if r.status_code == 200:
        check("Content-Type = pdf", "pdf" in r.headers.get("Content-Type", "").lower())
        check("PDF непустой > 5KB", len(r.content) > 5000, f"{len(r.content)} байт")
        check("PDF начинается с %PDF", r.content[:4] == b"%PDF")
    else:
        try:
            err = r.json().get("error", "")[:200]
            check("Ошибка PDF", False, err)
        except:
            pass
except Exception as e:
    check("PDF генерация", False, str(e))

section("6. STREAMING — потоковая генерация")

try:
    payload = {"text": "Скорость v = s / t. Ускорение a = delta_v / delta_t."}
    r = requests.post(f"{BASE}/stream", json=payload, stream=True, timeout=90)
    check("Stream → 200", r.status_code == 200)
    check("Content-Type text/event-stream", "event-stream" in r.headers.get("Content-Type", ""))
    
    chunks = []
    done_received = False
    cache_id_stream = None
    
    for line in r.iter_lines(decode_unicode=True):
        if line.startswith("data: "):
            try:
                ev = json.loads(line[6:])
                if ev.get("chunk"):
                    chunks.append(ev["chunk"])
                if ev.get("done"):
                    done_received = True
                    cache_id_stream = ev.get("cache_id")
                if ev.get("error"):
                    check("Stream без ошибок", False, ev["error"])
                    break
            except:
                pass
    
    check("Stream получает чанки", len(chunks) > 0, f"{len(chunks)} чанков")
    check("Stream получает done=True", done_received)
    check("Stream возвращает cache_id", bool(cache_id_stream))
    
    full_md = "".join(chunks)
    check("Stream: нет <think> тегов", "<think>" not in full_md)
    check("Stream: есть LaTeX формулы", "$" in full_md)
    
    # Проверяем что cache работает после stream
    if cache_id_stream:
        r2 = download_docx("", cache_id=cache_id_stream)
        check("DOCX из stream-кеша → 200", r2.status_code == 200)
        
except Exception as e:
    check("Streaming", False, str(e))

section("7. EDGE CASES — граничные случаи")

# Только формулы
try:
    r = preview("x^2 + y^2 = z^2. sum i=1 to n x_i. integral 0 to inf e^-x dx.")
    check("Только формулы → 200", r.status_code == 200)
    md = r.json().get("markdown", "")
    check("Формулы конвертированы", "$" in md)
except Exception as e:
    check("Только формулы", False, str(e))

# Длинный текст (много токенов)
try:
    long_text = "Это предложение для теста. " * 100
    r = preview(long_text)
    check("Длинный текст (2700 симв) → 200", r.status_code == 200)
except Exception as e:
    check("Длинный текст", False, str(e))

# Несуществующий скилл
try:
    r = preview("Текст", skill_id="nonexistent_skill_123")
    check("Несуществующий скилл → не падает", r.status_code == 200)
except Exception as e:
    check("Несуществующий скилл", False, str(e))

# Спецсимволы
try:
    special = 'Текст с "кавычками" и \'апострофами\', амперсанд & символ < > и формула α=β.'
    r = preview(special)
    check("Спецсимволы → 200", r.status_code == 200)
except Exception as e:
    check("Спецсимволы", False, str(e))

# Неверный формат
try:
    r = requests.post(f"{BASE}/process", json={"text": "тест", "format": "xlsx"}, timeout=30)
    check("Неверный формат → не падает сервер", r.status_code in [400, 500])
except Exception as e:
    check("Неверный формат", False, str(e))

section("8. УДАЛЕНИЕ СКИЛЛА — очистка")


try:
    if skill_id:
        r = requests.delete(f"{BASE}/skills/{skill_id}")
        check("DELETE /skills/:id → 200", r.status_code == 200)
        
        r = requests.get(f"{BASE}/skills")
        found = any(s.get("id") == skill_id for s in r.json())
        check("Скилл удалён из списка", not found)
except Exception as e:
    check("DELETE скилл", False, str(e))


section("ИТОГ")

total = len(results)
passed = sum(1 for _, ok, _ in results if ok)
failed = total - passed

print(f"\n  Всего проверок: {total}")
print(f"  {PASS} Прошло:  {passed}")
print(f"  {FAIL} Упало:   {failed}")

if failed > 0:
    print(f"\n  Проваленные тесты:")
    for name, ok, detail in results:
        if not ok:
            print(f"    {FAIL} {name}" + (f": {detail}" if detail else ""))

print()
