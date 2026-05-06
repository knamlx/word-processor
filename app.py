import os, re, subprocess, tempfile, uuid, json
from pathlib import Path
from flask import Flask, request, send_file, render_template, jsonify, Response, stream_with_context
from openai import OpenAI
from docx import Document
from docx.oxml.ns import qn
import shutil

app = Flask(__name__)

BASE_DIR    = Path(__file__).parent
SKILLS_FILE = BASE_DIR / 'skills.json'
REF_DOCX    = BASE_DIR / 'reference.docx'
CACHE_DIR   = Path(tempfile.gettempdir()) / 'wp_cache'
CACHE_DIR.mkdir(exist_ok=True)

# Загрузка .env
env_file = BASE_DIR / '.env'
if env_file.exists():
    for line in env_file.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith('#') and '=' in line:
            k, v = line.split('=', 1)
            os.environ.setdefault(k.strip(), v.strip())

API_KEY  = os.environ.get('API_KEY', '')
MODEL    = os.environ.get('LLM_MODEL', 'qwen/qwen3.6-plus')
BASE_URL = os.environ.get('LLM_BASE_URL', 'https://polza.ai/api/v1')

_client = None
def get_client():
    global _client
    if _client is None:
        _client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
    return _client


# — Skills —

def load_skills():
    if SKILLS_FILE.exists():
        return json.loads(SKILLS_FILE.read_text(encoding='utf-8'))
    return []

def save_skills(skills):
    SKILLS_FILE.write_text(json.dumps(skills, ensure_ascii=False, indent=2), encoding='utf-8')


# — Промпт —

BASE_PROMPT = """Преобразуй текст в Pandoc Markdown с LaTeX-формулами. Только Markdown, без пояснений, без ```.

ПРАВИЛА:

1. UNICODE: удали \u202f, \u00a0, невидимые символы перед % и °

2. СТРУКТУРА:
   - Заголовок -> YAML: ---\ntitle: "..."\n---
   - Заголовки разделов: # ## ###

3. ФОРМУЛЫ:
   - Встроенные в тексте: $формула$
   - Выключные (отдельная строка, пустые строки до и после): $$формула$$
   - Многострочные: $$\begin{aligned} x &= a + b \\ y &= c - d \end{aligned}$$
   - LaTeX: \frac{a}{b}, \sqrt{x}, \sum_{i=1}^{n}, \alpha, \beta, \Delta, x_{i}, x^{2}, \pm, \approx, \cdot, \leq, \geq, \Rightarrow
   - Переменные в тексте: значение $x$ равно...
   - НЕ используй \(...\) — только $...$

4. ТАБЛИЦЫ:
   - Заголовок НАД таблицей: **Таблица N — Название**
   - Формат: | col1 | col2 | и разделитель |---|---|
   - Пустая строка после таблицы

5. КОД:
   - Оборачивай в ``` с языком: python, matlab, c, cpp, java, r, sql, bash
   - Если язык неизвестен — ```text
   - Содержимое не менять

6. ПОДПИСИ к рисункам (ниже): *Рисунок N — Описание*
"""

def build_prompt(skill):
    prompt = BASE_PROMPT
    if skill and skill.get('instructions', '').strip():
        prompt += f"\n\nСКИЛЛ «{skill['name']}»:\n{skill['instructions']}"
    return prompt


# — Утилиты —

def clean_unicode(text):
    text = re.sub('\u202f', ' ', text)
    text = re.sub('[\u200b\u200c\u200d\ufeff]', '', text)
    text = re.sub('\u00a0', ' ', text)
    return re.sub('  +', ' ', text)

def strip_think(text):
    return re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL).strip()

def get_skill(skill_id):
    if not skill_id:
        return None
    return next((s for s in load_skills() if s['id'] == skill_id), None)

def call_llm(text, skill_id=None):
    resp = get_client().chat.completions.create(
        model=MODEL,
        messages=[
            {'role': 'system', 'content': build_prompt(get_skill(skill_id))},
            {'role': 'user', 'content': text},
        ],
        max_tokens=4096,
        temperature=0.2,
    )
    return clean_unicode(strip_think(resp.choices[0].message.content.strip()))


# — Pandoc —

def make_ref_docx(font):
    path = os.path.join(tempfile.gettempdir(), f'ref_{uuid.uuid4().hex[:8]}.docx')
    shutil.copy(str(REF_DOCX), path)
    try:
        doc = Document(path)
        for name in ['Normal', 'Heading 1', 'Heading 2', 'Heading 3']:
            try:
                style = doc.styles[name]
                style.font.name = font
                for rFonts in style._element.findall('.//' + qn('w:rFonts')):
                    rFonts.set(qn('w:ascii'), font)
                    rFonts.set(qn('w:hAnsi'), font)
                    rFonts.set(qn('w:cs'), font)
            except Exception:
                pass
        doc.save(path)
    except Exception:
        pass
    return path

def to_file(md, fmt, font, use_toc, toc_depth):
    uid = uuid.uuid4().hex[:8]
    tmp = tempfile.mkdtemp()
    md_path  = os.path.join(tmp, f'{uid}.md')
    out_path = os.path.join(tmp, f'{uid}.{fmt}')

    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md)

    cmd = ['pandoc', md_path, '-o', out_path]

    if fmt == 'docx':
        cmd += ['--reference-doc', make_ref_docx(font),
                '--metadata', 'lang=ru-RU',
                '--metadata', 'toc-title=Оглавление']
        if use_toc:
            cmd += ['--table-of-contents', f'--toc-depth={toc_depth}']
    else:
        cmd += ['--pdf-engine=xelatex',
                '-V', 'mainfont=DejaVu Serif',
                '-V', 'lang=ru',
                '-V', 'geometry:left=3cm,right=1.5cm,top=2cm,bottom=2cm',
                '-V', 'fontsize=14pt',
                '-V', 'linestretch=1.5',
                '-V', 'colorlinks=true']
        if use_toc:
            cmd += ['--table-of-contents', f'--toc-depth={toc_depth}']

    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(res.stderr)
    return out_path


# — Skills CRUD —

@app.route('/skills', methods=['GET'])
def skills_list():
    return jsonify(load_skills())

@app.route('/skills', methods=['POST'])
def skills_create():
    data = request.get_json()
    skill = {
        'id': uuid.uuid4().hex[:8],
        'name': data.get('name', '').strip(),
        'description': data.get('description', '').strip(),
        'instructions': data.get('instructions', '').strip(),
    }
    skills = load_skills()
    skills.append(skill)
    save_skills(skills)
    return jsonify(skill), 201

@app.route('/skills/<sid>', methods=['PUT'])
def skills_update(sid):
    data = request.get_json()
    skills = load_skills()
    for i, s in enumerate(skills):
        if s['id'] == sid:
            skills[i] = {**s,
                'name': data.get('name', s['name']).strip(),
                'description': data.get('description', s.get('description', '')).strip(),
                'instructions': data.get('instructions', s.get('instructions', '')).strip(),
            }
            save_skills(skills)
            return jsonify(skills[i])
    return jsonify({'error': 'Не найден'}), 404

@app.route('/skills/<sid>', methods=['DELETE'])
def skills_delete(sid):
    skills = load_skills()
    if not any(s['id'] == sid for s in skills):
        return jsonify({'error': 'Не найден'}), 404
    save_skills([s for s in skills if s['id'] != sid])
    return jsonify({'ok': True})


# — Routes —

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/status')
def status():
    return jsonify({'ok': bool(API_KEY), 'model': MODEL})

@app.route('/stream', methods=['POST'])
def stream():
    data = request.get_json()
    raw = clean_unicode(data.get('text', '').strip())
    if not raw:
        return jsonify({'error': 'Текст не передан'}), 400

    cache_id = data.get('cache_id', uuid.uuid4().hex[:12])
    skill = get_skill(data.get('skill_id'))

    def generate():
        full_text = ''
        in_think = False
        try:
            for chunk in get_client().chat.completions.create(
                model=MODEL,
                messages=[
                    {'role': 'system', 'content': build_prompt(skill)},
                    {'role': 'user', 'content': raw},
                ],
                max_tokens=4096,
                temperature=0.2,
                stream=True,
            ):
                delta = chunk.choices[0].delta.content or ''
                if not delta:
                    continue
                full_text += delta

                if '<think>' in delta:
                    in_think = True
                if in_think:
                    if '</think>' in delta:
                        in_think = False
                    continue

                yield f'data: {json.dumps({"chunk": delta})}\n\n'

            final_md = clean_unicode(strip_think(full_text))
            (CACHE_DIR / f'{cache_id}.md').write_text(final_md, encoding='utf-8')
            yield f'data: {json.dumps({"done": True, "cache_id": cache_id})}\n\n'

        except Exception as e:
            yield f'data: {json.dumps({"error": str(e)})}\n\n'

    return Response(
        stream_with_context(generate()),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'}
    )

@app.route('/preview', methods=['POST'])
def preview():
    data = request.get_json()
    raw = clean_unicode(data.get('text', '').strip())
    if not raw:
        return jsonify({'error': 'Текст не передан'}), 400
    try:
        md = call_llm(raw, data.get('skill_id'))
        cache_id = data.get('cache_id', uuid.uuid4().hex[:12])
        (CACHE_DIR / f'{cache_id}.md').write_text(md, encoding='utf-8')
        return jsonify({'markdown': md, 'cache_id': cache_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/process', methods=['POST'])
def process():
    data    = request.get_json()
    fmt     = data.get('format', 'docx')
    use_toc = data.get('toc', False)
    depth   = int(data.get('toc_depth', 3))
    font    = data.get('font', 'Times New Roman')

    md = None
    cache_id = data.get('cache_id')
    if cache_id:
        f = CACHE_DIR / f'{cache_id}.md'
        if f.exists():
            md = f.read_text(encoding='utf-8')

    if md is None:
        raw = clean_unicode(data.get('text', '').strip())
        if not raw:
            return jsonify({'error': 'Текст не передан'}), 400
        try:
            md = call_llm(raw, data.get('skill_id'))
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    try:
        out = to_file(md, fmt, font, use_toc, depth)
    except RuntimeError as e:
        return jsonify({'error': str(e), 'markdown': md}), 500

    mime = ('application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            if fmt == 'docx' else 'application/pdf')
    return send_file(out, as_attachment=True, download_name=f'formatted.{fmt}', mimetype=mime)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, threaded=True)
