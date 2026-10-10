from pathlib import Path
import re

ROOT = Path('downloads')

pools = {
    'input': ['input stuff', 'inputs here', 'get the inputs', 'input values'],
    'output': ['output stuff', 'output here', 'make the output', 'staad output'],
    'table': ['table stuff', 'tables', 'table part', 'data table'],
    'calc': ['calc part', 'main calc', 'do the calc', 'calculation stuff'],
    'gui': ['gui stuff', 'window stuff', 'screen layout', 'display stuff'],
    'file': ['file stuff', 'files here', 'path stuff', 'save files'],
    'abaqus': ['abaqus stuff', 'abaqus part', 'abaqus run', 'model stuff'],
    'job': ['job stuff', 'run the jobs', 'job part', 'jobs here'],
    'mesh': ['mesh stuff', 'meshing part', 'mesh here', 'geometry stuff'],
    'result': ['results', 'result stuff', 'get results', 'results here'],
    'default': ['this part', 'keep this here', 'main part', 'stuff below', 'needed here', 'leave this like this'],
}

def pick_pool(text):
    t = text.lower()
    if 'input' in t or 'paste' in t or 'value' in t:
        return 'input'
    if 'output' in t or 'write' in t or 'format' in t:
        return 'output'
    if 'table' in t or 'treeview' in t or 'row' in t or 'column' in t:
        return 'table'
    if 'calculate' in t or 'calculation' in t or 'equation' in t or 'factor' in t or 'load' in t:
        return 'calc'
    if 'gui' in t or 'window' in t or 'layout' in t or 'style' in t or 'button' in t or 'font' in t:
        return 'gui'
    if 'file' in t or 'folder' in t or 'path' in t or 'csv' in t or 'excel' in t:
        return 'file'
    if 'abaqus' in t or 'cae' in t or 'odb' in t:
        return 'abaqus'
    if 'job' in t or 'submit' in t or 'pressure' in t:
        return 'job'
    if 'mesh' in t or 'geometry' in t or 'step' in t or 'part' in t:
        return 'mesh'
    if 'result' in t or 'stress' in t or 'peeq' in t or 'failure' in t:
        return 'result'
    return 'default'

def clean_generated(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9 ]+', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def transform_file(path):
    lines = path.read_text(encoding='utf-8').splitlines(True)
    out = []
    comment_count = 0
    top_count = 0

    for i, line in enumerate(lines):
        stripped = line.lstrip()
        indent = line[:len(line) - len(stripped)]

        if i == 0 and stripped.startswith('#!'):
            out.append(line)
            continue
        if re.match(r'#.*coding[:=]', stripped):
            out.append(line)
            continue

        if stripped.startswith('#'):
            body = stripped[1:].strip()
            if body.startswith('type:') or body.startswith('noqa') or body.startswith('pylint'):
                out.append(line)
                continue

            comment_count += 1
            if not body or set(body) <= set('# -_='):
                patterns = ['#\n', '##\n', '####\n', '#\n', '# #\n']
                out.append(indent + patterns[comment_count % len(patterns)])
                continue

            key = pick_pool(body)
            phrases = pools[key]
            phrase = phrases[comment_count % len(phrases)]
            out.append(indent + '# ' + clean_generated(phrase) + '\n')
            continue

        # put a few plain separators around top level sections only
        if line.startswith(('class ', 'def ')):
            top_count += 1
            if top_count % 3 == 0:
                if out and out[-1].strip() != '':
                    out.append('\n')
                if top_count % 2:
                    out.extend(['#\n', '# main part\n', '#\n'])
                else:
                    out.extend(['####\n', '# this part\n', '####\n'])

        out.append(line)

    path.write_text(''.join(out), encoding='utf-8')

for path in sorted(ROOT.glob('*.py')):
    transform_file(path)
    print('updated', path)
