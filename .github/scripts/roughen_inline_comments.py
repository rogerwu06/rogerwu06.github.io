from pathlib import Path
import io
import re
import tokenize

phrases = ['keep this', 'main part', 'needed here', 'this part', 'leave this', 'calc part', 'input stuff', 'output stuff']

def make_comment(original, n):
    body = original[1:].strip().lower()
    if original.startswith('#!') or 'coding' in body or body.startswith(('type:', 'noqa', 'pylint')):
        return original
    if not body or set(body) <= set('# -_='):
        return ['#', '##', '####', '# #'][n % 4]
    text = '# ' + phrases[n % len(phrases)]
    if len(text) > len(original):
        return '#'
    return text

for path in sorted(Path('downloads').glob('*.py')):
    src = path.read_text(encoding='utf-8')
    toks = []
    count = 0
    for tok in tokenize.generate_tokens(io.StringIO(src).readline):
        if tok.type == tokenize.COMMENT:
            count += 1
            tok = tokenize.TokenInfo(tok.type, make_comment(tok.string, count), tok.start, tok.end, tok.line)
        toks.append(tok)
    path.write_text(tokenize.untokenize(toks), encoding='utf-8')
    print('updated inline comments', path)
