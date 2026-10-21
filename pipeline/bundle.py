"""Build preview.html: index.html with data.json inlined (single file, no server needed)."""
import os, json
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
html = open(os.path.join(ROOT, 'index.html')).read()
data = open(os.path.join(ROOT, 'data.json')).read().replace('</', '<\\/')
out = html.replace('<script>\n/* ====', '<script id="inline-data" type="application/json">' + data + '</script>\n<script>\n/* ====', 1)
open(os.path.join(ROOT, 'preview.html'), 'w').write(out)
print('preview.html', round(os.path.getsize(os.path.join(ROOT, 'preview.html')) / 1e6, 2), 'MB')
