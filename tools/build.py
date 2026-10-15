"""Đồng bộ khối EMS_GRADER vào index.html và sinh apps-script/Code.gs.

Chạy:  python tools/build.py   (từ thư mục dieu-do-htd)
"""
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parent.parent
grader = (ROOT / 'tools' / 'grader.js').read_text(encoding='utf-8').strip('\n')
server = (ROOT / 'tools' / 'server.gs').read_text(encoding='utf-8').strip('\n')

page_path = ROOT / 'index.html'
page = page_path.read_text(encoding='utf-8')
pattern = re.compile(r'(/\* <<EMS_GRADER_BEGIN>> \*/\n).*?(/\* <<EMS_GRADER_END>> \*/)', re.S)
if not pattern.search(page):
    raise SystemExit('Không tìm thấy marker EMS_GRADER trong index.html')
page = pattern.sub(lambda m: m.group(1) + grader + '\n' + m.group(2), page)
page_path.write_text(page, encoding='utf-8', newline='\n')

out = ROOT / 'apps-script' / 'Code.gs'
out.write_text(server + '\n\n' + grader + '\n', encoding='utf-8', newline='\n')
print('OK: EMS_GRADER synced into index.html and apps-script/Code.gs')
