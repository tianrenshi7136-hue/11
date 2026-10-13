import os,json
os.environ['QT_QPA_PLATFORM']='offscreen'
from pathlib import Path
from PySide6.QtWidgets import QApplication
from els.ui import Workbench
from els.supports import migrate_package
app=QApplication([]);w=Workbench(Path('outputs/support-ui-integration.db'));old=json.loads(Path('outputs/supports-integration-result.json').read_text(encoding='utf8'));p=w.project.copy();p.update(name='Software Validation Sample',subject_name='Direct Support Button Test',support_design=dict(migrate_package(old['request']['support_package']),enabled=False,include_blinding=False));w.load_form(p)
assert w.collect_steel()=={}
w.support_generate.click();assert w.support_worker is not None
assert w.support_worker.wait(55000)
app.processEvents()
text=w.support_result.toPlainText();assert '模板验算完成' in text,text
Path('outputs/support-ui-integration.txt').write_text(text,encoding='utf8');print(text);w.close()
