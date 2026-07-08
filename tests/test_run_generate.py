import os
import time
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import pyperclip
pyperclip.copy('宁波通商银行股份有限公司杭州分行\n中国民生银行股份有限公司杭州庆春支行')
from PySide6.QtWidgets import QApplication
from features.bank_classify.ui import BankClassifyWidget
from pathlib import Path
app = QApplication([])
w = BankClassifyWidget()
# Paste to fill input
w._paste_clipboard()
# Ensure out path temporary
out = Path('test_gen_output.xlsx')
if out.exists(): out.unlink()
w._out_path = out
# Simulate button click to start processing
w._process()
# Wait for thread to finish (timeout 10s)
start = time.time()
while True:
    app.processEvents()
    t = getattr(w, '_thread', None)
    if t is None or not t.isRunning():
        break
    if time.time() - start > 10:
        print('timeout waiting for thread')
        break
    time.sleep(0.1)
# Print log contents and file existence
print('log:\n', w._log.toPlainText())
print('progress label:', w._progress.text())
print('out exists:', out.exists())
if out.exists():
    print('build file size:', out.stat().st_size)
