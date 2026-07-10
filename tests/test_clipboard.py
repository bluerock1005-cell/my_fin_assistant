import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import pyperclip
pyperclip.copy('宁波通商银行股份有限公司杭州分行\n中国民生银行股份有限公司杭州庆春支行')
from PySide6.QtWidgets import QApplication
from features.bank_classify.bank_classify_ui import BankClassifyWidget
from features.bank_classify import classify_logic as c
from pathlib import Path
app=QApplication([])
w=BankClassifyWidget()
w._paste_clipboard()
text=w._txt_input.toPlainText()
print('pasted_len=', len(text))
banks=[line.strip() for line in text.splitlines() if line.strip()]
print('banks=', banks)
out=Path('test_bank_classify_output.xlsx')
if out.exists(): out.unlink()
res=c.build_workbook(banks, out)
print('result=', res)
print('out_exists=', out.exists())
