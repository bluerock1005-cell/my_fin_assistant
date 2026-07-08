from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from features.bank_classify import classify_logic as c
out=Path('test_direct_thread.xlsx')
if out.exists(): out.unlink()
with ThreadPoolExecutor(max_workers=1) as ex:
    fut = ex.submit(c.build_workbook, ['宁波通商银行股份有限公司杭州分行','中国民生银行股份有限公司杭州庆春支行'], out)
    res = fut.result()
    print('res=',res)
print('exists=', out.exists(), 'size=', out.stat().st_size if out.exists() else None)
