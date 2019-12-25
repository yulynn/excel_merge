from pathlib import Path
import xlwings as xw

root_path = Path.cwd()
first = 0
file_name = root_path / '汇总.xls'
if file_name.exists():
    file_name.unlink()
for dir_item in root_path.iterdir():
    if (dir_item.is_file()):
        if (dir_item.suffix.lower() == '.xls'
                or dir_item.suffix.lower() == '.xlsx'):
            if (first == 0):
                wb = xw.Book(str(dir_item))
                first = 1
            else:
                wb1 = xw.Book(str(dir_item))
                ws1 = wb1.sheets
                ws1.api.Copy(Before=wb.sheets(1).api)
                # wb1.app.quit()
wb.save(str(file_name))
print('done')
wb.app.quit()
