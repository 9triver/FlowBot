from robocorp import browser
from robocorp.tasks import task

from RPA.Excel.Files import Files as Excel
from RPA.HTTP import HTTP


@task
def solve_challenge():
    src = Excel()
    src.open_workbook(path='./data/2023.07 劳务税.xls')
    srcTable = src.read_worksheet_as_table(name='3413', header=True)

    noneResident = Excel()
    noneResident.create_workbook(path='./output/非居民.xls', fmt='xls', sheet_name='非居民')
    noneResidentTable = []
    local = Excel()
    local.create_workbook(path='./output/国内.xls', fmt='xls', sheet_name='国内')
    localTable = []
    foreigner = Excel()
    foreigner.create_workbook(path='./output/国外.xls', fmt='xls', sheet_name='国外')
    foreignerTable = []
    
    for row in srcTable:
        if int(row['劳务收入_劳务税非居民']) != 0 or int(row['劳务税率_劳务税非居民']) != 0 or \
            int(row['劳务实发_劳务税非居民']) != 0 or int(row['劳务税_劳务税非居民']) != 0 or \
            int(row['劳务应扣税_劳务税非居民']) != 0: 
            noneResidentTable.append(row)
            continue

        id = str(row['证件号'])
        if len(id) == 18 and not id.startswith('83'):
            localTable.append(row)
            continue

        foreignerTable.append(row)

    noneResident.append_rows_to_worksheet(content=noneResidentTable, header=True)
    noneResident.save_workbook()
    local.append_rows_to_worksheet(content=localTable, header=True)
    local.save_workbook()
    foreigner.append_rows_to_worksheet(content=foreignerTable, header=True)
    foreigner.save_workbook()