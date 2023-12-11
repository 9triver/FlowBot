from robocorp.tasks import task

from ExcelExtension import ExcelApplicationExtension as ExcelApplication


@task
def solve_challenge():
    src = ExcelApplication()
    src.open_application(visible=True)
    src.open_workbook("./data/2023.07 劳务税.xls")
    
    none_resident = ExcelApplication()
    none_resident.open_application(visible=True)
    none_resident.add_new_workbook()
    none_resident.add_new_sheet("非居民")
    
    resident = ExcelApplication()
    resident.open_application(visible=True)
    resident.add_new_workbook()
    resident.add_new_sheet("居民")
    
    # local = ExcelApplication()
    # local.open_application(visible=True)
    # local.add_new_workbook()
    
    # foreigner = ExcelApplication()
    # foreigner.open_application(visible=True)
    # foreigner.add_new_workbook()

    headers = src.read_row(row=1)
    none_resident.insert_row(row=1, row_content=headers)
    resident.insert_row(row=1, row_content=headers)
    
    src_row_num = src.find_first_available_row(row=1, column=2)
    src_row_num = src_row_num[0]
    
    none_resident_row_index = 2
    resident_row_index = 2
    for src_row_index in range(2, src_row_num + 1):
        src_row = src.read_row(row=src_row_index, column_from=1, column_to=70, header_row=1)
        if float(src_row['劳务收入_劳务税非居民']) != 0 or float(src_row['劳务税率_劳务税非居民']) != 0 or \
            float(src_row['劳务实发_劳务税非居民']) != 0 or float(src_row['劳务税_劳务税非居民']) != 0 or \
            float(src_row['劳务应扣税_劳务税非居民']) != 0:
            none_resident.insert_row(row=none_resident_row_index, row_content=src_row, header_row=1)
            none_resident_row_index = none_resident_row_index + 1
        else:
            resident.insert_row(row=resident_row_index, row_content=src_row, header_row=1)
            resident_row_index = resident_row_index + 1

        
    
    # save workbook
    none_resident.save_excel_as(filename="./output/none-resident.xls", file_format=56)
    # app.save_excel()