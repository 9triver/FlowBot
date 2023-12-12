from robocorp.tasks import task

from ExcelExtension import ExcelApplicationExtension as ExcelApplication


@task
def solve_challenge():
    # open {./data/2023.07 劳务税.xls} as {src}
    src = ExcelApplication()
    src.open_application(visible=True)
    src.open_workbook("./data/2023.07 劳务税.xls")

    # Add Workbook As {none_resident}
    none_resident = ExcelApplication()
    none_resident.open_application(visible=True)
    none_resident.add_new_workbook()
    # {none_resident} Create Sheet {非居民}
    none_resident.add_new_sheet('非居民')

    # Add Workbook As {resident}
    resident = ExcelApplication()
    resident.open_application(visible=True)
    resident.add_new_workbook()
    # {none_resident} Create Sheet {居民}
    resident.add_new_sheet("居民")

    # {src} Fetch Row {1} {} {} {} As {headers}
    headers = src.read_row(row=1)
    # {none_resident} Insert Row {1} {headers} {}
    none_resident.insert_row(row=1, row_content=headers)
    # {resident} Insert Row {1} {headers} {}
    resident.insert_row(row=1, row_content=headers)

    # Set {none_resident_row_index} to {2}
    none_resident_row_index = 2
    # Set {resident_row_index} to {2}
    resident_row_index = 2
    # for {src_row_index} from {2} to {100}
    for src_row_index in range(2, 100 + 1):
        # {src} Fetch Row {src_row_index} {1} {70} {1} As {src_row}
        src_row = src.read_row(row=src_row_index, column_from=1, column_to=70, header_row=1)
        # if
        # or 
        # Compare {float} {src_row['劳务收入_劳务税非居民']} {>} {0}
        # 另外四个 condition block
        if float(src_row['劳务收入_劳务税非居民']) > 0 or float(src_row['劳务税率_劳务税非居民']) > 0 or \
            float(src_row['劳务实发_劳务税非居民']) > 0 or float(src_row['劳务税_劳务税非居民']) > 0 or \
            float(src_row['劳务应扣税_劳务税非居民']) > 0:
            # {none_resident} Insert Row {none_resident_row_index} {src_row} {1}
            none_resident.insert_row(row=none_resident_row_index, row_content=src_row, header_row=1)
            # Set {none_resident_row_index} to {none_resident_row_index + 1}
            none_resident_row_index = none_resident_row_index + 1
        # else
        else:
            # {resident} Insert Row {resident_row_index} {src_row} {1}
            resident.insert_row(row=resident_row_index, row_content=src_row, header_row=1)
            # Set {resident_row_index} to {resident_row_index + 1}
            resident_row_index = resident_row_index + 1

        
    
    # {none_resident} Save Workbook {./output/none-resident.xls}
    none_resident.save_excel_as(filename="./output/none-resident.xls", file_format=56)
    # {resident} Save Workbook {./output/resident.xls}
    resident.save_excel_as(filename="./output/resident.xls", file_format=56)