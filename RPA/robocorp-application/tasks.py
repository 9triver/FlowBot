from robocorp.tasks import task

from ExcelExtension import ExcelApplicationExtension as ExcelApplication

@task
def solve_challenge():
    # above fixed
    
    # open workbook {./data/2023.07 劳务税.xls} as {app}
    app = ExcelApplication()
    app.open_application(visible=True)
    app.open_workbook('./data/2023.07 劳务税.xls')
    # for loop {i} from {0} to {10}
    contents = app.read_area(None, 9, 1, 11, True)
    for content in contents:
        print(content)
    print('---------------------------------------')
    contents = app.read_area(3, 9, None, 5, False)
    for content in contents:
        print(content)
    # save workbook
    app.save_excel()