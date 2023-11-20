from robocorp.tasks import task

from ExcelExtension import ExcelApplicationExtension as ExcelApplication

@task
def solve_challenge():
    app = ExcelApplication()
    app.open_application(visible=True)
    app.open_workbook('./data/2023.07 劳务税.xls')
    for i in range(0, 10):
        print(app.read_row(header=True))
        app.move_active_cell(1, 0)
    app.save_excel()