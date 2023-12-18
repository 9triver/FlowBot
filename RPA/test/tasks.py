from robocorp.tasks import task

from ExcelExtension import ExcelApplicationExtension as ExcelApplication

@task
def solve_challenge():
    name=ExcelApplication()
    name.open_application(visible=True)
    name.open_workbook('file')
    app.move_active_cell(row_change=1,column_change=2)
