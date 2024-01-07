from robocorp.tasks import task

from ExcelExtension import ExcelApplicationExtension as ExcelApplication

@task
def solve_challenge():
    =ExcelApplication()
    .open_application(visible=True)
    .open_workbook()
