print('Welcome to the "Trebo Hotel"!\n')
print('To start the program, please enter the account type: \n')
print("Worker: 'worker'")
print("Manager: 'manager'")
print("Director: 'director'")
print("Marketing specialist: 'marketing'")
print("Sale-manager: 'sales'")


import openpyxl
import entering
book = openpyxl.load_workbook("staff.xlsx")
running = True
while running:
    job_type = input('Account type: ').lower()
    if job_type == 'worker':
        entering.worker()
    elif job_type == 'manager':
        entering.manager()
    elif job_type == 'marketing':
        entering.markerting()
    elif job_type == 'director':
        entering.director()
    elif job_type == 'sales':
        entering.sales()
    else:
        print("Sorry, but we didn't find this type of account, please repeat")
