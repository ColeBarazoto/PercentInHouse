import openpyxl

inHouseWB = openpyxl.load_workbook(r'C:\Users\colebarazoto\Immedia, LLC\Immedia Development Group - Documents\Firmware Updating\Products In-House.xlsx')
inHouse = inHouseWB['Sheet1']

databaseWB = openpyxl.load_workbook(r'C:\Users\colebarazoto\OneDrive - Immedia, LLC\Documents\GitHub\dToolsIPCreate\Read Install Schedule (Database)\Database.xlsx')
database = databaseWB['Database']


# Returns list of projects
def SortProjectsList(refList):
    projects = [refList[0]]

    for proj in refList[1:]:
        if proj not in projects:
            projects.append(proj)

    return projects


if __name__ == "__main__":

    products = inHouse['A']
    jobs = inHouse['B']

    inHouseProducts = []
    inHouseJobs = []

    jobsCalendar = database['A']

    for i in range(len(products)):
        print(products[i].value, '|', jobs[i].value)
        inHouseProducts.append(products[i].value)
        inHouseJobs.append(jobs[i].value)

    myList = SortProjectsList(inHouseJobs)
    print(myList)

    for i in range(len(jobsCalendar)):
        for j in range(len(myList)):
            if jobsCalendar[i].value == myList[j]:
                print(jobsCalendar[i].value)  # Found product in house that has an upcoming install date
