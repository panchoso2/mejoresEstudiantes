import xlrd
import xlwt
from xlwt import Workbook

#   name of files
files = ['1A.xls', '1B.xls', '2A.xls', '2B.xls', '3A.xls', '3B.xls']
#   path of files
path = "E:\Descargas Chrome\Excels\\"

#   dictionary
students = {}

#   navigate through files
for i in files:

    filename = path + i

    #   open workbook
    wb = xlrd.open_workbook(filename)
    sheet = wb.sheet_by_index(0)

    #   find prom col (it is different on every file)
    for u in range(sheet.ncols):
        if sheet.cell_value(0,u) == 'Prom.':
            promCol = u

    #   navigate through rows
    for nrow in range(sheet.nrows):

        #   ignore headers
        if nrow == 0:
            continue
        
        studentName = sheet.cell_value(nrow,1)

        #   student exists on dictionary
        if studentName in students:
            students[studentName].append(sheet.cell_value(nrow,promCol))

        #   student didnt exist on dictionary
        else:
            students[studentName] = []
            students[studentName].append(sheet.cell_value(nrow,promCol))


#   eliminate students thar are no longer in this school
actualFiles = ['4A.xls', '4B.xls']
actualStudents = {}

for e in actualFiles:

    filename = path + e

    #   open workbook
    wb = xlrd.open_workbook(filename)
    sheet = wb.sheet_by_index(0)

    #   copy of the dictionary
    copyStudents = students

    for nrow in range(sheet.nrows):

        #   ignore headers
        if nrow == 0:
            continue
        
        actualStudentName = sheet.cell_value(nrow,1)
      
        if actualStudentName in copyStudents:
            actualStudents[actualStudentName] = students[actualStudentName]
            del students[actualStudentName]
    copyStudents = students



#   print(students)
#print("\n")
#print("Estudiantes que estan en 4to actualmente: \n")
#print(actualStudents)
for t in actualStudents:
     print("Nombre: " + t + " Total de notas: " + str(len(actualStudents[t])))


#   generate dictionary of proms
promStudents = {}

for student in actualStudents:
    temp = actualStudents[student]
    #prom = sum(temp)/len(temp)
    sum = 0
    cant = 0
    for u in actualStudents[student]:
        if type(u) is float:
            sum = sum + u
            cant = cant + 1
    prom = sum/cant
    promStudents[student] = round(prom,1)

#print("\n")
#print("Promedio de cada estudiante: \n")
#print(promStudents)


# export results to excel
final = Workbook()
finalSheet = final.add_sheet('Hoja 1')
finalSheet.write(0,0,"Nombre")
finalSheet.write(0,1,"Promedio")

row = 1
for student in promStudents:
    finalSheet.write(row,0,student)
    finalSheet.write(row,1,promStudents[student])
    row = row + 1
final.save("Resultados.xls")