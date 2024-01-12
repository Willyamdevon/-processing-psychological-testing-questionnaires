# Import `xlwt` 
import xlwt

# Initialize a workbook 
book = xlwt.Workbook(encoding="utf-8")

# Add a sheet to the workbook 
sheet = book.add_sheet("Итоги Лечбное-дело Юсупов") 

ans =[[5, 2, 14, 9, 11, 7], [7, 4, 6, 9, 13, 9], [6, 1, 4, 4, 10, 4], [4, 3, 7, 8, 6, 6], [7, 3, 6, 10, 10, 9], [4, 5, 3, 8, 9, 7], [6, 4, 10, 8, 4, 9], [5, 2, 6, 5, 8, 6], [8, 2, 5, 11, 12, 10], [3, 2, 9, 7, 9, 8], [5, 2, 5, 10, 11, 8], [11, 4, 13, 12, 13, 7], [8, 3, 11, 10, 9, 7], [5, 1, 13, 10, 6, 10], [8, 4, 7, 13, 10, 6]]

for i in range(len(ans)):
    lst = ans[i]
    for j in range(len(lst)):
        sheet.write(j+1, i+1, lst[j])


# Save the workbook 
book.save("resul.xls")