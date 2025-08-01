# importing openpyxl, to work with excel files
import openpyxl as xl

# loading in our data
wb = xl.load_workbook('info.xlsx')
sheet = wb['Sheet1']

# it does not matter how we define our cell
cell = sheet['a1']
cell1 = sheet.cell(1, 1)
if cell == cell1:
    print("As we can see, both of our defined cells are equal")
print(f"The value of cell A1 is {cell.value}")
