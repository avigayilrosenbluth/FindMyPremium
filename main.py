import openpyxl

# Giving the location of the file
path = '/Users/avigayilrosenbluth/Downloads/IllustrativeLifeTable.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb.active

Type = input('Life or Annuity?')
Age = int(input('How old are you?'))
Pay = int(input('What is your payout?'))
Row = int(Age - 19)
# Selecting the right column
if Type == 'Life' or Type == 'life':
    cell = ws.cell(row=Row, column=5)
    print('Your premium is: ' + str(Pay * cell.value))
elif Type == 'Annuity' or Type == 'annuity':
    cell = ws.cell(row=Row, column=4)
    print('Your premium is: ' + str(Pay * cell.value))
print('Thank you for using Find My Premium!')