import openpyxl
wb = openpyxl.load_workbook('fitnessTest.xlsx')
sheet = wb.get_sheet_by_name('FitnessTracker')

while True:
    day = input('What day is it? ' )
    if day.lower() not in ('sunday', 'monday', 'tuesday', 'wednesday',
                           'thursday', 'friday', 'saturday', 'sunday'):
        print('Please enter a valid day of the week')
    else:
        break
while True:    
    week = int(input('What week is it? '))
    if week <= 0 or week > 12:
        print('Please enter a valid week')
    else:
        break
    
weightRow = str((week * 4) - 2)
waistRow = str((week * 4) - 1)
bodyFatRow = str((week * 4))

if day in ('Monday', 'monday'):
    column = 'C'
elif day in ('Tuesday', 'tuesday'):
    column = 'D'
elif day in ('Wednesday', 'wednesday'):
    column = 'E'
elif day in ('Thursday', 'thursday'):
    column = 'F'
elif day in ('Friday', 'friday'):
    column = 'G'
elif day in ('Saturday', 'saturday'):
    column = 'H'
elif day in ('Sunday', 'sunday'):
    column = 'I'
else:
    print('Please enter a valid day')

sheet[column + weightRow] = 58
sheet[column + waistRow] = 65
sheet[column + bodyFatRow] = 100

wb.save('fitnessTest.xlsx')
