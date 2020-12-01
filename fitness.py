#! /usr/bin/env python3

# getting client stats
import math
import openpyxl

wb = openpyxl.load_workbook('fitness.xlsx')
sheet = wb.get_sheet_by_name('FitnessTracker')

# function that calculates body fat pecentage for men
def maleBodyFatPercentage(waist, neck, height):
    rtn  = 495 / (1.0324 - .19077 * (math.log((waist - neck), 10))
                  + .15456 * (math.log(height, 10))) - 450
    return rtn

# function that calculates body fat pecentage for women
def femaleBodyFatPercentage(weight, height, waist, hip):
    rtn = 495 / (1.29579 - .35004 * (math.log((waist + hip - neck), 10))
                 + .22100 * (math.log(height, 10))) - 450
    return rtn

# function that defines fat mass
def fatMass(bodyfat, weight):
    rtn = bodyfat * weight
    return rtn

# function that defines lean mass
def leanMass(weight, fatmass):
    rtn = weight - fatmass
    return rtn

# function that calculates calorie adjustments based on sex
def sexIdentity(sex):
    if sex in ('Male', 'male', 'm', 'M'):
        return 198
    elif sex in ('Female', 'female', 'f', 'F'):
        return 0

# function that calculates calorie adjustments based on age
def ageAdjustment(age):
    return (age * 3.351)

# function that defines various activity levels
def activityLevel(activity):
    if activity == 1:
        return 1.2
    elif activity == 2:
        return 1.375
    elif activity == 3:
        return 1.55
    elif activity == 4:
        return 1.725
    elif activity == 5:
        return 1.9

activityNumbers = {'Sedentary': 1, 'Lightly active': 2, 'Moderately active': 3,
                   'Very active': 4, 'Extremely active': 5}

def printActivityTable(itemsDict, leftWidth, rightWidth):
    print()
    print(' What is your activity level? '.center(leftWidth + rightWidth, '-'))
    for k, v in itemsDict.items():
        print(k.ljust(leftWidth, '.') + str(v).rjust(rightWidth))

print('First, let\'s get some information.')
sex = input('Are you male or female? ')
height = int(input('Enter your height in cm: '))
weight = float(input('Enter your weight in kg: '))
neck = int(input('Enter your neck (at narrowest): '))

if sex in ('Male', 'male', 'm', 'M'):
    waist = int(input('Enter your waist (at navel): '))
    print()
    bodyfat = maleBodyFatPercentage(waist, neck, height)
    fatmass = fatMass(bodyfat, weight)/100
    
    print('Your bodyfat percentage is: ' + str(round(maleBodyFatPercentage(waist, neck, height),2)) + '%')
    print('Your fat mass is: ' + str(round(fatMass(bodyfat, weight)/100,2)) + 'kg')
    print('Your lean mass is: ' + str(round(leanMass(weight, fatmass),2)) + 'kg\n')
    print('Now let\'s find out your calorie needs.')

elif sex in ('Female', 'female', 'f', 'F'):
    waist = int(input('Enter your waist (at narrowest): '))
    hip = int(input('Enter your hip (at widest): '))
    print()
    bodyfat = femaleBodyFatPercentage(weight, height, waist, hip)
    fatmass = fatMass(bodyfat, weight)/100

    print('Your bodyfat percentage is: ' + str(float(femaleBodyFatPercentage(weight, height, waist, hip))) + '%')
    print('Your fatmass is: ' + str(round(float((fatMass(bodyfat, weight)/100)))) + 'kg')
    print('Your lean mass is: ' + str(round(float(leanMass(weight, fatmass)))) + 'kg')
    print()
    print('Now let\'s find out your calorie needs.')
    print()

age = int(input('How old are you? '))

# mueller equation

mueller = round(((leanMass(weight, fatmass) * 13.587) + ((fatMass(bodyfat, weight)/100) * 9.163)
            + sexIdentity(sex) - ageAdjustment(age) + 674))

printActivityTable(activityNumbers, 35, 5)

activity = int(input('\nType the number associated with your activity: '))

print('\nYour basic metabolic rate is ' + str(mueller) + ' calories.')
print('Your maintenance calorie level is ' + str(int(mueller * activityLevel(activity))) + '.\n')

updateDatabase = input('Do you want to update your database with your measurements? ')
if updateDatabase.lower() in ('yes'):
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

    sheet[column + weightRow] = weight
    sheet[column + waistRow] = waist
    sheet[column + bodyFatRow] = round(bodyfat) / 100

    wb.save('fitness.xlsx')

    print()
    print('Database updated')





