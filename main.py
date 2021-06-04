#Made by Sophia

import XL
from pathlib import Path
from time import sleep
import shutil
def sleep_lines():
    for e in range(73):
        print('-', end='')
        sleep(0.00001)
    print('')
def sleep_words(list):
    for c in list:
        print(f'{c}', end=' ')
        sleep(0.2)
    print('')
title = False
First_Line_labels = []
First_Row_Values = []
First_Column_Values = []
number = 0
while True:
    print('-='*15 + ' XL Creator ' + '=-'*15 + '=')
    print("|" + " "*31 + "By Sofia" + " "*32 + "|")
    print("|" + " " * 14 + "A new way to create repetitive excel files" + " " * 15 + "|")
    ERROR = True
    print("|" + " " * 25 + "Enter 'exit' to close" + " " * 25 + "|")
    print("-" * 73)
    PACKAGE_NAME = input("""Enter the package name please
(if you don't want to create a package enter "pass"): """)
    if PACKAGE_NAME.upper().strip() != "PASS" and PACKAGE_NAME.upper().strip() != "EXIT":
        path = Path(PACKAGE_NAME)
        path.mkdir()
    elif PACKAGE_NAME.upper().strip() == "EXIT":
        break
    sleep_lines()
    XL_NAME = input('Enter the name of the excel file please: ')
    if XL_NAME.upper().strip() == "EXIT":
        break
    sleep_lines()
    XL_NAME += '.xlsx'
    SHEET_NAME = input('Enter the name of the excel sheet please: ')
    if SHEET_NAME.upper().strip() == "EXIT":
        break
    sleep_lines()
    while ERROR:
        try:
            COLUMNS_NUMBER = int(input('How many columns would you like to use? '))
            ERROR = False
        except ValueError:
            sleep_lines()
            print('|Invalid value, please try again. ' + ' ' * 38 + '|')
            print('-' * 73)
    ERROR = True
    sleep_lines()
    while ERROR:
        try:
            LINES_NUMBER = int(input('How many rows would you like to use? '))
            ERROR = False
        except ValueError:
            sleep_lines()
            print('|Invalid value, please try again. ' + ' ' * 38 + '|')
            print('-' * 73)
    sleep_lines()
    ERROR = True
    while ERROR:
        while ERROR:
            TITLE_B = input(f'Do you want titles in the first row of {SHEET_NAME} sheet? [Y/N] ')
            if TITLE_B.upper().strip() != 'Y' and TITLE_B.upper().strip() != 'N' and TITLE_B.upper().strip() != 'EXIT':
                sleep_lines()
                print('|Invalid value, please try again. ' + ' ' * 38 + '|')
                print('-' * 73)
            elif TITLE_B.upper().strip() == "Y":
                First_Line_labels = []
                title = True
                sleep_lines()
                for c in range(1, COLUMNS_NUMBER + 1):
                    First_Line_labels.append(input(f'Please enter the title of the {c}º column: '))
                    sleep_lines()
                ERROR = False
            elif TITLE_B.upper().strip() == 'EXIT':
                break
            else:
                title = False
        sleep_lines()
        rules_step1_phrase1 = ['First',  'step:']
        rules_step1_phrase2 = ['I', 'want', 'to', 'make', 'a', 'sequence', 'according', 'to', 'the', 'values', 'of', 'each', 'row', '[0]']
        rules_step1_phrase3 = ['I', 'want', 'to', 'make', 'a', 'sequence', 'according', 'to', 'the', 'values', 'of', 'each', 'column', '[1]']
        sleep_words(rules_step1_phrase1)
        sleep_words(rules_step1_phrase2)
        sleep_words(rules_step1_phrase3)
        step1 = 8
        while True and step1 != 0 and step1 != 1:
            try:
                step1 = int(input('Your answer [0/1]: '))
                sleep_lines()
                if step1 != 1 and step1 != 0:
                    print('|Invalid value, please try again. ' + ' ' * 38 + '|')
                    print('-' * 73)
                else:
                    break
            except ValueError:
                sleep_lines()
                print('|Invalid value, please try again. ' + ' ' * 38 + '|')
                print('-' * 73)
        rules_step2_phrase1 = ['Second', 'step:']
        rules_step2_phrase2 = ['I want', 'to add', 'a number', 'to a', 'cell according', 'to',
                            'the previous', "cell's", 'value', '(= previous_cell_value + number)', '[0]']
        rules_step2_phrase3 = ['I want', 'to subtract', 'a number', 'to a', 'cell according', 'to',
                                'the previous', "cell's", 'value', '(= previous_cell_value - number)', '[1]']
        rules_step2_phrase4 = ['I want', 'to', 'multiply a', 'number to', 'a cell', 'according to',
                                'the previous', "cell's", 'value', '(= previous_cell_value * number)', '[2]']
        rules_step2_phrase5 = ['I want', 'to divide', 'a', 'number to', 'a cell', 'according to',
                                'the previous', "cell's", 'value', '(= previous_cell_value / number)', '[3]']
        rules_step2_phrase6 = ['I want', 'to raise', 'the value', 'of', 'a cell', 'to a', 'number',
                               'according to', 'the', 'previous', "cell's", 'value', '(= previous_cell_value ^ number)', '[4]']
        rules_step2_phrase7 = ['I want', 'to exponentiate', 'the value', 'of the', 'first cell',
                               '(= first_cell_value ^ x)', '[5]']
        sleep_words(rules_step2_phrase1)
        sleep_words(rules_step2_phrase2)
        sleep_words(rules_step2_phrase3)
        sleep_words(rules_step2_phrase4)
        sleep_words(rules_step2_phrase5)
        sleep_words(rules_step2_phrase6)
        sleep_words(rules_step2_phrase7)
        step2 = 8
        while True and step2 != 0 and step2 != 1 and step2 != 2 and step2 != 3 and step2 != 4 and step2 != 5:
            try:
                step2 = int(input('Your answer [0/1/2/3/4/5]: '))
                sleep_lines()
                if step2 != 1 and step2 != 0 and step2 != 2 and step2 != 3 and step2 != 4 and step2 != 5:
                    print('|Invalid value, please try again. ' + ' ' * 38 + '|')
                    print('-' * 73)
                else:
                    break
            except ValueError:
                sleep_lines()
                print('|Invalid value, please try again. ' + ' ' * 38 + '|')
                print('-' * 73)
        if step2 != 5:
            while True:
                try:
                    number = float(input(f'Please enter the float "number" [your choice was "{step2}"]: '))
                    sleep_lines()
                    break
                except ValueError:
                    sleep_lines()
                    print('|Invalid value, please try again. ' + ' ' * 38 + '|')
                    print('-' * 73)
        ERROR = False
    if TITLE_B.upper().strip() == 'EXIT':
        break
    if step1 == 0:
       if title:
           for c in range(1, LINES_NUMBER):
               while True:
                   try:
                       value = float(input(f'Please set the first value (decimal number) of the {c+1}º row: '))
                       break
                   except ValueError:
                       sleep_lines()
                       print('|Invalid value, please try again. ' + ' ' * 38 + '|')
                       print('-' * 73)
               First_Row_Values.append(value)
       else:
        for c in range(1, LINES_NUMBER + 1):
            while True:
                try:
                    value = float(input(f'Please set the first value (decimal number) of the {c}º row: '))
                    break
                except ValueError:
                    sleep_lines()
                    print('|Invalid value, please try again. ' + ' ' * 38 + '|')
                    print('-' * 73)
            First_Row_Values.append(value)
    else:
        for c in range(1, COLUMNS_NUMBER + 1):
            while True:
                try:
                    value = float(input(f'Please set the first value (decimal number) of the {c}º column: '))
                    break
                except ValueError:
                    sleep_lines()
                    print('|Invalid value, please try again. ' + ' ' * 38 + '|')
                    print('-' * 73)
            First_Column_Values.append(value)
    XL.CREATE_XL_FILE(filename=XL_NAME , sheetname=SHEET_NAME, columns_number=COLUMNS_NUMBER,
    lines_number=LINES_NUMBER, choice=step1, choice2=step2, titles=title, numberx=number, labels=First_Line_labels,
    first_row_val=First_Row_Values, first_column_val=First_Column_Values)
    shutil.move(XL_NAME, PACKAGE_NAME)



