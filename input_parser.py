from openpyxl import load_workbook
from openpyxl.styles import Alignment
from env import *
import json
# import nltk
# from nltk import word_tokenize, sent_tokenize
# from pprint import pprint




wb = load_workbook(filename= filename)
ws = wb.active



def create_new_cols():

    '''Create 4 new columns on the left side of the spreadsheet (cols A - D). Column A is purely a cosmetic
    buffer on the left side of the sheet (width 4). Columns B - D are meant for extra date data (width 12). Also
    delete the header row.'''

    ws.insert_cols(0, 4)
    ws.delete_rows(1)



def format_col_widths():
    '''Set column widths to accept new data and display long strings in Column H (width 80).'''

    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['H'].width = 80



def set_alignment(cols):

    '''Set alignment of columns passed in to 'center'.

    ::param:: cols - tuple of strings representing column names

        ex: ('B', 'C', 'D')

    '''

    for row in range(1,1000):
        if ws[f'E{row}'].value == None:
            break
        for col in (cols):
            ws[f'{col}{row}'].alignment = Alignment(horizontal='center')



def format_dates():

    '''Format date data and insert into newly-created columns. '''

    month_hash = {
        1: 'Jan',
        2: 'Feb',
        3: 'March',
        4: 'April',
        5: 'May',
        6: 'June',
        7: 'July',
        8: 'Aug',
        9: 'Sept',
        10: 'Oct',
        11: 'Nov',
        12: 'Dec'}

    for row in range(1, 1000):
        if ws[f'E{row}'].value == None:
            break
        date = ws[f'E{row}'].value
        ws[f'B{row}'].value = date.year
        ws[f'C{row}'].value = f'{date.year} - {month_hash[date.month]}'
        ws[f'D{row}'].value = 'Long Island'



def combine_debits_credits():

    '''Combine Debit and Credit columns into a single column and preserve sign of values (+ or -).'''

    for row in range(1, 1000):
        if ws[f'E{row}'].value == None:
            break
        amount = ws[f'J{row}'].value
        if amount and ws[f'I{row}'].value == None:
            ws[f'I{row}'].value = amount



def delete_unused_cols():

    '''Delete column J after 'combine_debits_credits()' has been run to move credits values over to
    column I.'''

    ws.delete_cols(10)



def split_cell_values(column):

    '''Split all distinct alpha-numeric sequences in Column H (Description), append the resulting list to a list
    of lists, called 'text', and return it.'''

    text = []  # each line in the spreadsheet is stored here in separate lists
    for row in range(1, 20):
        cell = ws[f'H{row}'].value
        words = [x.lower() for x in cell.split() if not x.isdigit()]  # Only keep continuous string sequences if
        text.append(words)  # they are not entirely numbers
    return text



def words_hash(text):

    '''Creates hash entry (if one doesn't exist) that records the count of each textual sequence used in the column of the
    spreadsheet and returns a list, sorted by most common word.'''

    words_hash = {}  # each word in the collective description column will be counted and hashed here
    for line in text:
        for word in line:
            if word in words_hash:
                words_hash[word] += 1
            else:
                words_hash.update({word: 1})
    sorted_dict = (sorted(words_hash, key=words_hash.get, reverse=True))  # Sort the dictionary by the frequency of usage
    return sorted_dict



def categorize_description(sorted_dict, text):

    '''Handles the mapping of user-defined categories to unique textual sequences by printing out the lines in
    which the sequence occurs and then prompting the user for input. Results are stored in a database 'descriptions.db'
    or a json file 'descriptions.json' for later recall.'''

    try:
        with open(descriptions_filepath) as f:
            descript_hash = json.load(f)

            descript_hash = {}  # Store user-defined categories here
            for word in sorted_dict:
                if word not in descript_hash:
                    print(f'Here are the lines in which \'{word}\' occurs in the spreadsheet.')
                    for line in text:
                        if word in line:
                            print(line)
                    description = input(f'What description do you want to associate with the term \'{word}\' ?')
                    descript_hash.update({word: description})
            with open(descriptions_filepath, 'w') as f:
                json.dump(descript_hash, f)
            return descript_hash
    except:
        print('No previous description file found. Creating new...')

    descript_hash = {}  # Store user-defined categories here
    for word in sorted_dict:
        if word not in descript_hash:
            print(f'Here are the lines in which \'{word}\' occurs in the spreadsheet.')
            for line in text:
                if word in line:
                    print(line)
            description = input(f'What description do you want to associate with the term \'{word}\' ?')
            descript_hash.update({word: description})
    breakpoint()
    with open(descriptions_filepath, 'w') as f:
        json.dump(descript_hash, f)
    return descript_hash




def write_description(descript_hash, text):
    line_ref = 0
    for line in text:
        line_ref += 1
        for word in line:
            if word in descript_hash:
                print(f'Description found for word \'{word}\': \'{description}\'. Writing value to spreadsheet.')
                ws[f'F{line_ref}'].value = descript_hash[word]



def parse_description(column):

    text = split_cell_values(column)
    sorted_dict = words_hash(text)
    descript_hash = categorize_description(sorted_dict, text)
    write_description(descript_hash, text)


    # text = ''
    # for cell in ws['H']:
    #     words = cell.value.split()
    #     for word in words:
    #         if word.isalpha():
    #             text += (word.lower() + ' ')
    # tokens = word_tokenize(text)
    # text_new = nltk.Text(tokens)
    # print(text_new.concordance('deb'))









def format_spreadsheet():
    create_new_cols()
    format_col_widths()
    set_alignment(('B', 'C', 'D'))
    format_dates()
    combine_debits_credits()
    delete_unused_cols()
    parse_description('H')



#############
# Dashboard #
#############


# create_new_cols()
# format_col_widths()
# set_alignment(('B', 'C', 'D'))
# format_dates()
# combine_debits_credits()
# delete_unused_cols()
parse_description('H')

# format_spreadsheet()


wb.save(filename=filename)
wb.close()