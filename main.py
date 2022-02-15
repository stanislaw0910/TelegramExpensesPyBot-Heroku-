from flask import Flask, request
from flask_sslify import SSLify
import telebot
import gspread
from datetime import datetime
import logging
import os
from decimal import Decimal as dec
import re
import json


TG_API_TOKEN = os.environ.get('TG_API_TOKEN')
SERVICE_ACCOUNT = os.getenv('SERVICE_ACCOUNT')
SA = json.loads(os.getenv('SA'))
OWNER_ID = int(os.getenv('OWNER_ID'))
EMAIL = os.getenv('EMAIL')
TOKEN = TG_API_TOKEN
bot = telebot.TeleBot(TOKEN)
date_time = datetime(datetime.today().year, datetime.today().month, 1)
gc=gspread.service_account_from_dict(SA)
server = Flask(__name__)


@bot.message_handler(func=lambda message: True, regexp=r'^\w+( )\d{1,8}(\.\d{1,2})?$')
def add_current_month_existing_expense(message):
    """
    This function takes as input initial message that should be in 'expense price' format. Example 'bread 50'
    If expense is not in sheet, function prompts you to enter a category to create a new expense item
    """

    if message.from_user.id == OWNER_ID:
        try:
            expense, expense_value = message.text.split(' ')
            sh = gc.open(date_time.today().strftime("%Y.%m"))
            worksheet = sh.get_worksheet(0)
            try:
                cell = worksheet.find(expense)
                exp_val_cell = worksheet.cell(cell.row, cell.col + 1)  # cell with expense description
                added_value = dec(expense_value)  # expense cost
                exp_val_cell_value = str(dec(exp_val_cell.value.replace(',', '.')) + added_value).replace('.', ',')
                worksheet.update_cell(exp_val_cell.row, exp_val_cell.col, exp_val_cell_value)
                cell_name = gspread.utils.rowcol_to_a1(exp_val_cell.row, exp_val_cell.col)
                worksheet.format(cell_name, {"horizontalAlignment": "RIGHT",
                                             "textFormat": {
                                                 "fontSize": 10
                                             }
                                             })
                bot.send_message(message.chat.id, expense + ' ' + expense_value + ' has been successfully added')
            except AttributeError as ae:
                bot.send_message(message.chat.id, 'You entered the expense that is not in sheet\n')
                msg = bot.send_message(message.chat.id, 'If you want to add a new expense\n'
                                                        'please enter the expense category from message below:'
                                       )
                show_categories(message)
                bot.register_next_step_handler(msg, add_current_month_new_expense_by_category,
                                               expense, expense_value, worksheet)
                logging.error(str(ae))
        except Exception as e:
            if message.text.lower() in ['exit', 'start', 'help']:
                bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                                  'command /start or /help')
                return
            else:
                bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
                bot.send_message(message.chat.id, 'Please check if you have access and internet connection is stable: ')
                logging.error(str(e))
                return
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def add_current_month_new_expense_by_category(message, expense, expense_value, worksheet):
    """
    If the category exists
    This function takes name of category from message and adds expense in new cell above the last not empty row
    """
    try:
        cell = worksheet.find(message.text)
        values_list = worksheet.col_values(cell.col)
        cell = worksheet.cell(len(values_list) + 1, cell.col)
        exp_val_cell = worksheet.cell(cell.row, cell.col + 1)
        added_value = dec(expense_value)
        exp_val_cell_value = str(added_value).replace('.', ',')
        worksheet.update_cell(exp_val_cell.row, exp_val_cell.col, exp_val_cell_value)
        worksheet.update_cell(cell.row, cell.col, expense)
        bot.send_message(message.chat.id, expense + ' was successfully added to ' + message.text)
    except AttributeError as ae:
        msg = bot.send_message(message.chat.id, "You entered wrong category's name")
        bot.register_next_step_handler(msg, add_current_month_new_expense_by_category,
                                       expense, expense_value, worksheet)
        logging.error(str(ae))
    except Exception as e:
        if message.text.lower() in ['exit', 'start', 'help']:
            bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                              'command /start or /help')
            return
        else:
            bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
            bot.send_message(message.chat.id, 'Please check if you have access and internet connection is stable: ')
            logging.error(str(e))
            return


@bot.message_handler(commands=['start', 'Start', 'help', 'Help'])
def handle_start_help(message):
    """
    Shows start menu with all supported commands.
    """
    if message.from_user.id == OWNER_ID:
        bot.send_message(message.chat.id, "All available commands:\n"
                                          "/start or /help shows help menu\n"
                                          "/AddExpenseToDefinedMonth or /AEDM\n adds expense to defined month\n"
                                          "/CreateSpreadsheet or /CS creating spreadsheet for current month\n"
                                          "/CurrentMonthBalance or /CMB\n shows current month balance\n"
                                          "/CurrentMonthExpenseByCategory\n retrieves expenses for current month for "
                                          "every category\n"
                                          "/DefinedMonthBalance or /DMB\n shows defined month balance\n"
                                          "/DefinedMonthExpenseByCategory\n retrieves expenses for defined month for "
                                          "every "
                                          "category\n"
                                          "/FormatDefinedFile or /FDF\n restores correct formating for whole document"
                                          " defined by date, takes up to 5 minutes, do not use frequently\n"
                                          "/ShowCategories or /SC shows existing categories in current month \n"
                                          "/ShowExpenses or /SE show expenses in defined category of current month\n"

                         )
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['CreateSpreadsheet', 'CS'])
def create_spreadsheet(message):
    """
    This function creates new spreadsheet for current month if it doesn't exist
    """
    if message.from_user.id == OWNER_ID:
        try:
            gc.open(date_time.today().strftime("%Y.%m"))
            bot.send_message(message.chat.id, "Sheet for current month is already exists")
        except gspread.SpreadsheetNotFound:
            sh=gc.create(date_time.today().strftime("%Y.%m"))
            sh.share(EMAIL, perm_type='user', role='owner')
            bot.send_message(message.chat.id, date_time.today().strftime("%Y.%m") +
                             " spreadsheet was successfully created")
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['ShowCategories', 'SC'])
def show_categories(message):
    if message.from_user.id == OWNER_ID:
        sh = gc.open(date_time.today().strftime("%Y.%m"))
        worksheet = sh.get_worksheet(0)
        expenses_list = worksheet.row_values(1)[:-6:2]  # make list with no total values such as balance, income, total
        expenses_list = [x for x in expenses_list if x]  # getting rid of empty cells
        bot.send_message(message.chat.id, "\n".join([s for s in expenses_list]))
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['ShowExpenses', 'SE'])
def show_expenses(message):
    """
    This function is 1/2 step process of showing current month's expenses
    By every category
    Executes with command /ShowExpenses or /SE
    """
    if message.from_user.id == OWNER_ID:
        error_column = 0
        smile = u"\uE056"
        try:
            sh=gc.open(date_time.today().strftime("%Y.%m"))
            worksheet=sh.get_worksheet(0)
            for i in range(1, len(worksheet.row_values(1)) - 5, 2):
                exp_cell = worksheet.cell(1, i)
                price_cell = worksheet.cell(1, i+1)
                error_column = i
                if exp_cell.value:
                    expenses_list=worksheet.col_values(exp_cell.col)[1:]
                    values_list=worksheet.col_values(exp_cell.col + 1)[1:]
                    exps_vals_list=list(zip(expenses_list, values_list))
                    exps_vals_list=list(map(' -- '.join, exps_vals_list))
                    bot.send_message(message.chat.id, exp_cell.value + ' -- ' + price_cell.value)
                    bot.send_message(message.chat.id, "\n".join([s for s in exps_vals_list]))
            bot.send_message(message.chat.id, "That's All Folks!"+smile)
        except TypeError as te:
            bot.send_message(message.chat.id, "You have expense with no price or vice versa\n"
                                              "Make some changes with " + str(error_column) +
                                              "th or " + str(error_column+1) +
                                              "th column\n And you will not see this message again")
            logging.error(str(te))
            return
        except Exception as e:
            if message.text.lower() in ['exit', 'start', 'help']:
                bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                                  'command /start or /help')
                return
            else:
                msg=bot.send_message(message.chat.id, 'ERROR! Something went wrong! Please try again!')
                bot.register_next_step_handler(msg, show_expenses)
                logging.error(str(e))
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['CurrentMonthBalance', 'currentmonthbalance', 'cmb', 'CMB'])
def current_month_balance(message):
    """
    Sending 3 messages back to chat with balance statistics for current month.
    income, expenses, and balance
    executes with command /CurrentMonthBalance
    """
    if message.from_user.id == OWNER_ID:
        sh = gc.open(date_time.today().strftime("%Y.%m"))
        worksheet = sh.get_worksheet(0)
        bot.send_message(message.chat.id, "Current month income is: " + worksheet.acell('X1').value)
        bot.send_message(message.chat.id, "Current month expenses are: " + worksheet.acell('V1').value)
        bot.send_message(message.chat.id, "Current month balance is: " + worksheet.acell('Z1').value)
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['DefinedMonthBalance', 'definedmonthbalance', 'DMB', 'dmb'])
def defined_month_balance(message):
    """
    Sending 3 messages back with balance statistics for defined month.
    income, expenses, and balance
    This function is 1/2 step process of returning to chat balance statistics for defined month.
    executes with command /DefinedMonthBalance or /definedmonthbalance or /DMB or /dmb
    """
    if message.from_user.id == OWNER_ID:
        msg = bot.reply_to(message, 'Enter year and month in format YYYY.MM: ')
        bot.register_next_step_handler(msg, defined_month_balance_input)
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def defined_month_balance_input(message):
    """
    This function is 2/2 step process of returning to chat balance statistics for defined month.
    """
    try:
        sh = gc.open(message.text)
        worksheet = sh.get_worksheet(0)
        bot.send_message(message.chat.id, message.text + " month income is: " + worksheet.acell('X1').value)
        bot.send_message(message.chat.id, message.text + " month expenses are: " + worksheet.acell('V1').value)
        bot.send_message(message.chat.id, message.text + " month balance is: " + worksheet.acell('Z1').value)
    except Exception as e:
        if message.text.lower() in ['exit', 'start', 'help']:
            bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                              'command /start or /help')
            return
        else:
            bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nPlease try again!')
            msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
            bot.register_next_step_handler(msg, defined_month_balance_input)
            logging.error(str(e))


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['CurrentMonthExpenseByCategory', 'currentmonthexpensebycategory'])
def current_month_expense_by_category(message):
    """
    Returns message with balance statistics for current month expenses by every category.
    income, expenses, and balance
    executes with command /CurrentMonthExpenseByCategory
    """
    if message.from_user.id == OWNER_ID:
        sh = gc.open(date_time.today().strftime("%Y.%m"))
        worksheet = sh.get_worksheet(0)
        categories_list = worksheet.row_values(1)[:-6:2]
        values_list = worksheet.row_values(1)[1:-6:2]
        cats_vals_list = list(zip(categories_list, values_list))
        cats_vals_list = list(map(' -- '.join, cats_vals_list))
        cats_vals_list = filter(lambda s: len(s)>4, cats_vals_list)
        bot.send_message(message.chat.id, "Total by every category:\n" + "\n".join([s for s in cats_vals_list]))
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


# noinspection SpellCheckingInspection
@bot.message_handler(commands=['AddExpenseToDefinedMonth', 'addexpensetodefinedmonth', 'AEDM' 'aedm'])
def add_defined_month_expense(message):
    """
    Adds expense for defined month and for defined category.
    This function is 1/3 step process of adding expense to defined month
    executes with command /AddExpenseToDefinedMonth or /addexpensetodefinedmonth or /AEDM or /aedm
    Asks user to input year and month of document which will be used to add expense
    """
    if message.from_user.id == OWNER_ID:
        msg = bot.send_message(message.chat.id, 'Enter year and month in format YYYY.MM: ')
        bot.register_next_step_handler(msg, defined_month_expense_date)
    else:
        bot.send_message(message.chat.id, "Access denied!!!\nPlease ensure you have right to use this bot!")


def defined_month_expense_date(message):
    """
    This function is 2/3 step process of adding expense to defined month
    On this step it checks if year and month were entered correctly
    And if it does asks to input expense and cost
    """
    month=message.text
    try:
        bot.send_message(message.chat.id, 'You have entered: ' + month)
        sh = gc.open(month)
        worksheet = sh.get_worksheet(0)
        msg = bot.send_message(message.chat.id, "Input expense and cost in 'EXPENSE COST' format")
        bot.register_next_step_handler(msg, add_defined_month_existing_expense, worksheet)
    except gspread.SpreadsheetNotFound as fe:
        bot.reply_to(message, 'ERROR!\nFile ' + str(message.text) + ' not found!\nTry once more!')
        msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
        bot.register_next_step_handler(msg, defined_month_expense_date)
        logging.error(str(fe))
    except Exception as e:
        bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry again!')
        msg = bot.reply_to(message, 'Please choose existing document by correct date input: ')
        bot.register_next_step_handler(msg, defined_month_expense_date)
        logging.error(str(e))


def add_defined_month_existing_expense(message, worksheet):
    """
    This function 3/3 step process of adding expense to defined month
    It takes message with expense and price and trying to add them to sheet
    Checking if price have correct format and if expense is in sheet
    If expense is not on sheet it shows categories to choose and goes to the next additional step
    """
    try:
        expense, expense_value = message.text.split(' ')
        if re.match(r'^\d+((\.|,)\d{1,2})?$', expense_value):  # this if expession uses regex for price validation
            try:
                cell = worksheet.find(expense)
                exp_val_cell = worksheet.cell(cell.row, cell.col + 1)  # cell with expense description
                added_value = dec(expense_value)  # expense cost
                if exp_val_cell.value:
                    exp_val_cell_value = str(dec(exp_val_cell.value.replace(',', '.')) + added_value).replace('.', ',')
                else:
                    exp_val_cell_value = str(added_value).replace('.', ',')
                worksheet.update_cell(exp_val_cell.row, exp_val_cell.col, exp_val_cell_value)
                bot.send_message(message.chat.id, expense + ' ' + expense_value + ' has been successfully added')
            except AttributeError as ae:
                bot.send_message(message.chat.id, 'You entered the expense that is not in sheet\n')
                msg=bot.send_message(message.chat.id, 'If you want to add a new expense\n'
                                                      'please choose the expense category number from message above:'
                                     )
                show_categories(message)
                bot.register_next_step_handler(msg, add_defined_month_new_expense_by_category,
                                               expense, expense_value, worksheet
                                               )
                logging.error(str(ae))
        else:
            msg = bot.send_message(message.chat.id, "Wrong price format!\n Use only digits separated by dot or comma")
            bot.register_next_step_handler(msg, add_defined_month_existing_expense, worksheet)
    except Exception as e:
        bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
        msg=bot.reply_to(message, 'Try again! ')
        bot.register_next_step_handler(msg, add_defined_month_existing_expense, worksheet)
        logging.error(str(e))


def add_defined_month_new_expense_by_category(message, expense, expense_value, worksheet):
    """
    In this additional step function takes expense, expense_value
    And adds it to the bottom of category's column
    """
    if message.text.lower() in ['exit', 'start', 'help']:
        bot.send_message(message.chat.id, 'Exit initiated!!!\n You can start from beginning with '
                                                  'command /start or /help')
        return
    else:
        try:
            cell=worksheet.find(message.text)
            values_list=worksheet.col_values(cell.col)
            cell=worksheet.cell(len(values_list) + 1, cell.col)
            exp_val_cell=worksheet.cell(cell.row, cell.col + 1)
            added_value=dec(expense_value)
            exp_val_cell_value=str(added_value).replace('.', ',')
            worksheet.update_cell(exp_val_cell.row, exp_val_cell.col, exp_val_cell_value)
            worksheet.update_cell(cell.row, cell.col, expense)
            bot.send_message(message.chat.id, expense + ' was successfully added to ' + message.text)
        except AttributeError as ae:
            msg = bot.send_message(message.chat.id, "You entered wrong category's name")
            bot.register_next_step_handler(msg, add_defined_month_new_expense_by_category,
                                           expense, expense_value, worksheet)
            logging.error(str(ae))

        except Exception as e:
            bot.reply_to(message, 'ERROR!\nSomething wen\'t wrong !''\nTry once more!')
            bot.send_message(message.chat.id, 'Please check if you have access and internet connection is stable: ')
            logging.error(str(e))
            return


@server.route('/' + TOKEN, methods=['POST'])
def getMessage():
    json_string = request.get_data().decode('utf-8')
    update = telebot.types.Update.de_json(json_string)
    bot.process_new_updates([update])
    return "!", 200


@server.route("/")
def webhook():
    bot.remove_webhook()
    bot.set_webhook(url='https://salty-woodland-68506.herokuapp.com/' + TOKEN)
    return "!", 200


# Enable saving next step handlers to file "./.handlers-saves/step.save".
# Delay=2 means that after any change in next step handlers (e.g. calling register_next_step_handler())
# saving will happen after delay 2 seconds.
bot.enable_save_next_step_handlers(delay=2, filename="./step.save")

# Load next_step_handlers from save file (default "./.handlers-saves/step.save")
# WARNING It will work only if enable_save_next_step_handlers was called!
bot.load_next_step_handlers(filename="./step.save")

if __name__ == "__main__":
    server.run(host="0.0.0.0", port=int(os.environ.get('PORT', 5000)))
