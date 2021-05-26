import time
import pandas as pd
import pandas_datareader as web
import datetime as dt
import os


def setup():
    filename_excel = 'predictions.xlsx'
    sheet_update = 'Sheet4'
    date = '2021-05-25'
    return filename_excel, sheet_update, date


def get_data_from_excel(filename_excel, sheet_update):
    data_excel = pd.read_excel(filename_excel, sheet_update)
    return data_excel


def timer_start():
    start_tt = time.time()
    return start_tt


def timer_tt(start):
    tt = time.time() - start
    print(f'tt:{round(tt,2)} sec')
    return tt


def sort_data_excel(data_excel):
    list_tickets = data_excel['Ticket'].values
    list_price_before = data_excel['Price before'].values
    list_price_prediction = data_excel['Price prediction'].values
    # list_price_after = data_excel['Price after'].values
    list_percents_prediction = data_excel['Percentage 1'].values
    # list_percents_after = data_excel['%(17 - predict)'].values
    # list_result_of_predictions = data_excel['Result'].values
    return list_tickets, list_price_before, list_price_prediction, list_percents_prediction


def did_price_go_that_way(list_percents_prediction, list_percentage_after):
    i = 0
    counter_good_prediction = 0
    list_result_of_predict = []
    while i < len(list_percents_prediction):
        if list_percents_prediction[i][0] == '+' and list_percentage_after[i][0] == '+':
            list_result_of_predict.append('Predicted')
            counter_good_prediction += 1
        elif list_percents_prediction[i][0] == '-' and list_percentage_after[i][0] == '-':
            list_result_of_predict.append('Predicted')
            counter_good_prediction += 1
        else:
            list_result_of_predict.append('Not predicted')
        i += 1

    percent_of_success = '{:.2%}'.format(counter_good_prediction / len(list_percents_prediction))
    print(f'percent_of_success: {percent_of_success}')
    return list_result_of_predict, percent_of_success


def get_data_after(list_tickets, date):
    list_price_close_after = []
    list_price_high_after = []
    list_price_low_after = []
    counter = 0
    for company in list_tickets:
        data = web.DataReader(company, 'yahoo', (dt.datetime.now() - dt.timedelta(days=1)), dt.datetime.now())
        price_after = round(data['Close'][date], 2)
        list_price_close_after.append(price_after)
        price_high = round(data['High'][date], 2)
        list_price_high_after.append(price_high)
        price_low = round(data['Low'][date], 2)
        list_price_low_after.append(price_low)
        counter += 1
        print(f'№ {counter} {company}')
    return list_price_close_after, list_price_high_after, list_price_low_after


def percentage_after(list_price_close_after, list_price_before):
    i = 0
    list_percentage_after = []
    while i < len(list_price_before):
        percent = (list_price_close_after[i] - list_price_before[i]) / list_price_close_after[i]
        percent = '{0:+.2%}'.format(percent)
        list_percentage_after.append(percent)
        i += 1
    return list_percentage_after


def prediction_mistake(list_percents_prediction, list_percentage_after):
    i = 0
    list_mistakes_float = []
    list_mistakes = []
    while i < len(list_percents_prediction):
        percent_prediction = float(list_percents_prediction[i][:5])
        percent_after = float(list_percentage_after[i][:5])
        mistake = round(abs(percent_after - percent_prediction), 2)
        list_mistakes_float.append(mistake)
        i += 1

    total_mistake = round(sum(list_mistakes_float), 2)
    print(f'total mistake: {total_mistake}')

    for mistake in list_mistakes_float:
        mistake_percent = str(mistake) + '%'
        list_mistakes.append(mistake_percent)

    return list_mistakes, total_mistake


def df_to_append(list_price_close_after, list_percentage_after, list_result_of_predict, list_mistakes, list_price_high_after, list_price_low_after, percent_of_success, total_mistake):
    df1 = pd.DataFrame({  # Data frame with all arrays same lengths of columns
        'Price after': list_price_close_after,
        'Percentage 2': list_percentage_after,
        'Result': list_result_of_predict,
        'Mistake': list_mistakes,
        'High': list_price_high_after,
        'Low': list_price_low_after,
    })

    df2 = pd.DataFrame({  # Data frame with 1 length of column
        'Success %': [percent_of_success],
        'Total mistake': [total_mistake]
    })

    df_append = pd.concat([df1, df2], axis=1)
    return df_append


def excel_append_to_ex_sheet(filename_excel, df_append, sheet_update):
    to_update = {sheet_update: df_append}

    excel_reader = pd.ExcelFile(filename_excel)
    filename_excel_after = filename_excel[:-5] + '1' + '.xlsx'
    writer = pd.ExcelWriter(filename_excel_after)

    for sheet in excel_reader.sheet_names:
        sheet_df = excel_reader.parse(sheet)
        append_df = to_update.get(sheet)

        if append_df is not None:
            sheet_df = pd.concat([sheet_df, append_df], axis=1)
        sheet_df.to_excel(writer, sheet, index=False)
    writer.save()
    return filename_excel_after


def del_and_rename(filename_excel, filename_excel_after):
    os.remove(filename_excel)
    os.rename(filename_excel_after, filename_excel)


if __name__ == '__main__':
    start_tt = timer_start()
    filename_excel, sheet_update, date = setup()
    data_excel = get_data_from_excel(filename_excel, sheet_update)

    ''' Put excel data and take needed lists from there '''
    list_tickets, list_price_before, list_price_prediction, list_percents_prediction = sort_data_excel(data_excel)

    ''' Get prices after. Close, high and low price '''
    list_price_close_after, list_price_high_after, list_price_low_after = get_data_after(list_tickets, date)

    ''' Calculate percentage of fact price changed '''
    list_percentage_after = percentage_after(list_price_close_after, list_price_before)

    ''' If price went to side which was predicted:Predicted.In reverse: Not predicted. Than count percent of success '''
    list_result_of_predict, percent_of_success = did_price_go_that_way(list_percents_prediction, list_percentage_after)

    ''' Calculate mistake between prediction and fact of price change. Than find summary of mistakes '''
    list_mistakes, total_mistake = prediction_mistake(list_percents_prediction, list_percentage_after)

    ''' Create data frame for append to existing sheet '''
    df_append = df_to_append(list_price_close_after, list_percentage_after, list_result_of_predict, list_mistakes,
                             list_price_high_after, list_price_low_after, percent_of_success, total_mistake)

    ''' Append created data frame to existing sheet chosen in setup() '''
    filename_excel_after = excel_append_to_ex_sheet(filename_excel, df_append, sheet_update)

    ''' Delete file before and rename file after to name file before (Like didn't create any file just append data) '''
    del_and_rename(filename_excel, filename_excel_after)

    tt = timer_tt(start_tt)
