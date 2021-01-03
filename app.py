import os
import pandas as pd
import numpy as np
from pprint import pprint


def load_excel(path):
    df_data = pd.read_excel(path, sheet_name=None)

    df_instruments = df_data['INSTRUMENTS']
    df_internal_data = df_data['INTERNAL DATA']
    df_reminder_list = df_data['REMINDER LIST']
    df_comments = df_data['Comments']
    df_cbb = df_data['CBB availability']
    return df_instruments, df_internal_data, df_reminder_list, df_comments, df_cbb


comments = [
    'Already on reminder list',
    'No alternative source was found',
    'Stale ok - confirmed with another source where price is not stale',
    'Stale ok - confirmed with another source where price is stale'
]


def add_comment(isin, comment, df_instruments):
    df_instruments.loc[df_instruments['ISIN'] == isin, 'COMMENT'] = comment


def add_other_comment(isin, comment, df_instruments):
    df_instruments.loc[df_instruments['ISIN'] == isin, 'ADDITIONAL COMMENT'] = comment


def check_if_on_reminder_list(source, reminder_list):
    for idx, row in source.iterrows():
        if len(reminder_list[reminder_list['ISIN'] == row['ISIN']]):
            # print(comment, row['ISIN'])
            # df_instruments.loc[df_instruments['ISIN'] == row['ISIN'], 'COMMENT'] = comment
            add_comment(row['ISIN'], comments[0], source)


def check_same_price_within_6_days(source, internal_data, date):
    for idx, row in source.iterrows():

        results = internal_data[internal_data['ISIN'] == row['ISIN']]

        if len(results) < 6 and pd.isnull(row['COMMENT']):
            print(0, 'No alternative source was found', row['ISIN'])
            add_comment(row['ISIN'], comments[1], source)

        if len(results) == 6 and pd.isnull(row['COMMENT']):
            if row['SOURCE'] == results['BOERSE'].values[0]:
                print(1, 'No alternative source was found', row['ISIN'], results['BOERSE'].values[0])
                add_comment(row['ISIN'], comments[1], source)

        if len(results) == 6 and pd.isnull(row['COMMENT']):
            if row['SOURCE'] != results['BOERSE'].values[0]:

                threshold = abs(row['PRICE'] - results['PRICE'].values[0]) / results['PRICE'].values[0] * 100
                if threshold < 1:
                    if len(results['PRICE'].unique()) == 1:
                        print(2, 'Stale ok - confirmed with another source where price is stale', row['ISIN'],
                              row['SOURCE'])
                        add_comment(row['ISIN'], comments[3], source)
                        add_other_comment(row['ISIN'], row['SOURCE'], source)
                    else:
                        print(3, 'Stale ok - confirmed with another source where price is not stale', row['ISIN'],
                              results['BOERSE'].values[0])
                        add_comment(row['ISIN'], comments[2], source)
                        add_other_comment(row['ISIN'], results['BOERSE'].values[0], source)
                else:
                    print(4, 'No alternative source was found', row['ISIN'], results['BOERSE'].values[0])
                    add_comment(row['ISIN'], comments[1], source)

        if len(results) > 6 and pd.isnull(row['COMMENT']):
            sources = results['BOERSE'].unique()
            temp_sources_no_stale = []
            temp_sources_stale = []

            for boerse in sources:
                if len(results.loc[results['BOERSE'] == boerse, 'PRICE'].unique()) > 1:
                    temp_sources_no_stale.append(boerse)
                elif len(results.loc[results['BOERSE'] == boerse, 'PRICE'].unique()) == 1:
                    temp_sources_stale.append(boerse)

            if len(temp_sources_no_stale) >= 1:
                print(5, 'Stale ok - confirmed with another source where price is not stale', row['ISIN'],
                      temp_sources_no_stale)
                add_comment(row['ISIN'], comments[2], source)
                add_other_comment(row['ISIN'], ', '.join(temp_sources_no_stale), source)

            elif len(temp_sources_no_stale) == 0 and len(temp_sources_stale) >= 1:
                print(6, 'Stale ok - confirmed with another source where price is  stale', row['ISIN'],
                      temp_sources_stale)
                add_comment(row['ISIN'], comments[3], source)
                add_other_comment(row['ISIN'], ', '.join(temp_sources_stale), source)


def save_results(df_instruments, df_internal_data, df_reminder_list, df_comments, df_cbb, path):
    path = path.replace('.xlsx', '_results.xlsx')
    with pd.ExcelWriter(path) as writer:
        df_instruments.to_excel(writer, sheet_name='INSTRUMENTS')
        df_internal_data.to_excel(writer, sheet_name='INTERNAL DATA')
        df_reminder_list.to_excel(writer, sheet_name='REMINDER LIST')
        df_comments.to_excel(writer, sheet_name='Comments')
        df_cbb.to_excel(writer, sheet_name='CBB availability')



