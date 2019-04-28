"""
Course Paper XLS Parse

Developer: Stanislav Alexandrovich Ermokhin
"""

import pandas as pd
import os
import fnmatch

data_entry = object()

with open('../XLS/countries.txt') as file_obj:
    COUNTRIES = file_obj.readlines()
    COUNTRIES = [obj.rstrip() for obj in COUNTRIES]


def validate(filename,
             obj1, obj2):
    """
    Validation function.
    :param filename: str
    :param obj1: pandas.DataFrame
    :param obj2: pandas.DataFrame
    :return: str
    """

    series = ((obj1 == obj2) == obj1.notna()).values
    expression = False in series

    if not expression:
        return filename + ' IS VALID'

    else:
        return filename + ' IS NOT VALID'


def read_xls(countries=None,
             pattern='*.xls'):
    """
    Read excel docs function.
    :param countries: list
    :param pattern: str
    :param up: str
    :return: list
    """

    countries = countries or COUNTRIES
    library = set()
    xls_lst = list()

    for excel_file in os.listdir('../XLS'):
        if fnmatch.fnmatch(excel_file, pattern):
            library.add(excel_file)

    for excel_file in library:
        xls_lst.append(pd.read_excel('../XLS/'+excel_file,
                                     sheet_name='Data',
                                     header=3,
                                     index_col=1))
    
    data_frames_lst = [obj.loc[countries]
                       for obj in xls_lst]

    return data_frames_lst


def write_excel(df_obj, name='new.xls'):
    """
    Making excel file.
    :param df_obj: pandas.DataFrame
    :param name: str
    :return: None
    """

    df_obj.to_excel(excel_writer=name,
                    sheet_name='Data')


def write_html(df_obj):
    """
    Making html file.
    :param df_obj: pandas.DataFrame
    :return: None
    """

    with open('new.html', 'w') as file:
        file.write(df_obj.to_html())


def population_func(obj,
                    indicator, country, population):
    """
    Transforming population.
    :param obj: list
    :param indicator: str
    :param country: str
    :param population: dict
    :return: list
    """

    if '% of population' in indicator:
        for i in range(len(obj)):
            try:
                obj[i] = round(obj[i]*population[country][i]*0.01)
            except ValueError:
                pass

        indicator = indicator[:indicator.find(' (% of population)')] + \
                    ', total'
    elif 'GDP' in indicator:
        for i in range(len(obj)):
            obj[i] = obj[i]/(10**12)

        indicator = indicator[:indicator.find('US')] + \
                    'trln' + indicator[indicator.find('US'):]

    return [indicator]+obj


def parse_new(file='new.xls'):
    """
    Parsing excel file.
    :param file: str
    :return: pandas.DataFrame
    """

    indicators_dict = dict()
    population = dict()
    columns = ['Indicator Name'] + list(range(1998, 2018))
    xls_df = pd.read_excel(file,
                           sheet_name='Data',
                           index_col=1)
    
    for name in xls_df.iterrows():
        indicator = name[0]
        country = name[1][0]
        data_1998_to_2017_incl = name[1][1:-1]

        if 'Population, total' in indicator:
            population[country] = list(data_1998_to_2017_incl)

    for name in xls_df.iterrows():
        indicator = name[0]
        country = name[1][0]
        data_1998_to_2017_incl = name[1][1:-1]

        lst = population_func(list(data_1998_to_2017_incl),
                              indicator, country, population)

        if country not in indicators_dict:
            indicators_dict[country] = [lst]
        else:
            indicators_dict[country].append(lst)

    indicators_dict = {country: pd.DataFrame(indicators_dict[country],
                                             columns=columns)
                       for country in indicators_dict}

    for country in COUNTRIES:
        indicators_dict[country].index = [country
                                          for _ in range(len(indicators_dict[country]))]

    indicators = pd.DataFrame(index=[],
                              columns=columns)
    for country in indicators_dict:
        indicators = indicators.append(indicators_dict[country])

    return indicators


def main():
    """
    Main function.
    :return: pandas.DataFrame
    """

    global data_entry

    with open('../XLS/indicators.txt') as ind_file:
        indicators_lst = ind_file.readlines()
    indicators_lst = [obj.rstrip() for obj in indicators_lst]
    lst = ['Country Code', 'Indicator Name'] + [str(j)
                                                for j in range(1998,
                                                               2019)]

    df_countries = [obj.filter(items=lst)
                    for obj in read_xls()]
    
    df = pd.DataFrame()
    df = df.append(df_countries)

    data_entry = df.loc[df['Indicator Name'].isin(indicators_lst)].drop_duplicates()

    return data_entry


while True:
    main_obj = main()
    inp = input('Type xls, html or parse_new: ')

    if inp == 'xls':
        write_excel(main_obj)
        filename = 'new.xls'
        new_excel_file = pd.read_excel(filename,
                                       header=0,
                                       index_col=0)
        print(validate(filename, new_excel_file, data_entry))

    elif inp == 'html':
        write_html(main_obj)

    elif inp == 'parse_new':
        indicators = parse_new()
        write_excel(indicators, 'new_new_new.xls')

    else:
        try:
            exit()

        except KeyboardInterrupt:
            pass
