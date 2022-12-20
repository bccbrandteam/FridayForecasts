import os
import pandas
import numpy
from sqlalchemy import create_engine
from selenium import webdriver
class Scraper:
    def __init__(self):

        # Create a virtual display for the Web Driver to run 
        # display = Display(visible=1)
        # display.start()

        # Create an instance of the chromium webdriver
        dir_path = os.path.dirname(os.getcwd())
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument('--start-maximized')

        # Change this to watch the webscrapers work on your screen
        options.add_argument('--headless')

        prefs = {"download.default_directory": os.path.join(dir_path, 'temp')}
        options.add_experimental_option("prefs", prefs)

        # For local computers -- include the location of your chromedriver on your local computer
        # chromedriver = '/Users/haydenthomas/Documents/Hayden/01 School/chromedriver'
        # chromedriver = "C:/Users/jonah/Desktop/BCC/Data/chromedriver/chromedriver.exe"

        # For Raspberry Pi
        # maybe try just /usr/bin? it gives a different error. dont know which error is better tbh
        chromedriver = '/usr/bin/chromedriver'

        os.environ["webdriver.chrome.driver"] = chromedriver
        # self.driver = webdriver.Chrome(options=options, executable_path=chromedriver)
        self.driver = webdriver.Chrome(options=options, executable_path=chromedriver)


def merge_columns(dataframe):
    columns = dataframe.columns.tolist()
    for column in columns:
        column_count = 1
        while True:
            try:
                to_merge = dataframe[column + '.' + str(column_count)].name
            except KeyError:
                break
            dataframe[column] = dataframe[column].fillna(dataframe[to_merge])
            del dataframe[to_merge]
            columns.remove(to_merge)
            column_count += 1


def events_cleanup(dataframe):
    try:
        dataframe['MBA Majors Targeted'] = dataframe['MBA Majors Targeted'].fillna(dataframe['Which graduate business major(s) are you targeting?'])
        del dataframe['Which graduate business major(s) are you targeting?']
        return dataframe
    except KeyError:
        return dataframe


def utc_converter(dataframe, timezone):
    columns = dataframe.columns.tolist()
    for column in columns:
        try:
            s = pandas.to_datetime(dataframe[column], format='%Y-%m-%d %H:%M:%S UTC').dt.tz_localize('UTC')
        except ValueError:
            try:
                s = pandas.to_datetime(dataframe[column], format='%Y-%m-%dT%H:%M:%S.%fZ').dt.tz_localize('UTC')
            except ValueError:
                continue
        s = s.dt.tz_convert(timezone).dt.strftime('%Y/%m/%d %H:%M:%S')
        dataframe[column] = s
    dataframe = dataframe.replace(to_replace=pandas.NaT, value=numpy.nan)
    return dataframe

def tidy_split(df, column, sep='|', keep=False):
    """
    Split the values of a column and expand so the new DataFrame has one split
    value per row. Filters rows where the column is missing.

    Params
    ------
    df : pandas.DataFrame
        dataframe with the column to split and expand
    column : str
        the column to split and expand
    sep : str
        the string used to split the column's values
    keep : bool
        whether to retain the presplit value as it's own row

    Returns
    -------
    pandas.DataFrame
        Returns a dataframe with the same columns as `df`.
    """
    indexes = list()
    new_values = list()
    df = df.dropna(subset=[column])
    for i, presplit in enumerate(df[column].astype(str)):
        values = presplit.split(sep)
        if keep and len(values) > 1:
            indexes.append(i)
            new_values.append(presplit)
        for value in values:
            indexes.append(i)
            new_values.append(value)
    new_df = df.iloc[indexes, :].copy()
    new_df[column] = new_values
    return new_df

def mysqlalchemy(charset):
    user = 'iscareer_operations'
    pw = "BCCRecruiter1!"
    host = '159.65.98.126'
    db = 'iscareer_operations'

    enginestr = 'mysql+pymysql://{}:{}@{}/{}?charset={}'.format(user,pw,host,db,charset)
    alchemyengine = create_engine(enginestr)
    #alchemyengine = create_engine(f"mysql+pymysql://{user}:{pw}@{host}/{db}?charset=utf8")

    return alchemyengine