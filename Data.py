# -*- coding: utf-8 -*-
"""
Created on Thu May 23 11:51:17 2024

@author: mmorin
"""
from datetime import datetime, timedelta
import pandas as pd
import pyxlsb
import numpy as np
from Service import config

class DataManagement():
    
    def __init__(self):
        self.filepath_CAC40_index = config['filepath']['CAC40_Index.csv']['name']
        self.filepath_CAC40_Stock_Market = config['filepath']['CAC40_Stock_Market.txt']['name']
        self.filepath_CAC40_Closing = config['filepath']['CAC40_closing.xlsb']['name']
        self.filepath_data_reporting = config['filepath']['data_reporting.xlsx']['name']
    
    def __set_companies_features(self, df):
        company_names = df['CompanyName'].unique()
        isins = df['ISIN'].unique()
        tickers = df['Ticker'].unique()
        max_len = max(len(company_names), len(isins), len(tickers))
        company_names = list(company_names) + [None] * (max_len - len(company_names))
        isins = list(isins) + [None] * (max_len - len(isins))
        tickers = list(tickers) + [None] * (max_len - len(tickers))
        df_features = pd.DataFrame({
            'CompanyName': company_names,
            'ISIN': isins,
            'Ticker': tickers
        })
        
        return df_features
    
    def get_reporting_data(self, features=False):
        data_reporting = pd.read_excel(self.filepath_data_reporting)
        df_features = self.__set_companies_features(data_reporting)
        data_reporting = data_reporting[['ISIN', 'Close', 'Date']]
        # Pivot table to have a df with a line per date for each company
        df_data_reporting = data_reporting.pivot_table(index='Date', columns='ISIN', values='Close', aggfunc='first')
        df_data_reporting = df_data_reporting.reset_index()
        # Drop companie(s) with to much missing values
        cols_to_drop = self.__get_columns_todrop(df_data_reporting)
        df_data_reporting = df_data_reporting.drop(columns=cols_to_drop)
        # Drop line(s) with missing value(s)
        df_data_reporting = df_data_reporting.dropna()
        df_data_reporting = df_data_reporting.sort_values(by='Date')
        df_data_reporting.set_index('Date', inplace=True)
        if features==True:
            return df_data_reporting, df_features
        else:
            return df_data_reporting
    
    def get_CAC40_stock_market(self, features=False):
        data = pd.read_csv(self.filepath_CAC40_Stock_Market, sep='\t')
        data['Date'] = pd.to_datetime(data['Date'], format='%d/%m/%Y')
        df_features = self.__set_companies_features(data)
        df_features.to_excel('features_txt.xlsx')
        data = data[['ISIN', 'Close', 'Date']]
        df_CAC40_stock_market = data.pivot_table(index='Date', columns='ISIN', values='Close', aggfunc='first')
        df_CAC40_stock_market = df_CAC40_stock_market.reset_index()
        df_CAC40_stock_market = df_CAC40_stock_market.sort_values(by='Date')
        cols_to_drop = self.__get_columns_todrop(df_CAC40_stock_market)
        df_CAC40_stock_market = df_CAC40_stock_market.drop(columns=cols_to_drop)
        print(df_CAC40_stock_market)
        if features==True:
            return df_CAC40_stock_market, df_features
        else:
            df_CAC40_stock_market
    
    def __set_analysis_period(self, start_date):
        year = start_date.year
        year = year - 3
        start_date_analysis = "01/01/" + str(year)
        start_date_analysis = datetime.strptime(start_date_analysis, '%d/%m/%Y')
        end_date_analysis = start_date
        return start_date_analysis, end_date_analysis
    
    def __get_columns_todrop(self, df):
        periods_number = len(df['Date'])
        missing_value = df.isna().sum()
        cols_to_drop = missing_value[missing_value > (periods_number/4)].index
        print(cols_to_drop)
        return cols_to_drop
    
    def __convert_excel_date(self, excel_date):
        start_date = datetime(1899, 12, 30)
        converted_date = start_date + timedelta(days=excel_date)
        return converted_date
    
    def __convert_dates(self, date_str):
        try:
            excel_date = int(date_str)
            return self.__convert_excel_date(excel_date).strftime('%d/%m/%Y')
        except ValueError:
            return date_str
    
    def get_CAC40_index_historical_price(self, start_date, end_date):
        start_date = pd.to_datetime(start_date, format='%d/%m/%Y')
        end_date = pd.to_datetime(end_date, format='%d/%m/%Y')
        df_historical_price_CAC40 = pd.read_csv(self.filepath_CAC40_index, delimiter=';')
        df_historical_price_CAC40['Date'] = pd.to_datetime(df_historical_price_CAC40['Date'], format='%d/%m/%Y')
        #Arrange CAC40_Index_Data
        df_historical_price_CAC40 = df_historical_price_CAC40.dropna()
        df_historical_price_CAC40 = df_historical_price_CAC40.sort_values(by='Date')
        # Set Date in index
        df_historical_price_CAC40.set_index('Date', inplace=True)
        df_historical_price_CAC40 = df_historical_price_CAC40[['Adj Close']]
        df_historical_price_CAC40 = df_historical_price_CAC40.loc[start_date:end_date]
        
        return df_historical_price_CAC40
    
    def get_CAC40_stock_historical_price(self, features=False):
        df_features = pd.read_excel(self.filepath_CAC40_Closing, sheet_name="Input", engine='pyxlsb')
        df_cac40_stockprice = pd.read_excel(self.filepath_CAC40_Closing, sheet_name="Portfolio", engine='pyxlsb')
        df_cac40_stockprice = df_cac40_stockprice.rename(columns={'Name':'Date'})
        df_cac40_stockprice['Date'] = df_cac40_stockprice['Date'].apply(self.__convert_dates)
        isin_columns = df_cac40_stockprice.iloc[0]
        isin_columns = [item[:-3] for item in isin_columns]
        isin_columns = isin_columns[1:]
        df_cac40_stockprice = df_cac40_stockprice[1:]
        df_cac40_stockprice['Date'] = df_cac40_stockprice['Date'].apply(lambda x: pd.to_datetime(x))
        df_cac40_stockprice = df_cac40_stockprice.sort_values(by='Date')
        df_cac40_stockprice.set_index('Date', inplace=True)
        df_cac40_stockprice.columns = isin_columns
        if features==True:
            return df_features
        else:
            return df_cac40_stockprice
    
    def get_portfolio_reporting_data(self, isin_list, start_date, end_date):
        start_date_report = pd.to_datetime(start_date, format='%d/%m/%Y')
        end_date_report = pd.to_datetime(end_date, format='%d/%m/%Y')
        df_cac40_stockprice = self.get_CAC40_stock_historical_price()
        portfolio_reporting_data = df_cac40_stockprice[isin_list]
        portfolio_reporting_data = portfolio_reporting_data.loc[start_date_report:end_date_report]
        portfolio_reporting_data.dropna()
        return portfolio_reporting_data
    
    def get_portfolio_analysis_data(self, isin_list, start_date):
        start_date_report = pd.to_datetime(start_date, format='%d/%m/%Y')
        start_date_analysis, end_date_analysis = self.__set_analysis_period(start_date_report)
        df_cac40_stockprice = self.get_CAC40_stock_historical_price()
        portfolio_data = df_cac40_stockprice[isin_list]
        portfolio_data = portfolio_data.loc[start_date_analysis:end_date_analysis]
        portfolio_data.dropna()
        return portfolio_data