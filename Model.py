# -*- coding: utf-8 -*-
"""
Created on Thu May 23 11:52:26 2024

@author: mmorin
"""
import numpy as np
import pandas as pd
from scipy.stats import norm
from Data import DataManagement
from Service import config

class PortfolioManagement(DataManagement):
    
    def __init__(self):
        super().__init__()
    
    def __daily_return(self, df):
        daily_returns = df.pct_change()
        daily_returns = daily_returns.dropna()
        return daily_returns
    
    def __monte_carlo_simulation(self, daily_returns):
        numofassets = len(daily_returns.columns)
        num_simulations = int(config['parameters']['num_simulations']['name'])
        rf = int(config['parameters']['risk_free_rate']['name'])
        weights_list, list_mu, list_sigma, list_sharpe_ratio = [], [], [], []
        for i in range(num_simulations):
            # Generation of random weights for each asset
            weights = np.random.rand(numofassets)
            weights = weights / np.sum(weights)
            wts_list = weights.tolist()
            weights = weights[:,np.newaxis]
            # Mean calculation
            portfolio_returns = np.dot(daily_returns, weights)
            mu_portfolio = np.mean(portfolio_returns)
            # Volatility calculation
            cov_matrix = np.cov(daily_returns, rowvar=False)
            sigma_portfolio = np.sqrt(np.dot(weights.T, np.dot(cov_matrix, weights))) 
            # Sharpe Ratio calculation
            sharpe_ratio = (mu_portfolio - rf) / sigma_portfolio
            # Addition of calculations to the corresponding lists
            weights_list.append(wts_list)
            list_mu.append(mu_portfolio)
            list_sigma.append(sigma_portfolio [0][0])
            list_sharpe_ratio.append(sharpe_ratio [0][0])
        
        return weights_list, list_mu, list_sigma, list_sharpe_ratio
    
    def __get_monte_carlo_simulation_results(self, daily_returns):
        weights_list, list_mu, list_sigma, list_sharpe_ratio = self.__monte_carlo_simulation(daily_returns)
        # Insertion of monte carlo simulation results in a dataframe
        df_simulations = pd.DataFrame(weights_list, columns=daily_returns.columns)
        df_simulations['Yield'] = list_mu
        df_simulations['Volatility'] = list_sigma
        df_simulations['Sharpe_Ratio'] = list_sharpe_ratio
        return df_simulations
    
    def __get_max_sharpe_ratio_portfolio(self, daily_returns):
        df_simulations = self.__get_monte_carlo_simulation_results(daily_returns)
        # find the maximum sharpe ratio
        max_sharpe_ratio = df_simulations['Sharpe_Ratio'].max()
        index_max_sharpe_ratio = df_simulations.loc[df_simulations['Sharpe_Ratio'] == max_sharpe_ratio].index[0]
        optimal_portfolio = df_simulations.iloc[index_max_sharpe_ratio]
        
        return optimal_portfolio
    
    def __get_best_portfolio(self, isin_list, start_date):
        portfolio_data = super().get_portfolio_analysis_data(isin_list, start_date)
        daily_returns = self.__daily_return(portfolio_data)
        optimal_portfolio = self.__get_max_sharpe_ratio_portfolio(daily_returns)
        return optimal_portfolio
    
    def __get_optimal_weights(self, isin_list, start_date):
        optimal_portfolio = self.__get_best_portfolio(isin_list, start_date)
        optimal_weights = optimal_portfolio.drop(['Yield', 'Volatility', 'Sharpe_Ratio'])
        return optimal_weights
    
    def __compute_portfolio_daily_return(self, isin_list, start_date, end_date):
        df_portfolio = pd.DataFrame()
        portfolio_reporting_data = super().get_portfolio_reporting_data(isin_list, start_date, end_date)
        optimal_weights = self.__get_optimal_weights(isin_list, start_date)
        df_portfolio['VL'] = (portfolio_reporting_data * optimal_weights).sum(axis=1)
        df_portfolio['Return'] = (df_portfolio['VL'] - df_portfolio['VL'].shift(1)) / df_portfolio['VL'].shift(1)
        df_portfolio.fillna(0)
        return df_portfolio, optimal_weights
    
    def __get_cumulative_return(self, isin_list, start_date, end_date):
        portfolio_reporting_data = super().get_portfolio_reporting_data(isin_list, start_date, end_date)
        daily_returns = self.__daily_return(portfolio_reporting_data)
        cumulative_returns = (1 + daily_returns).cumprod()
        return cumulative_returns
    
    def __compute_Value_at_Risk(self, daily_returns, weights):
        confidence_level = float(config['VaR']['confidence_level']['name'])
        num_days = int(config['VaR']['num_days']['name'])
        # Calculate portfolio parameters mu (mean), sigma (standart deviation)
        weights = np.array(weights)[:,np.newaxis]
        portfolio_returns = np.dot(daily_returns, weights)
        mu_portfolio = np.mean(portfolio_returns)
        cov_matrix = np.cov(daily_returns, rowvar=False)
        sigma_portfolio = np.sqrt(np.dot(weights.T, np.dot(cov_matrix, weights)))
        # Calculation of the portfolio’s CV at 99% for a 1-month horizon
        VaR_99_10d = norm.ppf(confidence_level, mu_portfolio, sigma_portfolio) * np.sqrt(num_days)
        CVaR_99_10d = (mu_portfolio - (sigma_portfolio * norm.pdf(norm.ppf(confidence_level)) / (confidence_level))) * np.sqrt(num_days) 
        VaR_99_10d, CVaR_99_10d = VaR_99_10d[0][0], CVaR_99_10d[0][0]
        return VaR_99_10d, CVaR_99_10d
    
    def get_consolidated_quarter_reporting(self, isin_list, start_date, end_date):
        portfolio_reporting_data = super().get_portfolio_reporting_data(isin_list, start_date, end_date)
        daily_returns_reporting_period = self.__daily_return(portfolio_reporting_data)
        df_portfolio, optimal_weights = self.__compute_portfolio_daily_return(isin_list, start_date, end_date)
        period = str(start_date) + ' - ' + str(end_date)
        average_value = round(float(df_portfolio['VL'].mean()), 2)
        quarter_mu_portfolio = (df_portfolio['Return'].mean() + 1)**60 - 1
        quarter_mu_portfolio = round(float(quarter_mu_portfolio * 100), 2)
        sigma_portfolio = round(float(df_portfolio['Return'].std() * 100), 2)
        VaR_99_10d, CVaR_99_10d = self.__compute_Value_at_Risk(daily_returns_reporting_period, optimal_weights)
        VaR_99_10d = round(float(VaR_99_10d * 100), 2)
        CVaR_99_10d = round(float(CVaR_99_10d * 100), 2)
        data_reporting = {'Period': period,'Average Value (€)': average_value,'Average Return (%)': quarter_mu_portfolio,
                'Volatility (%)': sigma_portfolio,'VaR 99% 10days (%)': VaR_99_10d, 'CVaR 99% 10jours (%)': CVaR_99_10d}   
        df_quarter_reporting = pd.DataFrame(data_reporting, index=[0])
        df_quarter_reporting.to_excel("Quarter_report.xlsx")
        return df_quarter_reporting
        
    def reporting(self, isin_list, start_date, end_date):
        df_portfolio = self.__compute_portfolio_daily_return(isin_list, start_date, end_date)
        optimal_weights = self.__get_optimal_weights(isin_list, start_date)
        cumulative_returns = self.__get_cumulative_return(isin_list, start_date, end_date)
        return df_portfolio, optimal_weights, cumulative_returns