# -*- coding: utf-8 -*-
"""
Created on Thu May 23 11:52:01 2024

@author: mmorin
"""

config = {
    'filepath':{
        'CAC40_Index.csv':{'name':'C:/Users/morin/Desktop/Projet_VBA_Python/PTF_MNGT_PYTHON/Database/CAC40_Index.csv'},
        'CAC40_Stock_Market.txt':{'name':'C:/Users/morin/Desktop/Projet_VBA_Python/PTF_MNGT_PYTHON/Database/CAC40_Stock_Market_2010_2019.txt'},
        'CAC40_closing.xlsb':{'name':'C:/Users/morin/Desktop/Projet_VBA_Python/PTF_MNGT_PYTHON/Database/CAC40_closing_94_to_22.xlsb'},
        'data_reporting.xlsx':{'name':'C:/Users/morin/Desktop/Projet_VBA_Python/PTF_MNGT_PYTHON/Database/data_reporting.xlsx'}
    },
    'parameters':{
        'risk_free_rate':{'name':'0'},
        'num_simulations':{'name':'10000'}
    },
    'VaR':{
        'confidence_level':{'name':'0.01'},
        'num_days':{'name':'10'}
    }
    }
