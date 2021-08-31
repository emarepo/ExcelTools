# -*- coding: utf-8 -*-

"""

@author: Emanuele Luzzi

Date: 04/2021

"""

import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
import datetime as dt
# import sys


class Organizer():
    
    def __init__(self, file_path, logger):
        
        self.file_path = Path(file_path)
        self.data = self.load_excel_file("EPFR Output")
        self.logger = logger
        
    def load_excel_file(self, sheet_name):
        df = pd.read_excel(self.file_path, sheet_name=sheet_name, index_col=0)
        df = df.fillna(0)
        
        return df
    
    def __repr__(self):
        info = f"Organizer({self.file_path})"
        self.logger.info(info)
        return info
    
    def __str__(self):
        info = f"Organizer in use for the following file: {self.file_path}"
        self.logger.info(info)
        return info
    
    """
    
            ################################
    
             Metrics and Descriptive Stats
                                            
            ################################

    
    """
    
    def calc_sums_metrics(self, dataframe, list_of_addends_labels):
        sums_column = pd.Series(index=sorted(set(self.data.index.tolist()))).fillna(0)
        for addend in list_of_addends_labels:
            to_add_series = dataframe[addend].fillna(0)
            sums_column += to_add_series
        
        return sums_column
    
    """
    
            ################################
    
                Data Matrices Operations
                                            
            ################################

    
    """
    
    def calc_average_between_dataframes_sheets(self, list_of_dataframe_sheets):
        n = len(list_of_dataframe_sheets)
        list_of_dataframes = [self.load_excel_file(sheet).fillna(0) for sheet in list_of_dataframe_sheets]
        start_df = pd.DataFrame(index=list_of_dataframes[0].index, columns = list_of_dataframes[0].columns).fillna(0)
        for df in list_of_dataframes:
            sum_df = start_df.add(df)
            start_df = sum_df
        
        avg_df = sum_df/n
 
        return avg_df
    
    def calc_ratio_between_dataframes_sheets(self, duple_of_dataframe_sheets):
        duple_of_dataframes = [self.load_excel_file(sheet).fillna(0) for sheet in duple_of_dataframe_sheets]
        ratio_df = duple_of_dataframes[0].divide(duple_of_dataframes[1])
        
        return ratio_df
    
    def calc_rebased_dataframe_sheet(self, df_to_rebase, dd_mm_yy, marginal_return_label):
        df_to_rebase = self.load_excel_file(df_to_rebase)
        starting_date = dt.datetime.strptime(dd_mm_yy, "%d.%m.%y")
        marginal_return_dataframe = self.load_excel_file(marginal_return_label)
        rebased_dataframe = pd.DataFrame(index=df_to_rebase.loc[df_to_rebase.index >= starting_date].index, 
                                         columns=df_to_rebase.columns).fillna(0)
        rebased_dataframe.loc[starting_date] = [100] * len(df_to_rebase.columns)
        
        for date in rebased_dataframe.index:
            for column in rebased_dataframe.columns:
                if date != starting_date:
                    marginal_return = marginal_return_dataframe.loc[date, column]
                    index_list = rebased_dataframe.index.tolist()
                    current_position_of_date = index_list.index(date)
                    previous_date = index_list[current_position_of_date - 1]
                    
                    rebased_dataframe.loc[date, column] = (rebased_dataframe.loc[previous_date, column] 
                                                           * (1 + marginal_return) + marginal_return)
                
        return rebased_dataframe
    
    """
    
            ################################
    
                 Data Cleansing Methods
                                            
            ################################

    
    """
        
    def add_dataframe_sheet(self, dataframe, sheet_name):
        xlsx_file = load_workbook(self.file_path)
        writer = pd.ExcelWriter(self.file_path, engine = 'openpyxl')
        writer.book = xlsx_file
        
        try:
            xlsx_file.remove(xlsx_file[sheet_name])
        except Exception as e:
            print(f"Creating a new sheet, since {e}")
        finally:
            dataframe.to_excel(writer, sheet_name = sheet_name)
        
        self.logger.info(f"{sheet_name} added to the sheets")
        
        writer.save()
        writer.close()    
