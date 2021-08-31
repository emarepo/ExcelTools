# -*- coding: utf-8 -*-
"""
Created on Fri May  7 11:38:10 2021

@author: luzziem
"""

from Tools.setup_logging import logger
from Tools.organizer import Organizer
import configparser
from pathlib import Path

if __name__ == "__main__":
    
    # Read config File
    config = configparser.RawConfigParser()   
    config_path = r'config.ini'
    config.read(config_path)
    
    file_path = Path(config.get("Paths", "Main File Path"))
    test_organizer = Organizer(file_path, logger)
    how_we_define_europe = ["Austria",
                            "Belgium",
                            "Finland",
                            "France",
                            "Germany",
                            "Greece",
                            "Ireland"
                            "Italy", 
                            "Netherlands",
                            "Portugal"
                            "Spain"]
    
    how_we_define_usa = ["USA"]

    rebase_date = str(config.get("Parameters", "Rebase Date"))

    print(test_organizer)
    test_organizer.create_dataframe_for_every_country_and_category()
    test_organizer.create_dataframe_for_every_variable_country_and_category(europe_members=how_we_define_europe,
                                                                            usa_members=how_we_define_usa)

    avg_dataframe = test_organizer.calc_average_between_dataframes_sheets(["Total Net Assets Start", "Total Net Assets"])
    test_organizer.add_dataframe_sheet(avg_dataframe, "Average NAV")

    ratio_dataframe = test_organizer.calc_ratio_between_dataframes_sheets(["Flow US$ mill", "Total Net Assets"])
    test_organizer.add_dataframe_sheet(ratio_dataframe, "Flow %")

    rebased_index_dataframe = test_organizer.calc_rebased_dataframe_sheet("Flow %", rebase_date, "Flow %")
    test_organizer.add_dataframe_sheet(rebased_index_dataframe, "Rebased Flow %")

