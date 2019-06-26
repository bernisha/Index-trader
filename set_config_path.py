# -*- coding: utf-8 -*-
"""
Created on Wed Jun 26 13:24:15 2019

@author: BLala
"""

    
import csv
from pathlib import Path

base_path = Path(__file__).parent
file_path = (base_path / "../data/check.csv").resolve()
print(file_path)
#with open(file_path) as f:
    #test = [line for line in csv.reader(f)]