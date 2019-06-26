# -*- coding: utf-8 -*-
"""
Created on Thu Jun 20 11:09:26 2019

@author: BLala
"""

from configobj import ConfigObj
config = ConfigObj()
config.filename = 'c:\\data\\config.ini'
#
config['keyword1'] = 1
config['keyword2'] = "2"
#
config['section1'] = {}
config['section1']['keyword3'] = "s"
config['section1']['keyword4'] = 1
#
section2 = {
    'keyword5': 1,
    'keyword6': 1,
    'sub-section': {
        'keyword7': 1
        }
}
config['section2'] = section2
#
config['section3'] = {}
config['section3']['keyword 8'] = [1, 2, 3]
config['section3']['keyword 9'] = [1, 2, 3]
#
config.write()



import configparser
Config = configparser.ConfigParser()
Config
#<ConfigParser.ConfigParser instance at 0x00BA9B20>
Config.read("c:\\data\\config.ini")
Config.sections()

Config.options('section1')[1]


Config.get('section1',Config.options('section1')[4])
Config.get('section0',Config.options('section0')[0])