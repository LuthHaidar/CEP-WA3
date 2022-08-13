import numpy as np
import pandas as pd

global name, date, timein, timeout, weaponid, pelletcount

def req(): #Requests user for info
    #variables: name, date, timein, timeout, weaponid, pelletcount
    name = input("Please enter your name: ")
    date = input("Please enter the date: ")
    timein = input("Please enter the time you entered: ")
    timeout = input("Please enter the time you left: ")
    weaponid = input("Please enter the weapon ID: ")
    pelletcount = input("Please enter the pellet count: ")
    return name, date, timein, timeout, weaponid, pelletcount

def reset(): #Resets the variables
    name = ""
    date = ""
    timein = ""
    timeout = ""
    weaponid = ""
    pelletcount = ""
    return name, date, timein, timeout, weaponid, pelletcount