import numpy as np
import pandas as pd

global name, date, timein, timeout, weaponid, pelletcount

class New_Entry:
    def __init__(self, name, date, timein, timeout, weaponid, pelletcount):
        self.name = name
        self.date = date
        self.timein = timein
        self.timeout = timeout
        self.weaponid = weaponid
        self.pelletcount = pelletcount
