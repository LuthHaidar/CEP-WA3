import numpy as np
import pandas as pd

from classes import *
from functions import *

global name, date, timein, timeout, weaponid, pelletcount

list = req()

new_entry = New_Entry(list)

reset()