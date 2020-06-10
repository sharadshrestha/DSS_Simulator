'''
Created on May 26, 2020

@author: spate181
'''

import sys
import win32com.client #pywin32
from win32com.client import makepy
sys.argv = ["makepy", r"OpenDSSEngine.DSS"]
makepy.main()  # ensures early binding and improves speed


# Create a new instance of the DSS
print("Initiating opendss engine")
dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
if dssObj.Start(0) == False:
    print ("DSS Failed to Start")
else:
    print ("Python-dss interface successful")