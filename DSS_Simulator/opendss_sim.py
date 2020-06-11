'''
Created on May 21, 2020
@author: spate181
'''



import os
import csv
import sys
import win32com.client #pywin32
from win32com.client import makepy
import numpy as np
import math
from pdb import set_trace
import timeit
import copy
import pandas as pd
import socket
class opendsstools(object):
    
    '''
    classdocs
    '''


    def __init__(self, dssFileName):
        sys.argv = ["makepy", r"OpenDSSEngine.DSS"]
        makepy.main()  # ensures early binding and improves speed

        # Create a new instance of the DSS
        print("Initiating opendss engine")
        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
        if self.dssObj.Start(0) == False:
            print ("DSS Failed to Start")
        else:
            #Assign a variable to each of the interfaces for easier access
            self.dssText = self.dssObj.Text
            self.dssCircuit = self.dssObj.ActiveCircuit
            self.dssSolution = self.dssCircuit.Solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssMeters = self.dssCircuit.Meters
            self.dssPDElement = self.dssCircuit.PDElements
            self.dssLoads = self.dssCircuit.Loads
            self.dssLines = self.dssCircuit.Lines
            self.dssTransformers = self.dssCircuit.Transformers
            self.dssFileName=dssFileName
            self.dssPVs=self.dssCircuit.PVSystems
            self.loadnodes=[]
            self.l53flow=[]

        # Always a good idea to clear the DSS when loading a new circuit
        self.dssObj.ClearAll()

        # Loads the given circuit master file into OpenDSS
        self.dssText.Command = "compile " + dssFileName

        # Lists
        self.busesNames = self.dssCircuit.AllBusNames
        self.consumersNames = self.dssLoads.AllNames
        self.transformersNames = self.dssTransformers.AllNames
        self.busesNames_ltc = []
        self.consumersNames_ltc = []
        self.PVlines=[]
        self.lines=self.dssLines.AllNames
        self.capacitors=self.dssCircuit.Capacitors.AllNames
        self.capcontrols=self.dssCircuit.CapControls.AllNames
        self.regcontrols=self.dssCircuit.RegControls.AllNames
        self.loadoff=[]
        self.pvs=self.dssPVs.AllNames
        self.loads=self.dssLoads.AllNames
        self.bess_soc={}
        self.tieline={}
        
        self.tap_position={}
        self.der_output={}
        self.derIDs={}
        self.vsource=[]
        self.simulinkresults={}
        self.events=[{"ID":0, "Type":"Set Fault", "StartTime":1.1, "Bus":"650"},{"ID":1, "Type":"Remove Fault", "StartTime":1.15, "Bus":"650"}]
        #self.events=[{"ID":0, "Type":"None", "StartTime":1.1, "Bus":"650"},{"ID":1, "Type":"None", "StartTime":1.15, "Bus":"650"}]
        #self.events=[]
        self.derIDs=["DER1", "DER2", "DER3", "DER4", "DER5", "DER6", "DER7", "DER8", "DER9", "DER10", 
                     "DER11", "DER12", "DER13", "DER14", "DER15", "DER16", "DER17", "DER18", "DER19", "DER20", 
                     "DER21", "DER22", "DER23", "DER24", "DER25"]
        self.der_parameterlist=["Time", "Vfeederhead","Pfeederhead", "Qfeederhead", "P2grid", "Vpcc_a_dss", "Vpcc_b_dss", "Vpcc_c_dss", "angle_a_dss", "angle_b_dss", "angle_c_dss", "pflow_dss", "qflow_dss"]
        self.pvlines=["line1", "line2", "line3", "line4", "line5", "line6", "line7", "line8", "line9", "line10",
                      "line11", "line12", "line13", "line14", "line15", "line16", "line17", "line18", "lin19", "line20",
                      "line21", "line22", "line23", "line24", "line25"]
        self.pvpcc=["650", "651", "652", "653", "654", "655", "656", "657", "658","659" ,"660",
                    "661","662", "663", "664", "665", "666", "667", "668", "669", "670",
                    "671", "672", "673", "674"]
        self.dynamics_mode="Balanced"
        
        
    def setuppowerflow(self, time):
        print("Running powerflow")
        self.dssObj.AllowForms = "false"
        self.dssText.Command = "set controlmode=TIME"
        self.dssText.Command = "set maxcontroliter=300"
        self.dssText.Command = "set maxiterations=300"
        if time==0:
            self.dssText.Command = "set mode=Time stepsize=10s LoadShapeClass=Daily"
        else:
            [h,s]=self.run_time(time)
            self.dssText.Command = "set mode=Time stepsize=10s time=("+str(h)+','+str(s)+")"
        self.irr=[]
        self.pv_meas=[]
        #irradc or irr_ias10 or load_ias_sq, load_ias_tr
        if time==0:
            self.simulationmode="Timeseries"
            with open(r'C:\DSS_Simulator\ieee650v2_Renamed_withpv\Curves\irr_ias10.csv', 'r',newline='') as csvfile:
                    reader = csv.reader(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                    for row in reader:
                        self.irr.append(float(row[0]))
                        
            with open(r'C:\\DSS_Simulator\\vsource.csv', 'r',newline='') as csvfile:
            #with open(r'C:\\Users\\spate181\\eclipse-workspace\\Opendss2Matlab\\vsource_1.1_1.15.csv', 'r',newline='') as csvfile:
                    reader = csv.reader(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                    for row in reader:
                        self.vsource.append([float(row[0]), float(row[1])])
        
        print("Done")    
        
        
    def loadscaling(self):
        print("scaling load")
        '''
        Redirect the curves selected by the functions in the Settings Class
        '''
        #self.dssText.Command = r'New "LoadShape.LoadShape1" npts=1400 minterval=1 mult=(File=Curves\LoadShape1c.csv)'
        #self.dssText.Command = r'New "LoadShape.LoadShape1" npts=1400 minterval=1 mult=(File=Curves\test_load2_actualCopy.csv)'
        #load_ias10 or load_ias10c load_ias10s
        #self.dssText.Command = r'New "LoadShape.LoadShape1" npts=3000 sinterval=0.001 mult=(File=C:\Users\spate181\eclipse-workspace\Opendss2Matlab\ieee650v2_Renamed_withpv\Curves\load_ias10.csv)'
        self.dssText.Command = r'New "LoadShape.LoadShape1" npts=8640 sinterval=10 mult=(File=C:\DSS_Simulator\ieee650v2_Renamed_withpv\Curves\load_ias10.csv)'
        self.dssText.Command = "BatchEdit Load.. daily=LoadShape1"
        self.dssLoads.daily="LoadShape1"
        self.dssCircuit.LoadShapes.Name="LoadShape1"
        a=self.dssCircuit.LoadShapes
        print("pause")    
    
    def powerflow(self, time):
        self.change_irradiance("self_read", time)
        #self.set_source_voltage(time) #From PSSE in future
        self.dssText.Command = "Set Maxiterations=20" 
        self.dssText.Command = "solve" 
        print("Done")
           
        if self.dssSolution.Converged=="False":
            print("Powerflow did not converge")
            set_trace()
        else:
            print("Powerflow converged")
            _activesource=self.dssCircuit.CktElements("Vsource.source")
            P=-1*_activesource.SeqPowers[2]
            Q=-1*_activesource.SeqPowers[3]
            print("Feederhead P: "+str(P))
            print("Feederhead Q: "+str(Q))
            from feederheader import fhClass
            fd = fhClass(str(P),str(Q))
            return fd 
            
            
                         
        #self.log_measurements(time)
    def change_irradiance(self, irrad, t):
        prop="irradiance"
        self.dssCircuit.SetActiveClass('PVSystem')
        for i in self.pvs:
            self.dssCircuit.SetActiveElement('PVSystem.'+i)
            if irrad=="self_read":
                self.dssCktElement.Properties(prop).Val=self.irr[t]
            else:
                self.dssCktElement.Properties(prop).Val=irrad
    

    def run_time(self):
        s=str(self.dssSolution.Seconds)
        h=str(self.dssSolution.Hour)
        time_start=(float(h)*3600)+float(s)
        return time_start            
    
    def busdata(self, bus):
        _activebus=self.dssCircuit.SetActiveBus(bus)
        _distance=self.dssCircuit.ActiveBus.Distance
        _X=self.dssCircuit.ActiveBus.x
        _Y=self.dssCircuit.ActiveBus.y
        bus_data=[bus, _distance, _X, _Y]
        return bus_data
    #def get_coordinates(self, bus):
    def initialize_log(self,list_of_sensors):
        data=pd.read_csv(list_of_sensors)
        data.head()
        data=data.values.tolist()
        self.log={}
        _c=""
        for _data in data:
            if _c!=_data[0]:
                self.log[_data[0]]={}
            self.log[_data[0]][_data[1]]=[]
            _c=_data[0]
        self.last_step_measurement=copy.deepcopy(self.log)          
        print("Logs initialized")
    
    def log_cap(self, name):
        #|V|,del,Q, status (on/off)
        _name="capacitor."+name
        _activeelement=self.dssCircuit.CktElements("Capacitor."+name)
        if _activeelement.Name.upper()!=_name.upper():
            print(_name+" is not a valid name, check input file!!!")
            set_trace()
        self.dssCircuit.SetActiveClass('Capacitor')
        self.dssCircuit.SetActiveElement(name)
        status=self.dssCktElement.Properties("States").Val
        
        if _activeelement.NumPhases==1:
            Q=[_activeelement.Powers[1]]
            V=[_activeelement.VoltagesMagAng[0]]
        if _activeelement.NumPhases==2:
            Q=[_activeelement.Powers[1],_activeelement.Powers[3]]
            V=[_activeelement.VoltagesMagAng[0],_activeelement.VoltagesMagAng[2]]
        if _activeelement.NumPhases==3:
            Q=[_activeelement.Powers[1],_activeelement.Powers[3],_activeelement.Powers[5]]
            V=[_activeelement.VoltagesMagAng[0],_activeelement.VoltagesMagAng[2], _activeelement.VoltagesMagAng[4]]
        bus=_activeelement.BusNames[0].split(".")[0]
        busdata=self.busdata(bus)
        capdata={}
        capdata["busdata"]=busdata
        capdata["V"]=V
        capdata["Q"]=Q
        capdata["Status"]=status
        return capdata
        
    def log_transformer(self, name):
        #Ntap, Vpri, Vsec, violation warnings, power flow
        _name="Transformer."+name
        _activeelement=self.dssCircuit.CktElements("Transformer."+name)
        if _activeelement.Name.upper()!=_name.upper():
            print(_name+" is not a valid name, check input file!!!")
            set_trace()
            
        if _activeelement.HasVoltControl:
            self.dssCircuit.SetActiveClass('RegControl')
            self.dssCircuit.SetActiveElement(name)
            tap=self.dssCktElement.Properties("TapNum").Val
            _activeelement=self.dssCircuit.CktElements("Transformer."+name)
        else:
            _activeelement=self.dssCircuit.CktElements("Transformer."+name)
            tap=0
        
        if _activeelement.NumPhases==1:
            Pprim=[_activeelement.Powers[0]]
            Psec=[_activeelement.Powers[4]]
            Qprim=[_activeelement.Powers[1]]
            Qsec=[_activeelement.Powers[5]]
            
            Vprim=[_activeelement.VoltagesMagAng[0]]
            Vsec=[_activeelement.VoltagesMagAng[4]]
        if _activeelement.NumPhases==2:
            Pprim=[_activeelement.Powers[0],_activeelement.Powers[2]]
            Psec=[_activeelement.Powers[6],_activeelement.Powers[7]]
            Qprim=[_activeelement.Powers[1],_activeelement.Powers[3]]
            Qsec=[_activeelement.Powers[7],_activeelement.Powers[8]]
            
            Vprim=[_activeelement.VoltagesMagAng[0],_activeelement.VoltagesMagAng[2]]
            Vsec=[_activeelement.VoltagesMagAng[6],_activeelement.VoltagesMagAng[7]]
        if _activeelement.NumPhases==3:
            Pprim=[_activeelement.Powers[0],_activeelement.Powers[2],_activeelement.Powers[4]]
            Psec=[_activeelement.Powers[8],_activeelement.Powers[10],_activeelement.Powers[12]]
            Qprim=[_activeelement.Powers[1],_activeelement.Powers[3],_activeelement.Powers[5]]
            Qsec=[_activeelement.Powers[9],_activeelement.Powers[11],_activeelement.Powers[13]]
            
            Vprim=[_activeelement.VoltagesMagAng[0],_activeelement.VoltagesMagAng[2], _activeelement.VoltagesMagAng[4]]
            Vsec=[_activeelement.VoltagesMagAng[8],_activeelement.VoltagesMagAng[10], _activeelement.VoltagesMagAng[12]]
        bus=_activeelement.BusNames[0].split(".")[0]
        busdata=self.busdata(bus)
        transformerdata={}
        transformerdata["busdata"]=busdata
        transformerdata["Vprim"]=Vprim
        transformerdata["Vsec"]=Vsec
        transformerdata["Pprim"]=Pprim
        transformerdata["Psec"]=Psec
        transformerdata["Qprim"]=Qprim
        transformerdata["Qsec"]=Qsec
        transformerdata["Tap"]=tap
        return transformerdata

    def log_pv(self, name):
        #Pout, Qout, freq, Vmag
        
        _name="PVSystem."+name
        _activeelement=self.dssCircuit.CktElements("PVSystem."+name)
        if _activeelement.Name.upper()!=_name.upper():
            print(_name+" is not a valid name, check input file!!!")
            set_trace()
            
        
        if _activeelement.NumPhases==1:
            Pout=[_activeelement.Powers[0]]
            Qout=[_activeelement.Powers[1]]
            
            V=[_activeelement.VoltagesMagAng[0]]
        if _activeelement.NumPhases==2:
            Pout=[_activeelement.Powers[0],_activeelement.Powers[2]]
            Qout[_activeelement.Powers[1],_activeelement.Powers[3]]
        
            V=[_activeelement.VoltagesMagAng[0],_activeelement.VoltagesMagAng[2]]
        if _activeelement.NumPhases==3:
            Pout=[_activeelement.Powers[0],_activeelement.Powers[2],_activeelement.Powers[4]]
            Qout=[_activeelement.Powers[1],_activeelement.Powers[3],_activeelement.Powers[5]]
            
            V=[_activeelement.VoltagesMagAng[0],_activeelement.VoltagesMagAng[2], _activeelement.VoltagesMagAng[4]]
        bus=_activeelement.BusNames[0].split(".")[0]
        busdata=self.busdata(bus)
        pvdata={}
        pvdata["busdata"]=busdata
        pvdata["P"]=Pout
        pvdata["Q"]=Qout
        pvdata["V"]=V
        return pvdata
    
    def log_load(self, name):
        #P,Q,V
        
        _name="Load."+name
        _activeelement=self.dssCircuit.CktElements("Load."+name)
        if _activeelement.Name.upper()!=_name.upper():
            print(_name+" is not a valid name, check input file!!!")
            set_trace()
            
        
        if _activeelement.NumPhases==1:
            Pout=[_activeelement.Powers[0]]
            Qout=[_activeelement.Powers[1]]
            
            V=[_activeelement.VoltagesMagAng[0]]
        if _activeelement.NumPhases==2:
            Pout=[_activeelement.Powers[0],_activeelement.Powers[2]]
            Qout[_activeelement.Powers[1],_activeelement.Powers[3]]
        
            V=[_activeelement.VoltagesMagAng[0],_activeelement.VoltagesMagAng[2]]
        if _activeelement.NumPhases==3:
            Pout=[_activeelement.Powers[0],_activeelement.Powers[2],_activeelement.Powers[4]]
            Qout=[_activeelement.Powers[1],_activeelement.Powers[3],_activeelement.Powers[5]]
            
            V=[_activeelement.VoltagesMagAng[0],_activeelement.VoltagesMagAng[2], _activeelement.VoltagesMagAng[4]]
        bus=_activeelement.BusNames[0].split(".")[0]
        busdata=self.busdata(bus)
        loaddata={}
        loaddata["busdata"]=busdata
        loaddata["P"]=Pout
        loaddata["Q"]=Qout
        loaddata["V"]=V
        return loaddata


    def log_measurements(self):
        sim_time=self.run_time()
        for i in self.log:
            for j in self.log[i]:
                self.log[i][j]={}
                if i=="Capacitor":
                    val=self.log_cap(j)
                    self.log[i][j][str(sim_time)]=val
                    self.last_step_measurement[i][j]=val
                if i=="Transformer":
                    val=self.log_transformer(j)
                    self.log[i][j][str(sim_time)]=val
                    self.last_step_measurement[i][j]=val
                if i=="PVSystem":
                    val=self.log_pv(j)
                    self.log[i][j][str(sim_time)]=val
                    self.last_step_measurement[i][j]=val
                if i=="Load":
                    val=self.log_load(j)
                    self.log[i][j][str(sim_time)]=val
                    self.last_step_measurement[i][j]=val
        print("logging measurements")
        return [sim_time, self.last_step_measurement]