'''
Created on May 21, 2020
@author: spate181
'''

from opendss_sim import opendsstools
from feederheader import fhClass
import timeit
import socket
dssfilelocation='C:\\DSS_Simulator\\ieee650v2_Renamed_withpv\\' #change
ig=opendsstools(dssfilelocation+"Master.dss")
list_of_sensors='C:\\DSS_Simulator\\ieee650v2_Renamed_withpv\\sensor_location.csv'
ig.initialize_log(list_of_sensors)
starttime=0
endtime= 8#640 #6*60*24



''' Socket '''

host = '192.168.1.79'#socket.gethostname()  # as both code is running on same pc
port = 5000  # socket server port number
client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # instantiate
client_socket.connect((host, port))  # connect to the server
print("Server Address and port @ {}:{}".format(host, port))         


#print('Received from server: ' + dpingata)  # show in terminal

    
ig.setuppowerflow(0)
ig.loadscaling() #loading feeder load profiles
 
 
  
for t in range(starttime, endtime):
    print(t)
    fd = fhClass(0,0);
    fd=ig.powerflow(t)
    #print (fd.p + fd.q);
    [sim_time, measurements]=ig.log_measurements()
    
    message = "P = "+fd.p+" Q= " + fd.q;
    client_socket.send(message.encode())  # send message
    data = client_socket.recv(1024).decode()  # receive response
    print('Ack No server: ' + data)
print("Done")
client_socket.close()

#Measurement format:
#simtime: seconds
#For active/reactive power and voltage measurements, Type Array: [PhA, PhB, PhC]
#For Busdata, Type: List, Format=[Name, Distance from subsystem, X coordinate, Y coordinate]