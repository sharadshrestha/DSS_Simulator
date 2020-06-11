import socket
def server_program():


    server_socket =  socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # get instance
    
    # get the hostname
    #host = '192.168.1.94'#socket.gethostname()
    host = socket.gethostbyname(socket.gethostname())
    port = 5000  # initiate port no above 1024
    
    
    # look closely. The bind() function takes tuple as argument
    server_socket.bind((host, port))  # bind host address and port together
    print("Server Address and port @ {}:{}".format(host, port))
    # configure how many client the server can listen simultaneously
    server_socket.listen(2)
    conn, address = server_socket.accept()  # accept new connection
    print("Connection from: " + str(address))
    acknowledge =-1
    while True:
        # receive data stream. it won't accept data packet greater than 1024 bytes
        data = conn.recv(1024).decode()
        if not data:
            # if data is not received break
            break
        print("from connected user: " + str(data))
       
        acknowledge =acknowledge+1 #
        data = str(acknowledge)
        conn.send(data.encode())  # send data to the client

    conn.close()  # close the connection


if __name__ == '__main__':
    server_program()