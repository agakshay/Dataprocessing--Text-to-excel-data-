# -*- coding: utf-8 -*-
"""
Spyder Editor

Editor: Akshay

Reference: https://www.gasturb.de/download.html (User manual guide for smooth c)
"""



import glob
files=(glob.glob("Z:\\compressordata\\*.txt"))

for text in range(len(files)):    
    text=files[text]
    filename = open(text, "r").readlines()
    result = list(filename)
    # print (result)
    #Line 1
    line =result[0]
    file_input = line.split()
    # Map_Type_Indicator
    Map_type= int(file_input[0])
    print("Map_Type_Indicator:", Map_type)
    # Compressor_Name
    y= file_input[1:(len(file_input))]
    Compressor_Name = '_'.join(y)
    print("Compressor_Name:", Compressor_Name)
    
    # create excel file with headers
    import xlsxwriter
    workbook = xlsxwriter.Workbook('Z:\\compressordata\\preprocessed_excel_files\\'+Compressor_Name+'.xlsx') 
    worksheet = workbook.add_worksheet() 
    
    worksheet.write(0,0,'Map_Type_Indicator') 
    worksheet.write(0,1, 'Compressor_Name') 
    worksheet.write(0,2, 'RefSpeed') 
    worksheet.write(0,3, 'RefBeta')
    worksheet.write(0,4, 'RefMach')
    worksheet.write(0,5, 'RefPsi') 
    worksheet.write(0,6, 'RefPhi') 
    worksheet.write(0,7, 'No_of_speed_lines')
    worksheet.write(0,8, 'Keyword') 
    worksheet.write(0,9, 'Speed') 
    worksheet.write(0,10, 'No_of_points_1') 
    worksheet.write(0,11, 'No_of_points_2')
    worksheet.write(0,12, 'Mass_flow') 
    worksheet.write(0,13, 'Pressure_ratio') 
    worksheet.write(0,14, 'Efficiency')
    worksheet.write(0,15, 'Surge_Mass_flow')
    worksheet.write(0,16, 'Surge_Pressure_ratio')
    worksheet.write(0,17, 'Surge_Efficiency')
    
  # Counting total number of lines
    file = open(text, "r")
    number_of_lines = 0
    for line in filename:
      line = line.strip("\n")
      number_of_lines += 1
    file.close()
    print("lines:", number_of_lines)
    
  # Writing down map type and cmpressor name      
    worksheet.write(1, 0, Map_type)    
    worksheet.write_string(1, 1, Compressor_Name)
    
    with open(text) as f:
        if 'RefSpeed' in f.read():
            print("Output1")
            #Line 2
            line =result[1]
            file_input = line.split()
            x=str(file_input[0])
            s1 = x[x.find('=') + 1: ]
            RefSpeed= str(s1)#RefSpeed
            print("RefSpeed:", RefSpeed)
            worksheet.write(1, 2, RefSpeed) 
            x=str(file_input[1])
            s1 = x[x.find('=') + 1: ]
            RefBeta= str(s1)#RefBeta
            print("RefBeta:", RefBeta)
            worksheet.write(1, 3, RefBeta)
            #Line 3
            line =result[2]
            file_input = line.split()
            x=str(file_input[0])
            s1 = x[x.find('=') + 1: ]
            RefMach= str(s1)#RefMach
            print("RefMach:", RefMach)
            worksheet.write(1, 4, RefMach)
            x=str(file_input[1])    
            s1 = x[x.find('=') + 1: ]
            RefPsi= str(s1)#RefPsi
            print("RefPsi:", RefPsi)
            worksheet.write(1, 5, RefPsi)
            x=str(file_input[2])
            s1 = x[x.find('=') + 1: ]
            RefPhi= str(s1)#RefBeta
            print("RefPhi:", RefPhi)
            worksheet.write(1, 6, RefPsi)
            
            #line4
            line =result[3]
            file_input = line.split()
            Map_Type_Indicator = len(file_input)
            # Number_of_speed_lines
            Number_of_speed_lines= int(file_input[0])
            print("Number_of_speed_lines:", Number_of_speed_lines)
            worksheet.write(1, 7, Number_of_speed_lines)
            # Compressor_Name
            y= file_input[1:(len(file_input))]
            Keyword = ''.join(y)
            print("Keyword:", Keyword)
            worksheet.write(1, 8, Keyword)
            
            line =result[4]
            file_input = line.split()
            points=int(file_input[1])#Since num points_1=num points_2, taking first one
            Num_speed_lines= Number_of_speed_lines
            low_limit=4
            z=low_limit
            y=0
            pt=0
            Press_ratio_arr=[]
            Mass_flow_arr=[]
            Efficiency_arr=[]
            Pts_upda=0
     #Writing down the speed  lines and points    
            for num1 in range(0, Num_speed_lines):
                x=num1
                line =result[low_limit]
                file_input = line.split()
                # speed
                speed= float(file_input[0])
                print("speed:", speed)
                worksheet.write(Pts_upda+1, 9, speed)
                # points
                Num_points_1= int(file_input[1])
                print("Num_points_1:", Num_points_1)
                worksheet.write(Pts_upda+1, 10, Num_points_1)
                Num_points_2= int(file_input[2])
                print("Num_points_2:", Num_points_2)
                worksheet.write(Pts_upda+1, 11, Num_points_2)
                
    # to find the lower and upper limit
                z=z+pt+1
                y=z+points 
                Pts_upda=Pts_upda + points
                for num in range(z, y):            
                    line =result[num]
                    file_input = line.split()
                    Press_ratio=float(file_input[0])
                    Press_ratio_arr.append(Press_ratio)
                    print("Press_ratio:", Press_ratio)
                    Mass_flow=float(file_input[1]) 
                    Mass_flow_arr.append(Mass_flow)
                    print("Mass_flow:", Mass_flow)
                    Efficiency=float(file_input[2]) 
                    Efficiency_arr.append(Efficiency)
                    print("Efficiency:", Efficiency)
                
                low_limit=low_limit+points+1
                pt=points
                
            worksheet.write_column(1, 12, Press_ratio_arr)
            worksheet.write_column(1, 13, Mass_flow_arr) 
            worksheet.write_column(1, 14, Efficiency_arr)
            with open(text) as k:
                if 'Surge Line' in k.read():
    # writing surge line
                    p=5+Number_of_speed_lines+(points*Number_of_speed_lines)
                    q=number_of_lines
                    Pts_updat=0
                    for num2 in range(p, q):
                        line =result[num2]
                        file_input = line.split()
                        Press_ratio_surge=float(file_input[1])
                        print("Press_ratio:", Press_ratio_surge)
                        Mass_flow_surge=float(file_input[2])
                        print("Mass_flow:", Mass_flow_surge)
                        Efficiency_surge=float(file_input[3])
                        print("Efficiency:", Efficiency_surge)
                        worksheet.write(Pts_updat+1, 15, Press_ratio_surge)
                        worksheet.write(Pts_updat+1, 16, Mass_flow_surge) 
                        worksheet.write(Pts_updat+1, 17, Efficiency_surge)
    #Updating the points
                        Pts_updat=Pts_updat + points
                else:
                    p=5+Number_of_speed_lines+(points*Number_of_speed_lines)
                    q=number_of_lines
                    Pts_updat=0

                    for num2 in range(p, q):
                        line =result[num2]                       
                        worksheet.write(Pts_updat+1, 15, 'null')
                        worksheet.write(Pts_updat+1, 16, 'null') 
                        worksheet.write(Pts_updat+1, 17, 'null')
    #Updating the points
                        Pts_updat=Pts_updat + points
                    
                    
        else:
            print("Output2")
            #line1
            line =result[1]
            file_input = line.split()
            Map_Type_Indicator = len(file_input)
            #Since no reference values, all the unknown values are assigned to null
            worksheet.write(1, 2, 'null')
            worksheet.write(1, 3, 'null')
            worksheet.write(1, 4, 'null')
            worksheet.write(1, 5, 'null')
            worksheet.write(1, 6, 'null')
            # Number_of_speed_lines
            Number_of_speed_lines= int(file_input[0])
            print("Number_of_speed_lines:", Number_of_speed_lines)
            worksheet.write(1, 7, Number_of_speed_lines)
            # Compressor_Name
            y= file_input[1:(len(file_input))]
            Keyword = ''.join(y)
            print("Keyword:", Keyword)
            worksheet.write(1, 8, Keyword)
            
            line =result[2]
            file_input = line.split()
            points=int(file_input[1])
            Num_speed_lines= int(Number_of_speed_lines)
            low_limit=2
            z=low_limit
            y=0
            pt=0
            speed_upda=0
            Press_ratio_arr=[]
            Mass_flow_arr=[]
            Efficiency_arr=[]
            
            for num1 in range(0, Num_speed_lines):
                x=num1
                line =result[low_limit]
                file_input = line.split()
                # speed
                speed= float(file_input[0])
                print("speed:", speed)
                worksheet.write(speed_upda+1, 9, speed)
                # Compressor_Name
                Num_points_1= int(file_input[1])
                print("Num_points_1:", Num_points_1)
                worksheet.write(speed_upda+1, 10, Num_points_1)
                Num_points_2= int(file_input[2])
                print("Num_points_2:", Num_points_2)
                worksheet.write(speed_upda+1, 11, Num_points_2)
                
        # to find the lower and upper limit
                z=z+pt+1
                y=z+points 
                speed_upda=speed_upda + points
        #looping through the points
                for num in range(z, y):            
                    line =result[num]
                    file_input = line.split()
                    Press_ratio=float(file_input[0])   
                    print("Press_ratio:", Press_ratio)
                    Press_ratio_arr.append(Press_ratio)
                    Mass_flow=float(file_input[1])  
                    Mass_flow_arr.append(Mass_flow)
                    print("Mass_flow:", Mass_flow)
                    Efficiency=float(file_input[2])
                    Efficiency_arr.append(Efficiency)
                    print("Efficiency:", Efficiency)
                
                low_limit=low_limit+points+1
                pt=points
            worksheet.write_column(1, 12, Press_ratio_arr)
            worksheet.write_column(1, 13, Mass_flow_arr) 
            worksheet.write_column(1, 14, Efficiency_arr)
            with open(text) as k:
                if 'Surge Line' in k.read():
    # writing surge line
                    print("\nAdding Surge points")
                    p=3+Number_of_speed_lines+(points*Number_of_speed_lines)
                    q=number_of_lines
                    Pts_updat=0
                    for num2 in range(p, q):
                        line =result[num2]
                        file_input = line.split()
                        Press_ratio_surge=float(file_input[1])
                        print("Press_ratio:", Press_ratio_surge)
                        Mass_flow_surge=float(file_input[2])
                        print("Mass_flow:", Mass_flow_surge)
                        Efficiency_surge=float(file_input[3])
                        print("Efficiency:", Efficiency_surge)
                        worksheet.write(Pts_updat+1, 15, Press_ratio_surge)
                        worksheet.write(Pts_updat+1, 16, Mass_flow_surge) 
                        worksheet.write(Pts_updat+1, 17, Efficiency_surge)
    #Updating the points 
                        Pts_updat=Pts_updat + points
                else:
                    print("\nNo Surge points")
                    Pts_updat=0
                    for num2 in range(0, Number_of_speed_lines):
                        line =result[num2]                       
                        worksheet.write(Pts_updat+1, 15, 'null')
                        worksheet.write(Pts_updat+1, 16, 'null') 
                        worksheet.write(Pts_updat+1, 17, 'null')
    #Updating the points
                        Pts_updat=Pts_updat + points                        
            
    workbook.close()   
   
    
    
