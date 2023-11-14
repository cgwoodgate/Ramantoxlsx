import os
import numpy as np
import pandas as pd

import xlsxwriter



# Set directory and find Files
path1 = os.getcwd()

for file in os.listdir(path1):
    d = os.path.join(path1, file)
    if os.path.isdir(d):
        print(d)
    
        path = d
        os.chdir(path)
        
        files = os.listdir(path)
        
        noFiles = len(files)
        # For each file - FIles are out of order w.r.t. time and proc.
        i=0
        X = np.zeros(noFiles)
        Y = np.zeros(noFiles)
        Z = np.zeros(noFiles)
        for file in files:
            #print(file)
            
            # Split path into parts and Find X, Y (and Z if it exists).
            X1 = files[i].split("X_")
            X2 =  X1[1].split("_")
            X[i] =  float(X2[0])
            #   
            Y1 = files[i].split("Y_")
            Y2 =  Y1[1].split("_")
            Y[i] =  float(Y2[0])
        
            #
            Z1 = files[i].split("Z_")
            if len(Z1) > 1:
                Z2 =  Z1[1].split("_")
                Z[i] =  float(Z2[0])
            else:
                Z[i] = 0
                    
            
            #import data into table
            lines=np.loadtxt(file,delimiter='\t')
           #Initialize matrix
            if i == 0:
                data = np.zeros((len(lines),noFiles+1))
                data[:,0] = lines[:,0]
                temp=file.split("_proc")
                fName = temp[0]
            
            #endif
                
            data[:,i+1] = lines[:,1]
            
            i=i+1;
            
        #end for    
        
        
        
        #Sort list so X=0 is first.
        indices = np.argsort(X)
        Xs = np.zeros((noFiles,1))
        Ys = np.zeros((noFiles,1))
        Zs = np.zeros((noFiles,1))
        datas = np.zeros((len(lines),noFiles+1))
        datas[:,0] = data[:,0]
        
        for i in range(0,noFiles):
            Xs[i] = X[indices[i]]
            Ys[i] = Y[indices[i]]
            Zs[i] = Z[indices[i]]
            datas[:,i+1] = data[:,indices[i]+1]
        #export table.
        os.chdir(path1)#
        
        df = pd.DataFrame(datas)
        excelName = fName+'.xlsx'
        writer = pd.ExcelWriter(excelName,engine='xlsxwriter')
        
        df.to_excel(writer,sheet_name='Sheet1',startrow=4,index=False)
        
        
        
        workbook = writer.book
        worksheet =writer.sheets['Sheet1']
        worksheet.write('A1', fName)
        worksheet.write('A2', 'X:')
        worksheet.write('A3' , 'Y:')
        worksheet.write('A4' ,'Z:')
        worksheet.write('A5', 'Raman wavenumber')
        
        for i in range(1,noFiles+1):
            worksheet.write(4,i,'Intensity')
            worksheet.write(1,i,Xs[i-1])
            worksheet.write(2,i,Ys[i-1])
            worksheet.write(3,i,Zs[i-1])
        writer.close()