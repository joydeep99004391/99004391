from openpyxl import load_workbook,Workbook
def readworkbook():
      wb = load_workbook(filename = 'advancedpython.xlsx')
      sheet_names = wb.sheetnames
     # print(sheet_names)
     # ws=wb[sheet_names[0]]
      data={}
      
      for name in sheet_names:
        databuff=[]
        ws=wb[name]
     
        cellrange=ws['A1':'D5']

        r=[]
        
        for row in cellrange:
         r1=[]
     
         for col in row:
         
          r1.append(str(col.value))
         
         r.append(r1)   
         
        databuff.append(r)   
       # print("xxx",databuff)
        
        data[name]=databuff

      return data 
data=readworkbook()     
#print(data)           

def getdata(data,psno):
        total_data={}
        total_data_buffer=[]
        
        for i in data.keys():
                 data_buff=data[i]

                 
                 
                 for k in range(1,len(data_buff[0])):
                         
                           psn= data_buff[0][k][0]
                           #print("yyy",data_buff[0][k])
                           if psn==psno:
                             total_data_buffer.append( data_buff[0][k][1:])
                 total_data[psno]=total_data_buffer 
                               
        return total_data
tot=getdata(data,str(99004391.0))
print(tot)