from openpyxl import load_workbook, Workbook

'''
@author: joydeep ghosh 99004391
@description : read workbook and return data
'''

class workbookio():
 def readworkbook(self):
    try: 
      wb = load_workbook(filename='advancedpython.xlsx')
      sheet_names = wb.sheetnames
      data = {}
      for name in sheet_names:
        databuff = []
        ws = wb[name]
        cellrange = ws['A1':'D5']
        r = []
        for row in cellrange:
            r1 = []
            for col in row:
                r1.append(str(col.value))
            r.append(r1)
        databuff.append(r)
        data[name] = databuff
      return data
    except Exception as e:
          print("file reading error",e)  
 def write_data(self,data):
    try: 
      wb1 = Workbook()
      destination_file_name = 'selected_data.xlsx'
      ws1 = wb1.active
      ws1.title = 'selected'
      ws1['A1'] = 'data'
      list2 = [['psno', 'one', 'two', 'three']]
      for i in range(1, 2):
        for j in range(1, 5):
            ws1.cell(column=j, row=i, value=list2[i - 1][j - 1])
      for psn in data.keys():
        for row in range(2, 3):
            ws1.cell(row=row, column=1, value=psn)
            for col in range(2, 5):
                ws1.cell(row=row, column=col, value=data[psn][0][col - 2])
            # print(data[psn][0][1])
      wb1.save(filename=destination_file_name)
    except Exception as e :
          print("file writing error",e)


#data = readworkbook()


# print(data)
class workbook(workbookio):
 def getdata(self,data, psno):
   try:  
    total_data = {}
    total_data_buffer = []

    for i in data.keys():
        data_buff = data[i]

        for k in range(1, len(data_buff[0])):

            psn = data_buff[0][k][0]
            # print("yyy",data_buff[0][k])
            if psn == psno:
                total_data_buffer.append(data_buff[0][k][1:])
        total_data[psno] = total_data_buffer
   except Exception as e:
        print("error in reading data",e)
   return total_data


# psno=str(float(input("enter psno ")))
# tot=getdata(data,psno)
# print(tot)

 def getdata_select(self,data, psno, dataselect):
    try: 
     total_data = {}
     total_data_buffer = []

     for i in data.keys():
        if i == dataselect:
            data_buff = data[i]

            for k in range(1, len(data_buff[0])):

                psn = data_buff[0][k][0]
                # print("yyy",data_buff[0][k])
                if psn == psno:
                    total_data_buffer.append(data_buff[0][k][1:])
            total_data[psno] = total_data_buffer
    except Exception as e:
           print("error in reading selected data",e)
    return total_data



if __name__ == '__main__':
    psno = str(float(input("enter psno ")))
    data_to_be_selected = input("enter data to be selected ")
    workbook=workbook()
    data=workbook.readworkbook()    
    data_by_select =workbook.getdata_select(data,psno,data_to_be_selected)
    #databyselect = getdata_select(data, psno, data_to_be_selected)
    # print(databyselect)
    workbook.write_data(data_by_select)
