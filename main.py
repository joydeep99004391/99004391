from openpyxl import load_workbook, Workbook
class WorkbookIo():
    '''
       @author: joydeep ghosh 99004391
       @description :WorkbookIo class defining classes
       @date : 3-06-2021
    '''
    def readworkbook(self):
        '''
        @author: joydeep ghosh 99004391
        @description :method implementing read workbook
        @date : 3-06-2021
        '''
        try:
            workbookload = load_workbook(filename='advancedpython.xlsx')
            sheet_names = workbookload.sheetnames
            data = {}
            for name in sheet_names:
                databuff = []
                worksheet = workbookload[name]
                cellrange = worksheet['A1':'D5']
                row_list = []
                for row in cellrange:
                    row1 = []
                    for col in row:
                        row1.append(str(col.value))
                    row_list.append(row1)
                databuff.append(row_list)
                data[name] = databuff
            return data
        except IOError as ex:
            print("file reading error", ex)       
    '''
    @author: joydeep ghosh 99004391
    @description :method implementing writing data to workbook
    @date : 3-06-2021
    '''
    def write_data(self, data):
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
        except IOError as ex:
            print("file writing error", ex)




'''
@author: joydeep ghosh 99004391
@description : Workbook class inheriting from WorkbookIo
@date : 3-06-2021

'''
class WorkBook(WorkbookIo):
    '''
    @author: joydeep ghosh 99004391
    @description :method implementing getting all data
    @date : 3-06-2021
    '''
    def getdata(self, data, psno):
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
        except ValueError as ex:
            print("error in reading data", ex)
        return total_data    
    '''
    @author: joydeep ghosh 99004391
    @description :method implementing getting selected data 
    @date : 3-06-2021
    '''
    def getdata_select(self, data, psno, dataselect):
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
        except ValueError as ex:
            print("error in reading selected data", ex)
        return total_data

if __name__ == '__main__':
    PS_NO = str(float(input("enter psno ")))
    data_to_be_selected = input("enter data to be selected ")
    workbook = WorkBook()
    Data_Read = workbook.readworkbook()
    data_by_select = workbook.getdata_select(Data_Read, PS_NO, data_to_be_selected)
    # databyselect = getdata_select(data, psno, data_to_be_selected)
    # print(databyselect)
    workbook.write_data(data_by_select)
