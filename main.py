from openpyxl import load_workbook, Workbook

'''
@author: joydeep ghosh 99004391
@description : read workbook and return data
'''


def readworkbook():
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


data = readworkbook()


# print(data)

def getdata(data, psno):
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

    return total_data


# psno=str(float(input("enter psno ")))
# tot=getdata(data,psno)
# print(tot)

def getdata_select(data, psno, dataselect):
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

    return total_data


def write_data(data):
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


if __name__ == '__main__':
    psno = str(float(input("enter psno ")))
    data_to_be_selected = input("enter data to be selected ")
    databyselect = getdata_select(data, psno, data_to_be_selected)
    # print(databyselect)
    write_data(databyselect)
