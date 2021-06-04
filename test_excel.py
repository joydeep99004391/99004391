from openpyxl import load_workbook, Workbook as Workbook_Open, workbook
from main import WorkBook 


def test_read_data():
          workbook=WorkBook()
          read_data=workbook.readworkbook()

          assert len(read_data) > 0

def test_get_data():
          workbook=WorkBook()
          read_data=workbook.readworkbook()
          PS_NO=str(float(99004390))

          getdata=workbook.getdata(read_data,PS_NO)

          assert len(getdata.keys()) >0

def test_get_selected_data():
          workbook=WorkBook()
          read_data=workbook.readworkbook()
          PS_NO=str(float(99004390))
          data_to_be_selected='games'
          getdata=workbook.getdata_select(read_data,PS_NO,data_to_be_selected)

          assert len(getdata.keys()) >0
def test_incorrect_psno():
          workbook=WorkBook()
          read_data=workbook.readworkbook()
          PS_NO=str(float(99004389))
          data_to_be_selected='games'
          getdata=workbook.getdata_select(read_data,PS_NO,data_to_be_selected)

          assert len(getdata.keys()) ==1
def test_incorrect_data_selected():
          workbook=WorkBook()
          read_data=workbook.readworkbook()
          PS_NO=str(float(99004389))
          data_to_be_selected='test'
          getdata=workbook.getdata_select(read_data,PS_NO,data_to_be_selected)

          assert len(getdata.keys()) ==0          