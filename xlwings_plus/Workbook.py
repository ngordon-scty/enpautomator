from xlwings import Workbook as xlwings_Workbook
import win32com

class Workbook(xlwings_Workbook):
    def __init__(self, fullname=None, newinstance=False, **kwargs):
        if newinstance:
            xl_app, xl_workbook = self.open_workbook_in_new_instance(fullname)
            super(Workbook,self).__init__(xl_workbook = xl_workbook, **kwargs)
        else:
            super(Workbook,self).__init__(fullname=fullname, **kwargs)

    def open_workbook_in_new_instance(self,fullname):
        self.xl_app = self._get_new_excel()
        self.xl_workbook = self.xl_app.Workbooks.Open(fullname)
        return self.xl_app, self.xl_workbook
        
    def _get_new_excel(self):
        return win32com.client.DispatchEx('Excel.Application')