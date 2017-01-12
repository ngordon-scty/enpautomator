from xlwings import Workbook as xlwings_Workbook
import win32com

class Workbook(xlwings_Workbook):
    def __init__(self, fullname=None, newinstance=False, readonly=False, **kwargs):
        self.xl_app = None
        self.xl_workbook = None
        self.window_handle = None
        if newinstance:
            self.xl_app = self._get_new_excel()
            self.window_handle = self.xl_app.Hwnd
            self.xl_workbook = self.open_workbook_in_new_instance(fullname,readonly)
            super(Workbook,self).__init__(xl_workbook = self.xl_workbook, **kwargs)
        else:
            super(Workbook,self).__init__(fullname=fullname, **kwargs)

    def open_workbook_in_new_instance(self,fullname,readonly=False):
        return self.xl_app.Workbooks.Open(fullname,None,readonly)
        
    def _get_new_excel(self):
        return win32com.client.DispatchEx('Excel.Application')