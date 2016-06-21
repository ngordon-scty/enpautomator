from . import Workbook
from xlwings import Range, Sheet
import threading
import pythoncom
import Queue
import time
import traceback

class ThreadedWorkbook(Workbook):
    def __init__(self, **kwargs):
        self.busy = True
        self.alive = True
        pythoncom.CoInitialize()
        self.q = Queue.Queue()
        self.thread = threading.Thread(target=self._start_thread, kwargs=kwargs)
        self.thread.start()
    
    def _start_thread(self, **kwargs):
        pythoncom.CoInitialize()
        self.busy = True
        try:
            super(ThreadedWorkbook, self).__init__(newinstance=True, **kwargs)
        except Exception as e:
            self._quit(True)
            raise e
        while self.alive:
            try:
                task = self.q.get(True, 0.01)
                self.busy = True
                task.status = "busy"
                try:
                    task.retval = task.function(*task.args, **task.kwargs)
                except Exception as e:
                    traceback.print_exc()
                    task.error = e
                task.status = "finished"
            except Queue.Empty:
                self.busy = False
        self.xl_workbook = None
        self.xl_app = None
    
    def _execute_threaded(self, task):
        self.q.put(task)
    
    def run_macro(self,macroname):
        self._execute_threaded(WorkbookTask(self._run_macro,macroname))
        
    def unprotect(self,sheet,password):
        self._execute_threaded(WorkbookTask(self._unprotect,sheet,password))
        
    def quit(self,force=True):
        self._execute_threaded(WorkbookTask(self._quit,force))
        
    def calculate(self):
        self._execute_threaded(WorkbookTask(self._calculate))
    
    def get_value(self, *args, **kwargs):
        task = WorkbookTask(self._get_value,*args,**kwargs)
        self._execute_threaded(task)
        while task.status != "finished" and self.alive:
            pass
        return task.retval
    
    def set_value(self, sheetname, address, value):
        if value != None:
            self._execute_threaded(WorkbookTask(self._set_value,sheetname, address, value))
    
    def save_as(self,filename):
        self._execute_threaded(WorkbookTask(self._save_as,filename))
    
    def _run_macro(self,macroname):
        self.xl_app.Run(macroname)
        
    def _unprotect(self,sheetname,password):
        sheet = Sheet(sheetname)
        sheet.xl_sheet.Unprotect(password)
        
    def _calculate(self):
        self.xl_app.Calculate()
    
    def _get_value(self, *args, **kwargs):
        ra = self._get_range(*args, **kwargs)
        return ra.value
    
    def _set_value(self, sheetname, address, value):
        ra = self._get_range(sheetname, address)
        ra.value = value
    
    def _get_range(self, *args, **kwargs):
        return Range(wkb=self, *args, **kwargs)
    
    def _save_as(self, filename):
        self.xl_workbook.SaveAs(filename)
    
    def _quit(self,force):
        killa = ExcelKiller(self.xl_app)
        killa.get_handles()
        if force:
            self.xl_app.DisplayAlerts = False
        self.xl_app.Quit()
        killa.kill()
        self.alive = False

class ExcelKiller(object):
    def __init__(self, excel):
        self.excel = excel
        self.hwnd = None
    
    def get_handles(self):
        if self.excel:
            self.hwnd = self.excel.Hwnd
    
    def kill(self):
        import win32process
        import win32gui
        import win32api
        import win32con
        if self.excel and self.hwnd:
            try:
                t, p = win32process.GetWindowThreadProcessId(self.hwnd)
                handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, 0, p)
                if handle:
                    win32api.TerminateProcess(handle, 0)
                    win32api.CloseHandle(handle)
            except Exception as e:
                print("exception in ExcelKiller.kill:")
                print(e)
      
class WorkbookTask(object):
    def __init__(self, function, *args, **kwargs):
        self.status = "waiting"
        self.function = function
        self.args = args
        self.kwargs = kwargs
        self.retval = 0