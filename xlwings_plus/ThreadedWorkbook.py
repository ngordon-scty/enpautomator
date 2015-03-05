from . import Workbook
from xlwings import Range
import threading
import pythoncom
import Queue
import time

class ThreadedWorkbook(Workbook):
    def __init__(self, **kwargs):
        pythoncom.CoInitialize()
        self.alive = True
        self.busy = False
        self.q = Queue.Queue()
        self.thread = threading.Thread(target=self._start_thread, kwargs=kwargs)
        self.thread.start()
    
    def _start_thread(self, **kwargs):
        pythoncom.CoInitialize()
        self.busy = True
        super(ThreadedWorkbook, self).__init__(newinstance=True, **kwargs)
        self.busy = False
        while self.alive:
            try:
                task = self.q.get()
                self.busy = True
                task.status = "busy"
                try:
                    task.retval = task.function(*task.args, **task.kwargs)
                except Exception as e:
                    task.error = e
                task.status = "finished"
            except Queue.Empty:
                self.busy = False
                self.idle()
    
    def _execute_threaded(self, task):
        self.q.put(task)
    
    def run_macro(self,macroname):
        self._execute_threaded(WorkbookTask(self._run_macro,macroname))
        
    def quit(self,force=True):
        self._execute_threaded(WorkbookTask(self._quit,force))
    
    def get_value(self, *args, **kwargs):
        task = WorkbookTask(self._get_value,*args,**kwargs)
        self._execute_threaded(task)
        while task.status != "finished":
            pass
        return task.retval
    
    def set_value(self, sheetname, address, value):
        if value != None:
            self._execute_threaded(WorkbookTask(self._set_value,sheetname, address, value))
    
    def _run_macro(self,macroname):
        self.xl_app.Run(macroname)
    
    def _get_value(self, *args, **kwargs):
        ra = self._get_range(*args, **kwargs)
        return ra.value
    
    def _set_value(self, sheetname, address, value):
        ra = self._get_range(sheetname, address)
        ra.value = value
    
    def _get_range(self, *args, **kwargs):
        return Range(wkb=self, *args, **kwargs)
    
    def _quit(self,force):
        if force:
            self.xl_app.DisplayAlerts = False
        self.xl_app.Quit()
        self.alive = False
    
    def idle(self):
        pass

class WorkbookTask(object):
    def __init__(self, function, *args, **kwargs):
        self.status = "waiting"
        self.function = function
        self.args = args
        self.kwargs = kwargs
        self.retval = 0