import os.path
import shutil
import glob
from xlwings_plus import ThreadedWorkbook

class ENP(object):
    def __init__(self, projectnumber):
        self.projectnumber = projectnumber
        self._filename = None
        self._path = None
        self.localdestination = 'c:\\common\\ENPCache\\' + projectnumber + '\\'
        self.useUNC = False
        self.workbook = None
    
    def find_latest_revision(self):
        ENPs = sorted(glob.glob(self.get_path() + 'JB-' + self.projectnumber + '-00*.xl*'), key=os.path.getmtime, reverse=True)
        if len(ENPs) > 0:
            self._filename = os.path.basename(ENPs[0])
    
    def open(self):
        if self.workbook == None and self.exists():
            self.workbook = ENPWorkbook(fullname=self.get_full_path(),app_visible=False)
        return self.workbook
    
    def saveas(self,newpath,newfilename):
        if self.workbook == None:
            return self.copy_to(newpath,newfilename)
        else:
            self.workbook.save_as(os.path.join(newpath,newfilename))
            self._path = newpath
            self._filename = newfilename
            return True
    
    def close(self):
        if not self.workbook == None:
            self.workbook.quit(force=True)
            self.workbook = None
            
    def exists(self):
        return os.path.isfile(self.get_full_path())
    
    def copy_to_local(self):
        return self.copy_to(self.localdestination,self.get_filename())
    
    def copy_to(self, newpath, newfilename):
        if self.exists():
            if not os.path.exists(newpath):
                os.makedirs(newpath)
            try:
                shutil.copy2(self.get_full_path(),os.path.join(newpath,newfilename))
                self._filename = newfilename
                self._path = newpath
                return True
            except:
                return False
        return False

    def get_full_path(self):
        return self.get_path() + self.get_filename()
    
    def get_filename(self):
        if self._filename != None:
            return self._filename
        return self.get_default_filename()
    
    def get_default_filename(self):
        return 'JB-' + self.projectnumber + '-00.xlsm'
    
    def get_path(self):
        if self._path != None:
            return self._path
        return self.get_default_path()
    
    def get_default_path(self):
        pathsuffix = self.projectnumber[:3] + '\\' + self.projectnumber + '\\Drawings\\Structural\\'
        if self.useUNC:
            return '\\\\triton\\jobs\\' + pathsuffix
        else:
            return 'Z:\\' + pathsuffix

class ENPWorkbook(ThreadedWorkbook):
    def __init__(self, *args, **kwargs):
        self.name = ""
        super(ENPWorkbook,self).__init__(*args, **kwargs)

    def get_mps(self):
        mps = self.get_value('ENP','E82:K82')
        return filter(None,mps)