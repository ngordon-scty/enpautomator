import os.path
import shutil
import glob
import logging
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
    
    def open(self,readonly=False):
        if self.workbook == None and self.exists():
            self.workbook = ENPWorkbook(fullname=self.get_full_path(),app_visible=False,readonly=readonly)
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
        return os.path.join(self.get_path(),self.get_filename())
    
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

    def __repr__(self):
        return "ENP('{path}')".format(path=self.get_full_path())

class ENPWorkbook(ThreadedWorkbook):
    def __init__(self, *args, **kwargs):
        self.name = ""
        self.ids = None
        super(ENPWorkbook,self).__init__(*args, **kwargs)
        
    def enp_version(self):
        version = self.get_value('Rev','Version') 
        return version if version > 0 else None 
        
    def get_mp_cell_location(self):
        return 'E82:K82'

    def get_mps(self):
        mps = self.get_value('ENP',self.get_mp_cell_location())
        return filter(None,mps)
    
    def get_mp_index_for_mp(self, mp_name):
        mps = self.get_value('ENP',self.get_mp_cell_location())
        return mps.index(mp_name) + 1 if mp_name in mps else None
        
    def get_enp_id_for_mp(self,enp_id,mp_index):
        row = self.get_enp_row_for_id(enp_id)
        if row > 0:
            return self.get_value('ENP',(row,4+mp_index))
        return None
    
    def set_enp_id_for_mp(self,enp_id,mp_index,value):
        row = self.get_enp_row_for_id(enp_id)
        if row > 0:
            return self.set_value('ENP',(row,4+mp_index),value)
        return None        
    
    def get_ids(self):
        if self.ids == None:
            self.ids = self.get_value('ENP','A3:A211')
        return self.ids        
    
    def get_enp_row_for_id(self,enp_id):
        row = 3
        for id in self.get_ids():
            if id == enp_id:
                return row
            row = row + 1
        return 0
    
    def clear_enp_tab(self):
        self.set_value('ENP','E3:K211',"")
        
    def is_crackled(self):
        mps = self.get_mps()
        crackled = True
        for mp in mps:
            crackled = crackled and self.sheet_exists("{mp}-S".format(mp=mp))
        return crackled
    
    def rename_duplicate_mps(self):
        mps = self.get_mps()
        newmps = []
        for idx, mp in enumerate(mps):
            countmps = mps[:(idx+1)].count(mp)
            if countmps > 1:
                mp = "{mp}-DUPE{n}".format(mp=mp,n=(countmps-1))
            newmps.append(mp)
        for idx, mp in enumerate(newmps):
            if newmps[idx] != mps[idx]:
                self.set_enp_id_for_mp(5400,(idx+1),newmps[idx])
        return mps != newmps
        
    def crackle(self,force=False):
        if force or not self.is_crackled():
            self.rename_duplicate_mps()
            self.run_macro('Button_Crackle')
    
    def analyze_mp(self,mpname):
        if self.activate_sheet(mpname + "-S"):
            self.run_macro('Button_RunAnalysis_CurrentSelectedMP_WithUpgrades')
            return True
        return False

class ENPCopier(object):
    def __init__(self,src_wb,dest_wb):
        self.src_wb = src_wb
        self.dest_wb = dest_wb
    
    def copy(self):
        if self.src_wb is not None and self.dest_wb is not None:
            srcmps = self.src_wb.get_mps()
            index = 0
            for mp in srcmps:
                index = index + 1
                for id in self.src_wb.get_ids():
                    if id is not None:
                        val = self.src_wb.get_enp_id_for_mp(id,index)
                        if val is not None:
                            logging.debug("copying value '{val}' for mp '{mp} ({mpidx})' to new workbook".format(val=val,mp=mp,mpidx=index))
                            self.dest_wb.set_enp_id_for_mp(id,index,val)
        
        