import openpyxl as xl
import datetime as dt
import os

class Logger:
    def __init__(self,path = "~",prefix = "log",name = None,overwrite=False,meta_data = None,headers:list= None):
        self.path = path
        self.prefix = prefix
        self.name = name
        self.counter = 0
        if self.name is None:
            self.name = dt.datetime.now().strftime("%Y-%m-%d")    
        if not os.path.exists(self.path):
            os.makedirs(self.path)
        self.file_name = self.prefix + "_" + self.name + ".xlsx"
        self.file_path = os.path.join(self.path,self.file_name)
        self.file_path = os.path.expanduser(self.file_path)
        if  os.path.exists(self.file_path):
            self.wb = xl.load_workbook(self.file_path)
        else:
            self.wb = xl.Workbook()
        self.ws = self.wb.create_sheet(dt.datetime.now().strftime("%H-%M-%S"))
        if meta_data is not None:
            self.ws.append(meta_data.keys())
            self.ws.append(meta_data.values())
        self.ws.append([])
        if headers is not None:
            self.ws.append(["SI No" + "Time"] + headers)
        else:
            self.ws.append(["SI No","Time","Data"])
    
    def log(self,data:list):
        self.counter += 1
        self.ws.append([self.counter,dt.datetime.now().strftime("%H:%M:%S")] + data)    
    
    def save(self):
        self.wb.save(self.file_path)
    
        
