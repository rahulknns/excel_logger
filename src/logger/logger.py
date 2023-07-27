import openpyxl as xl
import datetime as dt
import os
""" 
    @brief: This class is used to log data in excel file.
    @description: Logs the data in excel file. The data is saved in the form of rows. By default the file is saved with date in its name along with prefix. If such a file already exxists creates a new sheet in the same file. If the file does not exist then creates a new file with the given name. The data is saved in the form of rows. The first row is the header row. The first column is the serial number of the data. The second column is the time at which the data is logged. The rest of the columns are the data.
    @param: path: path where the file is to be saved.
    @param: prefix: prefix of the file name.
    @param: name: name of the file.
    @param: overwrite: if True then the file will be overwritten.
    @param: meta_data: meta data to be saved in the file.
    @param: headers: headers of the data to be saved in the file.
    @return: None
"""
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
    
    """
        @brief: This function is used to log data in excel file.
        @param: data: data to be logged.
        @return: None
    """
    def log(self,data:list):
        self.counter += 1
        self.ws.append([self.counter,dt.datetime.now().strftime("%H:%M:%S")] + data)    
    
    """
        @brief: This function is used to save the file.
        @param: None
        @return: None
    """
    def save(self):
        self.wb.save(self.file_path)
    
    """
        @brief: This function is used to delete the object.
        @param: None
        @return: None
    """
    def __del__(self):
        self.save()