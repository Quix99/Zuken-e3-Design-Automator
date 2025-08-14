import glob
import pandas as pd
import json
import win32com.client as win32
from json2json import get_value_by_path, find_key_paths
import time
import logging
import os




def find_key_paths(data, target_key):
    paths = []

    def recurse(obj, current_path):
        if isinstance(obj, dict):
            for key, value in obj.items():
                new_path = current_path + [key]
                if key == target_key:
                    paths.append(new_path)
                recurse(value, new_path)
        elif isinstance(obj, list):
            for index, item in enumerate(obj):
                recurse(item, current_path + [index])

    recurse(data, [])
    return paths

def extract_paths(string,level,all_paths):
    paths=[]
    for path in all_paths:
        if string in path[level]:
            paths.append(path)
    return paths

def device_iterator(func):
    def wrapper(self):
        clipboard = self.job.CreateClipboardObject()
        clipboardIds = self.get_clipboardIds()
        try:
            clitem = clipboardIds[-1]
        except:
            clitem=0
        clipboard.SetId(clitem)
        itemIds = []
        result = clipboard.GetAnyIds(0, itemIds)
        device = self.job.CreateDeviceObject()
        self.devc = 1

        for item in result[1]:
            if item is not None:
                #device.SetId(item)
                func(self, item,clitem)

    return wrapper

def symbol_iterator(func):
    def wrapper(self, smu_id, module_place):
        clipboard = self.job.CreateClipboardObject()
        clipboardIds = self.get_clipboardIds()
        clitem = clipboardIds[-1]
        clipboard.SetId(clitem)
        itemIds = []
        result = clipboard.GetAnyIds(0, itemIds)
        symbol = self.job.CreateSymbolObject()

        for item in result[1]:
            if item is not None:
                symbol.SetId(item)
                func(self, symbol, smu_id, module_place)

    return wrapper



class Draw():

    def __init__(self,**kwargs):
        self.zuken = win32.gencache.EnsureDispatch('CT.Application')
        self.zuken.PutInfo(0,"Hello-Zuken-on-Steroids")
        self.projectname = []
        self.job=self.zuken.CreateJobObject()
        #self.job.Open(project)
        cwd = os.getcwd()
        self.modulepath = os.path.join(cwd,"zuken_devices")
        self.modulepath=str(self.modulepath).replace("\\","/")+"/"
        self.sheetpath = os.path.join(cwd,"zuken_sheets")
        self.sheetpath=str(self.sheetpath).replace("\\","/")+"/"

        self.sheet = self.job.CreateSheetObject()
        self.connection = self.job.CreateConnectionObject()
        self.core = self.job.CreatePinObject()
        self.node = self.job.CreatePinObject()
        self.netseg = self.job.CreateNetSegmentObject()
        self.net = self.job.CreateNetObject()
        self.device = self.job.CreateDeviceObject()


        sheetIds = []
        result = self.job.GetSheetIds(sheetIds)
        if result[1][-1]!=None:
            sheetId = self.sheet.SetId(result[1][-1]) 
            result = self.sheet.Delete() 


        with open('data1.json', 'r', encoding='utf-8') as f:
                    self.data = json.load(f)


        self.last = 0
        self.con_id = 0

        self.path=[]

        self.namespaces={}

        if "log" in kwargs:
            self.log=kwargs["log"]
        else:
            self.log = logging.getLogger(__name__)
            logging.basicConfig(filename="json2zuken.log",format='%(asctime)s %(message)s',level=logging.DEBUG)

    def get_clipboardIds(self):
        clipboardIds = []
        clipboardCount = self.job.GetClipboardIds(clipboardIds)
        return clipboardCount[1][1:]
    
    def commit(self):
        clipboard = self.job.CreateClipboardObject()
        clipboardIds = self.get_clipboardIds()
        for clitem in clipboardIds:
            clipboard.SetId(clitem)
            clipboard.CommitToProject(1)     

    def add_sheets(self):
        self.log.info("Starting add_sheets...")
        paths = find_key_paths(self.data,"Sheet Template")
        for path in paths:
            try:
                template = get_value_by_path(self.data,path)
                result1=self.job.ImportDrawing(self.sheetpath+template,1)
                sheetIds = []
                result = self.job.GetSheetIds(sheetIds)
                sheetId = self.sheet.SetId(result[1][-1]) 
                self.sheet.SetName(path[-2])
                self.log.info(f"Set sheet name to {path[-2]} with Zuken ID {result[1][-1]}")
                sheet_item = get_value_by_path(self.data,path[:-1])
                if "Sheet Attributes" in sheet_item:
                    pass
                else: 
                    sheet_item["Sheet Attributes"]={}
                sheet_item["Sheet Attributes"]["Zuken ID"] = result[1][-1]
            except Exception as e:
                self.log.warning(f"Error adding sheet for path {path}: {e}")

        self.log.info("add_sheets completed.")


    def add_parts(self):
         self.log.info("Starting add_parts...")
         paths = find_key_paths(self.data,"Part Item")
         side=["Header Field", "Header Text","Schematics Text"]
         for path in paths:
            self.path = path
            part = get_value_by_path(self.data,path[:-1])
            element = get_value_by_path(self.data,path[:-3])
            sheet = get_value_by_path(self.data,path[:-5])
            xPos = element["Element Attributes"]["xPos"]
            yPos = element["Element Attributes"]["yPos"]
            xOffset = 0
            yOffset = 0
            if "Part Attributes" in part:
                if "xOffset" in part["Part Attributes"] and  "yOffset" in part["Part Attributes"]:
                 xOffset = part["Part Attributes"]["xOffset"]
                 yOffset = part["Part Attributes"]["yOffset"]
            try:
                sheetId = sheet["Sheet Attributes"]["Zuken ID"]
                self.sheet.SetId(sheetId)
            except Exception as e:
                self.log.warning(f"Sheet Id was not found for part {part['Part Item']} at path {path}: {e}")
            try:    
                result = self.sheet.PlacePartEx(self.modulepath+part["Part Item"], 1, 4,xPos+xOffset,yPos+yOffset,0.0 )
            except Exception as e:
                self.log.error(f"Failed to place part {part['Part Item']} at path {path}: {e}")
                continue
            self.commit()
            self.log.info("add_parts completed.")


    def add_connections(self):
        self.log.info("Starting add_connections...")
        paths = find_key_paths(self.data,"Connection")
        for path in paths:
            self.path=path
            connection = get_value_by_path(self.data,path[:-2])
            sheet = get_value_by_path(self.data,path[:-4])
            try:
                xPos1 = connection["Element Attributes"]["CON 1 xPos"]
                yPos1 = connection["Element Attributes"]["CON 1 yPos"]
                xPos2 = connection["Element Attributes"]["CON 2 xPos"]
                yPos2 = connection["Element Attributes"]["CON 2 yPos"]
            except Exception as e:
                self.log.warning(f"Coordinates were not found for connection at path {path}: {e}")
                continue
            try:
                sheetId = sheet["Sheet Attributes"]["Zuken ID"]
                self.sheet.SetId(sheetId)
            except Exception as e:
                self.log.warning(f"Sheet Id was not found for connection at path {path}: {e}")
                continue
            xPos = [0,xPos1,xPos2]
            yPos = [0,yPos1,yPos2]
            pt = [0,0,0,0,0]
            netSegments=[]
            #result3 = self.connection.CreateConnection( 0, sheetId, 2, xPos, yPos, netSegments, pt )
            result3 = self.connection.CreateConnectionBetweenPoints( sheetId,xPos[1], yPos[1], xPos[2], yPos[2],0)
            #result3 = -1
            if result3<1:
                result3 = self.connection.CreateConnection( 0, sheetId, 2, xPos, yPos, netSegments, pt )
            id =[]
            try:
                if type(result3)==int:
                    connection["Element Attributes"]["Connection ID"]=result3
                else:
                    connection["Element Attributes"]["Connection ID"]=result3[1][-1]
                    result3=result3[1][-1]
            except Exception as e:
                self.log.warning(f"Could not assign Connection ID for path {path}: {e}")
            segId=result3
            ids=[]
            #this is problematic, might not hit the right connection Id
            conIds = self.job.GetAllConnectionIds(ids,2)
            conId = conIds[1][-1]
            try:
                self.add_wire(connection,conId,segId)
                self.log.info(f"Added wire for connection at path {path}, conId: {conId}, segId: {segId}")
            except Exception as e:
                self.log.warning(f"Failed to add wire for connection at path {path}: {e}")
            self.log.info("add_connections completed.")

    def add_texts(self):
        self.log.info("Starting add_texts...")
        paths = find_key_paths(self.data,"Text Item")
        for path in paths:
            text = get_value_by_path(self.data,path[:-1])
            element = get_value_by_path(self.data,path[:-3])
            sheet = get_value_by_path(self.data,path[:-5])
            xPos = element["Element Attributes"]["xPos"]
            yPos = element["Element Attributes"]["yPos"]
            xOffset = 0
            yOffset = 0
            if "Text Attributes" in text:
                if "xOffset" in text["Text Attributes"] and  "yOffset" in text["Text Attributes"]:
                 xOffset = text["Text Attributes"]["xOffset"]
                 yOffset = text["Text Attributes"]["yOffset"]
            try:
                sheetId = sheet["Sheet Attributes"]["Zuken ID"]
                self.sheet.SetId(sheetId)
            except Exception as e:
                self.log.warning(f"Sheet Id was not found for text '{text['Text Item']}' at path {path}: {e}")
                continue
            try:
                graphic=self.job.CreateGraphObject()
                sheetId = self.sheet.GetId()
                result = graphic.CreateText(sheetId,list(text["Text Item"].keys())[0],xPos+xOffset,yPos+yOffset)
                if "Text Colour" in text["Text Attributes"]:
                    textcolor = text["Text Attributes"]["Text Colour"]
                    graphic.SetColour(textcolor)
                if "Text Height" in text["Text Attributes"]:
                    textheight = text["Text Attributes"]["Text Height"]
                    graphic.SetTextHeight(textheight)
            except Exception as e:
                self.log.warning(f"Failed to create/set attributes for text '{text['Text Item']}' at path {path}: {e}")

        self.log.info("add_texts completed.")          

    def add_wire(self,connection,conId,segId):
            self.log.info("Starting add_wire...")
            id = []
            try:
                self.connection.SetId(conId)
                if type(segId)==int:
                    self.netseg.SetId(segId)

                if "Line Color" in connection["Element Attributes"]:
                    self.netseg.SetLineColour(connection["Element Attributes"]["Line Color"])
                if "Line Width" in connection["Element Attributes"]:
                    self.netseg.SetLineWidth(connection["Element Attributes"]["Line Width"]) 
                if "Wire Item" in connection["Element Attributes"]:
                    wire_item=connection["Element Attributes"]["Wire Item"]

                    con_id = self.con_id
                    self.con_id = self.con_id+1
                    #last = int(deviceName[-2])*10+int(deviceName[-1])
                    if con_id > 99:
                        wire_num ="-W"+str(con_id)
                    elif con_id > 9:
                        wire_num ="-W0"+str(con_id)
                    elif con_id < 10:
                        wire_num = "-W00"+str(con_id)

                    location=""
                    if "COOLING" in self.path[3]:
                        location="Cooling Circuit"
                    if "DASHBOARD" in self.path[3]:
                        location="Bridge"
                    devId = self.device.Create(wire_num,self.path[3],location,wire_item,0,0)
                    corIds = self.device.GetCoreIds(id)
                    self.core.SetId(corIds[1][1])
                    self.log.info("Core: "+self.core.GetName())
                    #pinIds=self.connection.GetPinIds(id)
                    self.netseg.SetId(segId)
                    self.net.SetId(self.netseg.GetNetId())
                    pinIds=self.net.GetPinIds([])

                    if len(pinIds[1])>=2:
                        result1=self.core.SetEndPinId(1,pinIds[1][1])
                        self.log.info("End Pins Set")
                    else:
                        self.log.info("No End Pins found")
                    if len(pinIds[1])==3:
                        result2=self.core.SetEndPinId(2,pinIds[1][2])
                    else:
                        self.log.info("No Second End Pin found")
                    self.set_assignment(self.device,self.sheet.GetName())
            except Exception as e:
                self.log.error(f"Error in add_wire: {e}")


    def save(self):
        name = self.data["Project Attributes"]["Project Name"]
        # Step 1: Get the current working directory
        cwd = os.getcwd()

        # Step 2: Construct the path for 'generated projects'
        folder_name = 'generated_projects'
        folder_path = os.path.join(cwd, folder_name)

        # Step 3: Create the folder if it doesn't exist
        os.makedirs(folder_path, exist_ok=True)

        # Step 4: Set the filename and complete file path
        file_path = os.path.join(folder_path, name)

        #export pdf
        sheetIds = []
        result = self.job.GetSheetIds(sheetIds)
        self.job.ExportPDF( file_path+".pdf", result[1][1:],0)
        # Save data to the file
        self.job.SaveAs(file_path+".e3s")   


    def draw(self):

        
        self.add_sheets()

        self.add_parts()

        self.add_connections()

        self.add_texts()

        self.commit()


        with open("data2.json", "w") as file:
            json.dump(self.data, file, indent=4)

        self.save()
        self.log.info("Drawing completed")


    
    
if __name__ == "__main__":

    h = Draw()
    h.draw()