import json
import pandas as pd
import glob
import copy
import re
import logging
import numpy as np



def find_first_occurrence_indices(df: pd.DataFrame, search_value):
    """
    Find the first occurrence of search_value in the DataFrame.
    Returns (row_index, col_name) if found, otherwise None.
    """
    # Convert to string for safe comparison even with NaN values
    positions = np.where(df.astype(str) == str(search_value))
    if positions[0].size > 0:
        row_index = df.index[positions[0][0]]
        col_name = df.columns[positions[1][0]]
        return int(row_index), int(col_name)
    else:
        return None


def get_value_by_path(data, path):
    for key in path:
        data = data[key]
    return data

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

def check_for_string_in_value(data, val, search_string):
    for key, value in data.items():
        if val in value:
            element_item = value['Element Item']
            if search_string in element_item:
                #print(f"Found '{search_string}' in key: {key}")
                return True  # Stop after finding the first occurrence
    return False  # Return False if not found

def check_for_string_in_key(data, val, search_string):
    for key, value in data.items():
        if val in value:
            element_item = list(value['Element Item'].keys())[0]
            if search_string in element_item:
                #print(f"Found '{search_string}' in key: {key}")
                return True  # Stop after finding the first occurrence
    return False  # Return False if not found

def any_key_values(data,target_key,target_value):
    paths = find_key_paths(data,target_key)
    for path in paths:
        value = get_value_by_path(data,path)
        if target_value in value:
            return 1
    return 0

def delete_keys(data,keys):
    if isinstance(keys,list):
        for key in keys:
            del data[key]
    else:
        del data[keys]



class TransJsonBase():
    def __init__(self,**kwargs):
        with open('data.json', 'r', encoding='utf-8') as f:
                    self.data = json.load(f)

        with open('connector_positions.json', 'r', encoding='utf-8') as f:
                self.pos_data = json.load(f)

        self.xlsx_files = glob.glob("*.xlsm")
    # Excel-Datei lesen
        self.excel_file = self.xlsx_files[0]
        self.excel_file = "conf.xlsm"
        self.df1 = pd.read_excel(self.excel_file, "Item Positions", header=None,usecols='A:EU')
        self.df2 = pd.read_excel(self.excel_file, "Run", header=None)


        self.modulepath = self.data["Project Attributes"]["Module Path"]
        self.sheetpath = self.data["Project Attributes"]["Sheet Path"]


        self.fgroup1_flag = int(self.df2.loc[11,10])
        self.fgroup2_flag = int(self.df2.loc[12,10])

        paths = find_key_paths(self.data,"Group Item")
        functional_group_item = get_value_by_path(self.data, paths[0][:-2])     
        keys = list(functional_group_item.keys())

        for key in keys:
            if "FUNCTIONAL_GROUP_1" in key:
                if self.fgroup1_flag == 0:
                    del functional_group_item[key]
            elif "FUNCTIONAL_GROUP_2" in key:
                if self.fgroup2_flag == 0:
                    del functional_group_item[key]


        if "log" in kwargs:
            self.log=kwargs["log"]
        else:
            self.log = logging.getLogger(__name__)
            logging.basicConfig(filename="json2json.log",format='%(asctime)s %(message)s',level=logging.DEBUG)


    def cleanup(self,part_key):
        #data =self.data
        all_paths = find_key_paths(self.data,"Set")
        paths = extract_paths(part_key,3,all_paths)
        for path in paths:
             try:
                  item = get_value_by_path(self.data,path[:-3])
                  if item[path[-3]]["Element Attributes"]["Set"]!="+":
                    del item[path[-3]]
             except Exception as e:
                self.log.error(f"Error processing Set Item at {path}: {e}", exc_info=True)
               

    def add_drawing(self,part_key):
        self.log.info(f"Starting add_drawing...")
        data = self.data
        try:
            all_paths = find_key_paths(data,"Sheet Item") 
            paths = extract_paths(part_key,3,all_paths)
            self.log.info(f"Found {len(paths)} Sheet Item paths for part_key '{part_key}'.")
            for path in paths:
                try:
                    sheet_item = get_value_by_path(data,path)
                    all_paths = find_key_paths(sheet_item,"Element Item") 
                    for path2 in all_paths:
                        try:
                            element_item = get_value_by_path(sheet_item,path2)
                            for key in list(element_item.keys()):
                                if key !="Connection" and not "Text" in key:
                                    key1=key
                                    element_item[key] = {"Part Item":part_key+" "+key1+".e3p"}     
                                    self.log.debug(f"Set Part Item for key '{key}' in Element Item at {path2}.")
                        except Exception as e:
                            self.log.error(f"Error processing Element Item at {path2} in Sheet Item at {path}: {e}", exc_info=True)
                except Exception as e:
                    self.log.error(f"Error processing Sheet Item at {path}: {e}", exc_info=True)
        except Exception as e:
            self.log.error(f"Critical error in add_drawing for part_key '{part_key}': {e}", exc_info=True)

    def add_sheet_template(self,sheet_key):
        self.log.info(f"Starting add_sheet_template...")
        data=self.data
        try:
            all_paths = find_key_paths(data,"Sheet Item")
            paths = extract_paths(sheet_key,3,all_paths)


            for path in paths:
                try:
                    sheet_item = get_value_by_path(data,path[:-1])
                    sheet_item["Sheet Template"] = sheet_key+".e3p"
                except Exception as e:
                    self.log.error(f"Error setting Sheet Template for path {path[:-1]}: {e}", exc_info=True)
        except Exception as e:
            self.log.critical(f"Critical error in add_sheet_template for sheet_key '{sheet_key}': {e}", exc_info=True)
        return data
    
    def add_positions(self,position_key):
        self.log.info("Starting add_positions for "+position_key+" ...")
        n_modules = []
        xPoss = []
        yPoss = []
        i = 0

        coords=find_first_occurrence_indices(self.df1,position_key)
        #imax
        paths= find_key_paths(self.data,"Element Type Device") 
        posrow=self.df1.loc[coords[0]+1,coords[1]+1:]
        imax=int(posrow.isna().to_numpy().argmax())
        #imax=len(paths)*2
        item_line=coords[0]
        position_line=item_line+1
        try:
            while True:
                #n_modules.append(str(self.df1.loc[21,3+i]).replace(" ",""))
                n_modules.append(str(self.df1.loc[item_line,coords[1]+1+i]))
                xPoss.append(int(self.df1.loc[position_line,coords[1]+1+i]))
                yPoss.append(int(self.df1.loc[position_line,coords[1]+2+i]))
                i=i+2
                if i > imax:
                    break
            self.log.info(f"Loaded {len(n_modules)} module positions from dataframe.")
        except Exception as e:
            self.log.error(f"Error loading positions from dataframe: {e}")

        data = self.data
        try:
            all_paths = find_key_paths(data,"Sheet Item") 
            paths = extract_paths(position_key,3,all_paths)
            for path in paths:
                sheet_item = get_value_by_path(data,path)
                for key in sheet_item.keys():
                    for n_module,xPos,yPos in zip(n_modules,xPoss,yPoss):
                        if key == n_module:
                            if "Element Attributes" in sheet_item[key]:
                                pass
                            else: sheet_item[key]["Element Attributes"]={}
                            sheet_item[key]["Element Attributes"]["xPos"]=xPos
                            sheet_item[key]["Element Attributes"]["yPos"]=yPos
        except Exception as e:
            self.log.error(f"Error assigning module positions: {e}")
        #add connection positions
        try:
            for path in paths:
                sheet_item = get_value_by_path(data,path)
                for key in sheet_item.keys():
                    if "Connection" in sheet_item[key]["Element Item"]:
                        try:
                            xPos1 = sheet_item[sheet_item[key]["Element Attributes"]["CON 1 Sheet Item"]]["Element Attributes"]["xPos"]
                            yPos1 = sheet_item[sheet_item[key]["Element Attributes"]["CON 1 Sheet Item"]]["Element Attributes"]["yPos"]

                            port_offset = self.pos_data[position_key+" "+sheet_item[key]["Element Attributes"]["CON 1 Part"]][sheet_item[key]["Element Attributes"]["CON 1 Port"]]["Position"]
                            xPos1 = xPos1+port_offset[0]
                            yPos1 = yPos1+port_offset[1]

                            xPos2 = sheet_item[sheet_item[key]["Element Attributes"]["CON 2 Sheet Item"]]["Element Attributes"]["xPos"]
                            yPos2 = sheet_item[sheet_item[key]["Element Attributes"]["CON 2 Sheet Item"]]["Element Attributes"]["yPos"]   

                            port_offset = self.pos_data[position_key+" "+sheet_item[key]["Element Attributes"]["CON 2 Part"]][sheet_item[key]["Element Attributes"]["CON 2 Port"]]["Position"]
                            xPos2 = xPos2+port_offset[0]
                            yPos2 = yPos2+port_offset[1]

                            sheet_item[key]["Element Attributes"]["CON 1 xPos"]=xPos1
                            sheet_item[key]["Element Attributes"]["CON 1 yPos"]=yPos1

                            sheet_item[key]["Element Attributes"]["CON 2 xPos"]=xPos2
                            sheet_item[key]["Element Attributes"]["CON 2 yPos"]=yPos2
                        except Exception as e:
                                                self.log.warning(f"Error setting connection positions for {key} in sheet at {path}: {e}")

        except Exception as e:
            self.log.error(f"Error assigning connection positions: {e}")

        self.log.info("add_positions for "+position_key+" completed.")

class TransJson():
    def __init__(self,**kwargs):
        if "log" in kwargs:
            self.log=kwargs["log"]
        self.fgroup1 = Transfgroup1()
        self.fgroup2 = Transfgroup2()




    def transformJson(self):
        


        if self.fgroup1.fgroup1_flag == 1:
            self.fgroup1.data = self.fgroup1.data 
            self.fgroup1.log.info("Preparing fgroup1")
            self.fgroup1.add_drawing("FUNCTIONAL_GROUP_1")
            self.fgroup1.add_positions("FUNCTIONAL_GROUP_1")
            self.fgroup1.add_sheet_template("FUNCTIONAL_GROUP_1")
            self.fgroup1.add_wire()
            self.fgroup1.cleanup("FUNCTIONAL_GROUP_1")
                
            self.data = self.fgroup1.data

        if self.fgroup2.fgroup2_flag == 1:
            self.fgroup2.data = self.fgroup1.data
            self.fgroup2.log.info("Preparing fgroup2")
            self.fgroup2.complete_structure()
            self.fgroup2.cleanup()
            self.fgroup2.add_drawing("FUNCTIONAL_GROUP_2")
            self.fgroup2.add_positions("FUNCTIONAL_GROUP_2")
            self.fgroup2.add_sheet_template("FUNCTIONAL_GROUP_2")    


        with open("data1.json", "w") as file:
            json.dump(self.data, file, indent=4)
        print("finished")

















































class Transfgroup1(TransJsonBase):



    def add_wire(self):
        self.log.info("Starting add_wire for FUNCTIONAL_GROUP_1...")
        data = self.data
        try:
            all_paths = find_key_paths(data,"Sheet Item") 
            paths = extract_paths("FUNCTIONAL_GROUP_1",3,all_paths)
            for path in paths:
                sheet_item = get_value_by_path(data,path)
                try:
                    for key in sheet_item.keys():
                        if "Connection" in sheet_item[key]["Element Item"]:
                            sheet_item[key]["Element Attributes"]["Wire Item"]="036-00375K"
                            sheet_item[key]["Element Attributes"]["Line Color"]=59
                            sheet_item[key]["Element Attributes"]["Line Width"]=2
                        
                except Exception as e:
                                    self.log.warning(f"Error adding wire attributes to {key} in sheet at {path}: {e}")
                        # add core (if needed, add your logic here and log accordingly)
        except Exception as e:
            self.log.error(f"Error in add_wire for FUNCTIONAL_GROUP_1: {e}")

        self.log.info("add_wire for FUNCTIONAL_GROUP_1 completed.")




class Transfgroup2(TransJsonBase):

    def complete_structure(self):
        self.log.info(f"Starting complete_structure...")
        data = self.data
        #remove unwished cooling items
        all_paths = find_key_paths(data,"Sheet Item")
        paths = extract_paths("FUNCTIONAL_GROUP_2",3,all_paths)      
        try:
            for path in paths:
                group_item = get_value_by_path(data,path[:-3])
                if group_item["Group Attributes"]["Cooling Circuit Label"]=="-":
                    fgroup_item = get_value_by_path(data,path[:-4])  
                    del fgroup_item[path[-4]]
        except Exception as e:
            self.log.error(f"Error removing cooling circuits: {e}")
            
        self.log.info("Processing cooling sheet items...")
        try:
            all_paths = find_key_paths(data,"Sheet Item")
            paths = extract_paths("FUNCTIONAL_GROUP_2",3,all_paths)
            for path in paths:
                sheet_item = get_value_by_path(data,path)
                missing_items = [" 1.5", " 1.6", " 1.7", " 1.8"," 3.9"," 3.10"," 3.11"," 3.12"," 5.9"," 5.10", " 5.11", " 5.12", " 7.9", " 7.10", " 7.11", " 7.12"," 9.5", " 9.6", " 9.7", " 9.8"," 10.3"," 10.4"," 11.2"," 11.3"," 12.3"," 12.4"," 12.5"]
                missing_content = {'Element Attributes': {'Set': 0}, 'Element Item': {'Connection': {}}}
                for item in missing_items:
                    sheet_item[item] = copy.deepcopy(missing_content)
                #wenn 1.* or 3.* or 5.* or 7.* or 9.* ; * e [1:4] {"Set":"+"} is True then add Pipe Tee" to Element Item
                ids = [" 3."," 5."," 7.", " 9."]
                poss = ["1","2","3","4"]
                for id in ids:
                    for pos in poss:
                        if sheet_item[id+pos]["Element Attributes"]["Set"]=="+":
                            #sheet_item[id+pos]["Element Item"]["Pipe Tee"]={}#
                            sheet_item[id+str(int(pos)+12)]={"Element Attributes" : {},"Element Item":{"Pipe Tee":{}}}

                ids = [" 1."]
                poss = ["9","10","11","12"]
                for id in ids:
                    for pos in poss:
                        if sheet_item[id+pos]["Element Attributes"]["Set"]=="+":
                            sheet_item[id+str(int(pos)+4)]={"Element Attributes" : {},"Element Item":{"Pipe Tee":{}}}

                #wenn 1.1 {"Set":"+"} 0.5 {"Set":"+"} add Pipe , 1.2 -> 0.6, 1.3 -> 0.7, 1.4 -> 0.8
                id1 = " 1."
                id2 = " 1."
                poss1 = ["9","10","11","12"]
                poss2 = ["5","6","7","8"]
                for pos1, pos2 in zip(poss1,poss2):
                    if sheet_item[id1+pos1]["Element Attributes"]["Set"]=="+":
                        sheet_item[id2+pos2]["Element Attributes"]["Set"]="+"

                #wenn 4.* or 6.* or 8.*; * e [1:4] dann 3.*, 5.*, 7.*, * e [9:12] {"Set":"+"} 
                ids1 = [" 4.", " 6.", " 8."]
                ids2 = [" 3.", " 5.", " 7."]
                poss1 = ["1","2","3","4"]
                poss2 = ["9","10","11","12"]
                for id1,id2 in zip(ids1,ids2):
                    for pos1,pos2 in zip(poss1,poss2):
                        if list(sheet_item[id1+pos1]["Element Item"].keys())[0]=="-" or list(sheet_item[id1+pos1]["Element Item"].keys())[0]=="0":
                            pass
                        else:
                            sheet_item[id2+pos2]["Element Attributes"]["Set"]="+"

                #wenn 9.*, * e [1,4] {"Set":"+"} dann 9.* e [5,8] {"Set":"+"}
                id = " 9."

                poss1 = ["1","2","3","4"]
                poss2 = ["5","6","7","8"]
                for pos1,pos2 in zip(poss1,poss2):
                    if sheet_item[id+pos1]["Element Attributes"]["Set"]=="+":
                        sheet_item[id+pos2]["Element Attributes"]["Set"]="+"

                #wenn 3.* or 5.* or 7.* , * e [1,4] {"Set":"+"}  und 3.* or 5.* or 7.* , * e [9,12] {"Set":"+"} dann 3.* or 5.* or 7.* , * e [5,8] {"Set":"+"}
                #ids = [" 3.", " 5.", " 7."]
                #poss1 = ["1","2","3","4"]
                #poss2 = ["9","10","11","12"]
                #poss3 = ["5","6","7","8"]

                #for id in ids:
                #    for pos1,pos2,pos3 in zip(poss1,poss2,poss3):
                #        if sheet_item[id+pos1]["Element Attributes"]["Set"]=="+" and sheet_item[id+pos2]["Element Attributes"]["Set"]=="":
                #            sheet_item[id+pos3]["Element Attributes"]["Set"]="+"

                #add Heat Exchanger
                #add Seawater
                #wenn Heat Exchanger -> Seawater Pump
                group = get_value_by_path(data,path[:-3])
                sheet_item[" 10.1"]={'Element Item': {'Pump': {}}}
                sheet_item[" 12.1"]= {'Element Item': {group["Group Attributes"]["External Circuit"]: {}}}
                if group["Group Attributes"]["External Circuit"]=="Heat Exchanger":
                    sheet_item[" 12.2"]={'Element Item': {'Pump': {}}}  
                    sheet_item[" 12.6"]={'Element Item': {'Seawater': {}}}
                sheet_item[" 10.2"]= {'Element Item': {'Pipe Tee': {}}}


                #add connection items
                ids=[" 1."," 3."," 5."," 7."," 9."]
                poss1=["5","6","7","8"]
                for id in ids:
                    for pos1 in poss1:
                        if sheet_item[id+pos1]["Element Attributes"]["Set"]=="+":
                            sheet_item[id+pos1]["Element Attributes"]["CON 1 Sheet Item"]=id+str(int(pos1)+8)                                           
                            sheet_item[id+pos1]["Element Attributes"]["CON 1 Port"]="S"

                            sheet_item[id+pos1]["Element Attributes"]["CON 1 Part"]=list(sheet_item[sheet_item[id+pos1]["Element Attributes"]["CON 1 Sheet Item"]]["Element Item"].keys())[0]
                            if id+pos1==" 1.5":
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Sheet Item"]=" 10.1"       
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Port"]="N"    
                            elif pos1=="5":
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Sheet Item"]=" 12.1"       
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Port"]="B" 
                            else:
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Sheet Item"]=id+str(int(pos1)+7)                                           
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Port"]="N"        

                            sheet_item[id+pos1]["Element Attributes"]["CON 2 Part"]=list(sheet_item[sheet_item[id+pos1]["Element Attributes"]["CON 2 Sheet Item"]]["Element Item"].keys())[0]

                ids=[" 1."," 3."," 5."," 7."]
                poss1=["9","10","11","12"]
                for id in ids:
                    for pos1 in poss1:
                        if sheet_item[id+pos1]["Element Attributes"]["Set"]=="+":
                                sheet_item[id+pos1]["Element Attributes"]["CON 1 Sheet Item"]=id+str(int(pos1)+4)                                           
                                sheet_item[id+pos1]["Element Attributes"]["CON 1 Port"]="E"        
                                sheet_item[id+pos1]["Element Attributes"]["CON 1 Part"]=list(sheet_item[sheet_item[id+pos1]["Element Attributes"]["CON 1 Sheet Item"]]["Element Item"].keys())[0]
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Sheet Item"]=" "+str(1+int(id[-2]))+"."+str(int(pos1)-8)                                           
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Port"]="W"

                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Part"]=list(sheet_item[sheet_item[id+pos1]["Element Attributes"]["CON 2 Sheet Item"]]["Element Item"].keys())[0]


                ids=[" 3."," 5."," 7."," 9."]
                poss1=["1","2","3","4"]
                for id in ids:
                    for pos1 in poss1:
                        if sheet_item[id+pos1]["Element Attributes"]["Set"]=="+":
                                sheet_item[id+pos1]["Element Attributes"]["CON 1 Sheet Item"]=" "+str(-1+int(id[-2]))+"."+pos1                                         
                                sheet_item[id+pos1]["Element Attributes"]["CON 1 Port"]="E"   
                                sheet_item[id+pos1]["Element Attributes"]["CON 1 Part"]=list(sheet_item[sheet_item[id+pos1]["Element Attributes"]["CON 1 Sheet Item"]]["Element Item"].keys())[0]     
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Sheet Item"]=id+str(int(pos1)+12)                                            
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Port"]="W"
                                sheet_item[id+pos1]["Element Attributes"]["CON 2 Part"]=list(sheet_item[sheet_item[id+pos1]["Element Attributes"]["CON 2 Sheet Item"]]["Element Item"].keys())[0]


                #find visually highest pipe tee above pipe tee connecting to connection 9.5
                ids = [" 3."," 5."," 7."," 9."]
                topright = ""
                for id in ids:
                    if sheet_item[id+"5"]["Element Attributes"]["Set"]=="+":
                        poss = ["16","15","14","13"]
                        for pos in poss:
                            if id+pos in sheet_item.keys():
                                topright = id+pos
                                break



                        


                #remove unnecessary Pipe Tees interrupting Connections
                ids = [" 3."," 5."," 7."]
                poss = ["1","2","3","4"]
                for id in ids:
                    for pos in poss:
                        # connect devices directly when no pipe tee is necessary
                        if sheet_item[id+pos]["Element Attributes"]["Set"]=="+":
                            if sheet_item[id+str(int(pos)+8)]["Element Attributes"]["Set"]=="+" and sheet_item[id+str(int(pos)+4)]["Element Attributes"]["Set"]==0:
                                    if (pos == "1" or pos =="2" or pos=="3"):
                                        if sheet_item[id+str(int(pos)+5)]["Element Attributes"]["Set"]==0:
                                            sheet_item[id+pos]["Element Attributes"]["CON 2 Sheet Item"]=sheet_item[id+str(int(pos)+8)]["Element Attributes"]["CON 2 Sheet Item"]                                      
                                            sheet_item[id+pos]["Element Attributes"]["CON 2 Port"]=sheet_item[id+str(int(pos)+8)]["Element Attributes"]["CON 2 Port"]
                                            sheet_item[id+pos]["Element Attributes"]["CON 2 Part"]=sheet_item[id+str(int(pos)+8)]["Element Attributes"]["CON 2 Part"]
                                        
                                            sheet_item[id+str(int(pos)+8)]["Element Attributes"]["Set"]=0
                                            sheet_item[id+str(int(pos)+12)]["Element Attributes"]["Set"]=0
                                    else:
                                            sheet_item[id+pos]["Element Attributes"]["CON 2 Sheet Item"]=sheet_item[id+str(int(pos)+8)]["Element Attributes"]["CON 2 Sheet Item"]                                      
                                            sheet_item[id+pos]["Element Attributes"]["CON 2 Port"]=sheet_item[id+str(int(pos)+8)]["Element Attributes"]["CON 2 Port"]
                                            sheet_item[id+pos]["Element Attributes"]["CON 2 Part"]=sheet_item[id+str(int(pos)+8)]["Element Attributes"]["CON 2 Part"]
                                        
                                            sheet_item[id+str(int(pos)+8)]["Element Attributes"]["Set"]=0 
                                            sheet_item[id+str(int(pos)+12)]["Element Attributes"]["Set"]=0                               
                for id in ids:
                    for pos in poss:
                        # connect directly to pipe tee when line goes one layer downward
                        if sheet_item[id+pos]["Element Attributes"]["Set"]=="+":
                            if sheet_item[id+str(int(pos)+8)]["Element Attributes"]["Set"]==0 and sheet_item[id+str(int(pos)+4)]["Element Attributes"]["Set"]=="+":
                                #ignore top-right pipe tee
                                if id+str(int(pos)+12) != topright:
                                    if (pos == "1" or pos =="2" or pos=="3"):    
                                        if sheet_item[id+str(int(pos)+5)]["Element Attributes"]["Set"]==0:
                                                sheet_item[id+pos]["Element Attributes"]["CON 2 Sheet Item"]=sheet_item[id+str(int(pos)+4)]["Element Attributes"]["CON 2 Sheet Item"]                                      
                                                sheet_item[id+pos]["Element Attributes"]["CON 2 Port"]=sheet_item[id+str(int(pos)+4)]["Element Attributes"]["CON 2 Port"]
                                                sheet_item[id+pos]["Element Attributes"]["CON 2 Part"]=sheet_item[id+str(int(pos)+4)]["Element Attributes"]["CON 2 Part"]

                                                sheet_item[id+str(int(pos)+12)]["Element Attributes"]["Set"]=0
                                                sheet_item[id+str(int(pos)+4)]["Element Attributes"]["Set"]=0      
                                    else:
                                                sheet_item[id+pos]["Element Attributes"]["CON 2 Sheet Item"]=sheet_item[id+str(int(pos)+4)]["Element Attributes"]["CON 2 Sheet Item"]                                      
                                                sheet_item[id+pos]["Element Attributes"]["CON 2 Port"]=sheet_item[id+str(int(pos)+4)]["Element Attributes"]["CON 2 Port"]
                                                sheet_item[id+pos]["Element Attributes"]["CON 2 Part"]=sheet_item[id+str(int(pos)+4)]["Element Attributes"]["CON 2 Part"]

                                                sheet_item[id+str(int(pos)+12)]["Element Attributes"]["Set"]=0   
                                                sheet_item[id+str(int(pos)+4)]["Element Attributes"]["Set"]=0                              
                
                #remove unwanted device items
                ids = [" 2."," 4."," 6."," 8."]
                poss = ["1","2","3","4"]       
                for id in ids:
                    for pos in poss:
                        # connect directly to next connectable item
                        if list(sheet_item[id+pos]["Element Item"].keys())[0]=="+":
                            del sheet_item[id+pos]
                            done = 0
                            id1=id
                            while done==0:
                                if sheet_item[" "+str(int(id1[-2])-1)+"."+str(int(pos)+8)]["Element Attributes"]["Set"]=="+":
                                    sheet_item[" "+str(int(id1[-2])-1)+"."+str(int(pos)+8)]["Element Attributes"]["CON 2 Sheet Item"]=sheet_item[" "+str(int(id[-2])+1)+"."+pos]["Element Attributes"]["CON 2 Sheet Item"]                                      
                                    sheet_item[" "+str(int(id1[-2])-1)+"."+str(int(pos)+8)]["Element Attributes"]["CON 2 Port"]=sheet_item[" "+str(int(id[-2])+1)+"."+pos]["Element Attributes"]["CON 2 Port"]
                                    sheet_item[" "+str(int(id1[-2])-1)+"."+str(int(pos)+8)]["Element Attributes"]["CON 2 Part"]=sheet_item[" "+str(int(id[-2])+1)+"."+pos]["Element Attributes"]["CON 2 Part"]

                                    sheet_item[" "+str(int(id[-2])+1)+"."+pos]["Element Attributes"]["Set"]=0
                                    done = 1

                                if sheet_item[" "+str(int(id1[-2])-1)+"."+pos]["Element Attributes"]["Set"]==0 and sheet_item[" "+str(int(id[-2])-1)+"."+pos]["Element Attributes"]["Set"]=="+":
                                    sheet_item[" "+str(int(id1[-2])-1)+"."+pos]["Element Attributes"]["CON 2 Sheet Item"]=sheet_item[" "+str(int(id[-2])+1)+"."+pos]["Element Attributes"]["CON 2 Sheet Item"]                                      
                                    sheet_item[" "+str(int(id1[-2])-1)+"."+pos]["Element Attributes"]["CON 2 Port"]=sheet_item[" "+str(int(id[-2])+1)+"."+pos]["Element Attributes"]["CON 2 Port"]
                                    sheet_item[" "+str(int(id1[-2])-1)+"."+pos]["Element Attributes"]["CON 2 Part"]=sheet_item[" "+str(int(id[-2])+1)+"."+pos]["Element Attributes"]["CON 2 Part"]

                                    sheet_item[" "+str(int(id[-2])+1)+"."+pos]["Element Attributes"]["Set"]=0
                                    done = 1
                                id1 = " "+str(int(id1[-2])-2)+"."

                #add connections between  pump and heat exchanger
                        #add connection between tank and topright pipe tee
                poss = [" 10.3"," 10.4"," 11.2"," 11.3"]#11.3 fehlt
                cons = [[" 10.1"," 10.2"],[" 10.2"," 12.1"],[" 10.2"," 11.1"],[topright," 11.1"]]
                ports =[["S","N"],["E","A"],["W","S"],["N","E"]]
                for pos,con,port in zip(poss,cons,ports):
                    sheet_item[pos]["Element Attributes"]["CON 1 Sheet Item"]=con[0]                                   
                    sheet_item[pos]["Element Attributes"]["CON 1 Port"]=port[0]      
                    sheet_item[pos]["Element Attributes"]["CON 1 Part"]=list(sheet_item[con[0]]["Element Item"].keys())[0]

                    sheet_item[pos]["Element Attributes"]["CON 2 Sheet Item"]=con[1]                                
                    sheet_item[pos]["Element Attributes"]["CON 2 Port"]=port[1]
                    sheet_item[pos]["Element Attributes"]["CON 2 Part"]=list(sheet_item[con[1]]["Element Item"].keys())[0]
                    sheet_item[pos]["Element Attributes"]["Set"]="+"

                
                poss = [" 10.3"," 10.4"," 11.2"]#," 12.4"," 12.5"]#11.3 fehlt
                for pos in poss:
                    sheet_item[pos]["Element Attributes"]["Set"]="+"


                #add connections between Heat Exchanger and Seawater
                if group["Group Attributes"]["External Circuit"]=="Heat Exchanger":
                    poss = [" 12.3"," 12.4"," 12.5"]#11.3 fehlt
                    cons = [[" 12.1"," 12.2"],[" 12.2"," 12.6"],[" 12.1"," 12.6"]]
                    ports =[["C","N"],["S","W"],["D","E"]]
                    for pos,con,port in zip(poss,cons,ports):
                        sheet_item[pos]["Element Attributes"]["CON 1 Sheet Item"]=con[0]                                   
                        sheet_item[pos]["Element Attributes"]["CON 1 Port"]=port[0]      
                        sheet_item[pos]["Element Attributes"]["CON 1 Part"]=list(sheet_item[con[0]]["Element Item"].keys())[0]

                        sheet_item[pos]["Element Attributes"]["CON 2 Sheet Item"]=con[1]                                
                        sheet_item[pos]["Element Attributes"]["CON 2 Port"]=port[1]
                        sheet_item[pos]["Element Attributes"]["CON 2 Part"]=list(sheet_item[con[1]]["Element Item"].keys())[0]
                        sheet_item[pos]["Element Attributes"]["Set"]="+"

        except Exception as e:
            self.log.error(f"Error processing cooling sheet items: {e}")

        self.log.info("Cooling sheet processing completed.")
    def cleanup(self):
        self.log.info(f"Starting cleanup...")
        #remove all unnecessary structure
        data = self.data
        try:
            all_paths = find_key_paths(data,"Sheet Item") 
            paths = extract_paths("FUNCTIONAL_GROUP_2",3,all_paths)
            for path in paths:
                sheet_item = get_value_by_path(data,path)
                for key in list(sheet_item.keys()):
                    try:
                        if "Element Attributes" in sheet_item[key]:
                            if "Set" in sheet_item[key]["Element Attributes"]:
                                if sheet_item[key]["Element Attributes"]["Set"]==0:
                                    del sheet_item[key]
                        if key in sheet_item:
                            if "Element Item" in sheet_item[key]:
                                if list(sheet_item[key]["Element Item"].keys())[0]== "0":
                                    del sheet_item[key]
                    except Exception as e:
                        self.log.error(f"Error during cleanup: {e}")

            self.log.info("Cleanup completed.")
        except Exception as e:
            self.log.error(f"Error during cleanup: {e}")

        self.log.info("Cleanup completed.")

if __name__ == "__main__":

     h = TransJson()
     h.transformJson()