
import glob
import pandas as pd
import json
import logging

def set_nested_value(dic, keys, value):
    """
    Sets a value in a nested dictionary using a list of keys.

    Parameters:
        dic (dict): The dictionary to modify.
        keys (list): A list of keys representing the path in the nested dictionary.
        value: The value to set.

    Example:
        dic = {}
        set_nested_value(dic, ["world", "a", "b", "c"], "hello")
        print(dic)  # Output: {'world': {'a': {'b': {'c': 'hello'}}}}
    """
    temp = dic
    for key in keys[:-1]:  # Traverse to the second-to-last key
        temp = temp.setdefault(key, {})  # Create intermediate dictionaries if they don't exist
    temp[keys[-1]] = value  # Set the value at the final key      

 



class Config2Json:

    def __init__(self,**kwargs):
        xlsx_files = glob.glob("*.xlsm")
        # first found excel file is assumed to be config file
        self.configFile = xlsx_files[0]
        self.configFile = "conf.xlsm"

        if "log" in kwargs:
            self.log=kwargs["log"]
        else:
            self.log = logging.getLogger(__name__)
            logging.basicConfig(filename="config2json.log",format='%(asctime)s %(message)s',level=logging.DEBUG)


    def getAutoDrawingConfig(self):
        pass

    def getIndex(self,base,el):
        ix = base.index(el)
        if ix >4:
            ix=5
        return ix



    def populateJson(self):
            self.log.info("Start Config File Import")
            self.df = pd.read_excel(self.configFile, "json_prep", header=None)
            self.df=self.df.loc[8:]
            self.project = {}
            self.base = {
                 "Project Attributes" : {},
                 "Project Item" : {
                      "Functional Group Attributes" : {},
                      "Functional Group Item" : {
                           "Sheet Attributes" : {},
                           "Sheet Item" : {
                                "Group Attributes" : {},
                                "Group Items" : {
                                     "Element Attributes" : {},
                                     "Element Items" : {}
                                }
                           }
                      }
                 }
            }
            self.h_base = ["Project", "Functional Group", "Group", "Sheet", "Element","Part","Text"]
            element_old = "Project"
            higher_level = []
            level = []
            self.property = ["Attributes","Item"]
            for element in self.df.iterrows():
                line = list(element[1][3:])
                for element in self.h_base:
                     for property in self.property:
                        key = element+" "+property
                        if  line[0]==key:
                            if self.getIndex(self.h_base,element) - self.getIndex(self.h_base,element_old) == 1:
                                level = nested_key_old
                                if pd.isna(line[1])==False:
                                    nested_key = level + [key] + [line[1]]
                                else:
                                    nested_key = level + [key]

                            if self.getIndex(self.h_base,element) - self.getIndex(self.h_base,element_old)  == 0:
                                if pd.isna(line[1])==False:
                                    nested_key = level+[key] + [line[1]]
                                else:
                                    nested_key = level+[key]
 
                            if self.getIndex(self.h_base,element) - self.getIndex(self.h_base,element_old)  < 0:
                                delta = -2+2*(self.getIndex(self.h_base,element) - self.getIndex(self.h_base,element_old) )
                                level = nested_key_old[:delta]
                                if pd.isna(line[1])==False:
                                    nested_key = level + [key] +[line[1]]
                                else:
                                    nested_key = level + [key]
                                    #higher_levels = level[:delta]                             
                            if pd.isna(line[1]) == False and pd.isna(line[2]) == True:
                                value = {}
                            elif pd.isna(line[1]) == False and pd.isna(line[2]) == False:
                                value = line[2]
                            if pd.isna(line[1]) == True and pd.isna(line[2]) == True:
                                value = {}
                            #print("higher levels"+ str(higher_levels))  
                            self.log.info("Import Key: %s", str(nested_key))
                            self.log.info("Import Value: %s",str(value))  
                            set_nested_value(self.project,nested_key,value)                           
                            element_old = element
                            nested_key_old = nested_key
                            break
            with open("data.json", "w") as file:
                json.dump(self.project, file, indent=4)
            self.log.info("Config imported into json File %s","data.json")



            
if __name__ == "__main__":

     h = Config2Json()
     h.populateJson()

     





# json transformer nochmal bauen
# json2zuken bauen   