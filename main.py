import config2json
import json2json
import json2zuken
import pandas as pd
import logging
import os

excel_file = "conf.xlsm"
df = pd.read_excel(excel_file, "Run", header=None)


config2json_flag = int(df.loc[11,13])
json2json_flag = int(df.loc[12,13])
json2zuken_flag = int(df.loc[13,13])

filename = "config2zuken.log"
filepath = os.path.join(os.getcwd(), filename)
if os.path.isfile(filepath):
    os.remove(filepath)

logger = logging.getLogger(__name__)
logging.basicConfig(filename="config2zuken.log",format='%(asctime)s %(message)s',level=logging.DEBUG)

if config2json_flag==1:
    print("importing config file...")
    for _ in range(10):
        logger.info("hi")
    h1 = config2json.Config2Json(log=logger)
    h1.populateJson()
    print("import done.")

if json2json_flag==1:
    print("preparing project for drawing...")
    for _ in range(10):
        logger.info("")
    h2 = json2json.TransJson(log=logger)
    h2.transformJson()
    print("preparation done")

if json2zuken_flag==1:
    print("drawing...")
    for _ in range(10):
        logger.info("")
    h3 = json2zuken.Draw(log=logger)
    h3.draw()
    print("drawing done.")