import os
import shutil
import pandas as pd
import numpy as np
from tqdm import tqdm
import json
import glob
import win32com.client as w32
import xml.etree.ElementTree as ET
from object.metadata import Metadata
from object.enumerations import dataTypeConstants
from object.iSurvey import iSurvey

import collections.abc

#hyper needs the four following aliases to be done manually.
collections.Iterable = collections.abc.Iterable
collections.Mapping = collections.abc.Mapping
collections.MutableSet = collections.abc.MutableSet
collections.MutableMapping = collections.abc.MutableMapping

#Load config
f = open('config.json', mode = 'r', encoding="utf-8")
config = json.loads(f.read())
f.close()

main_protoid_final = config["main"]["protoid_final"]

isurveys = {}

#Đọc file xml của phần main
for proto_id, xml_file in config["main"]["xmls"].items(): 
    isurveys[proto_id] = iSurvey(f'source\\xml\\{xml_file}') 


if config["run_mdd_source"]:
    #Create mdd/ddf file based on xmls file
    source_mdd_file = r"template\TemplateProject.mdd"
    current_mdd_file = "data\\{}.mdd".format(config["project_name"])
    source_dms_file = r"dms\OutputDDFFile.dms"

    mdd_files = glob.glob(os.path.join("data", "*.mdd"))
    ddf_files = glob.glob(os.path.join("data", "*.ddf"))

    for f in mdd_files:
        os.remove(f)
    for f in ddf_files:
        os.remove(f)

    if not os.path.exists(current_mdd_file):
        shutil.copy(source_mdd_file, current_mdd_file)

    mdd_source = Metadata(mdd_file=current_mdd_file, dms_file=source_dms_file)

    mdd_source.addScript("InstanceID", "InstanceID \"InstanceID\" text;")

    for question_name, question in isurveys[main_protoid_final]["questions"].items():
        if "syntax" in question.keys():
            
            if question["attributes"]["objectName"] in ["SHELL_BLOCK"]:
                a = ""

            mdd_source.addScript(question["attributes"]["objectName"], question["syntax"], is_defined_list=question["is_defined_list"], parent_nodes=list() if "parents" not in question.keys() else [q["attributes"]["objectName"] for q in question["parents"]])

            if "comment_syntax" in question.keys():
                mdd_source.addScript(f'{question["attributes"]["objectName"]}{question["comment"]["objectName"]}', question["comment_syntax"], parent_nodes=list() if "parents" not in question.keys() else [q["attributes"]["objectName"] for q in question["parents"]])

    mdd_source.runDMS()

###################################################################################

current_mdd_file = "data\\{}_EXPORT.mdd".format(config["project_name"])
current_ddf_file = "data\\{}_EXPORT.ddf".format(config["project_name"])

df_main = pd.read_csv(f'source\\csv\\{config["main"]["csv"]}', encoding="utf-8")
df_main.set_index(['InstanceID'], inplace=True)

adoConn = w32.Dispatch('ADODB.Connection')
conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(current_ddf_file, current_mdd_file)
adoConn.Open(conn)

sql_delete = "DELETE FROM VDATA"
adoConn.Execute(sql_delete)

for i, row in df_main[list(df_main.columns)].iterrows():
    try:
        sql_insert = "INSERT INTO VDATA(InstanceID) VALUES(%s)" % (row.name)
        adoConn.Execute(sql_insert)

        isurvey = isurveys[str(row["ProtoSurveyID"])]

        c = list()
        v = list()

        for key, question in isurvey["questions"].items():
            if question["attributes"]["objectName"] in ["Q6a_Text"]:
                if row.name == 1609219:
                    a = ""

            if question["datatype"] not in [dataTypeConstants.mtNone, dataTypeConstants.mtLevel]:
                for i in range(len(question["columns"])):
                    for mdd_col, csv_obj in question["columns"][i].items():
        
                        if (question["datatype"].value == dataTypeConstants.mtCategorical.value) or (question["datatype"].value == dataTypeConstants.mtObject.value and csv_obj["datatype"].value == dataTypeConstants.mtCategorical.value):
                            if bool(int(question["answers"]["answerref"]["attributes"]["isMultipleSelection"])):
                                if row[csv_obj["csv"]].sum() > 0:
                                    c.append(mdd_col)
                                    v.append("{%s}" % (",".join([k.split(".")[-1] for k, v in dict(row[csv_obj["csv"]]).items() if v == 1])))
                            else:
                                if not pd.isnull(row[csv_obj["csv"][0]]): 
                                    c.append(mdd_col)
                                    v.append("{%s}" % (question["answers"]["options"][str(int(row[csv_obj["csv"][0]]))]["objectname"]))
                            
                            for mdd_other_col, csv_other_obj in csv_obj["others"].items():

                                if not pd.isnull(row[csv_other_obj["csv"][0]]):
                                    c.append(mdd_other_col)

                                    match int(csv_other_obj["datatype"]):
                                        case 2:
                                            v.append("{}".format(row[csv_other_obj["csv"][0]]))
                                        case 3:
                                            v.append("'{}'".format(row[csv_other_obj["csv"][0]]))
                                        case 4:
                                            v.append("'{}'".format(row[csv_other_obj["csv"][0]]))
                        else:
                            if not pd.isnull(row[csv_obj["csv"][0]]):
                                match question["datatype"].value:
                                    case dataTypeConstants.mtText.value:
                                        c.append(mdd_col)
                                        v.append("'{}'".format(". ".join(str(row[csv_obj["csv"][0]]).split('\n'))))
                                    case dataTypeConstants.mtDate.value:
                                        c.append(mdd_col)
                                        v.append("'{}'".format(row[csv_obj["csv"][0]]))
                                    case dataTypeConstants.mtDouble.value:
                                        c.append(mdd_col)
                                        v.append("{}".format(row[csv_obj["csv"][0]]))
                                    case dataTypeConstants.mtObject.value:
                                        c.append(mdd_col)
                                        v.append("{}".format(row[csv_obj["csv"][0]]))
        
        sql_update = "UPDATE VDATA SET " + ','.join([cx + str(r" = %s") for cx in c]) % tuple(v) + " WHERE InstanceID = {}".format(row.name)
        adoConn.Execute(sql_update)    
    except Exception as ex:
        #print(sql_insert, ex, sep="-")
        sys.exit(1)