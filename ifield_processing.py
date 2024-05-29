import sys
import os
import shutil
import pandas as pd
import numpy as np
from tqdm import tqdm
import re
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

main_protoid_final = int(config["main"]["protoid_final"])

isurveys = {}

csv_files = glob.glob(os.path.join("source\\csv", "*.csv"))
csv_files = sorted(csv_files, key=lambda x: os.path.getctime(x), reverse=False)

for csv_file in csv_files:
    df = pd.read_csv(csv_file, encoding="utf-8", low_memory=False)
    
    for proto_id in list(np.unique(list(df["ProtoSurveyID"]))):
        if proto_id not in isurveys.keys():
            isurveys[proto_id] = {
                "csv" : csv_file,
                "survey" : None
            }

#Read the xml file for the main section
for proto_id, xml_file in config["main"]["xmls"].items():
    isurveys[int(proto_id)]["survey"] = iSurvey(f'source\\xml\\{xml_file}') 

#Read the xml file for the placement + recall section
for stage_id, stage_obj in config["stages"].items():
    for proto_id, xml_file in stage_obj["xmls"] .items():
        isurveys[int(proto_id)]["survey"] = iSurvey(f'source\\xml\\{xml_file}') 

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

    for question_name, question in isurveys[main_protoid_final]["survey"]["questions"].items():
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

try:
    if not os.path.isfile(current_mdd_file) or not os.path.isfile(current_ddf_file):
        raise Exception("File Error", "File mdd/ddf is not exist.")

    #Open the update data file
    df_update_data = pd.read_csv("source\\update_data.csv", encoding="utf-8")

    if not df_update_data.empty:
        df_update_data.set_index(["InstanceID"], inplace=True)

    adoConn = w32.Dispatch('ADODB.Connection')
    conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(current_ddf_file, current_mdd_file)
    adoConn.Open(conn)

    #Delete all data before inserting new data (default is FALSE)
    if config["source_initialization"]["delete_all"]:
        sql_delete = "DELETE FROM VDATA"
        adoConn.Execute(sql_delete)
        adoConn.Close()

    #csv_files = glob.glob(os.path.join("source\\csv", "*.csv"))
    #csv_files = sorted(csv_files, key=lambda x: os.path.getctime(x), reverse=True)

    for proto_id, xml_file in config["main"]["xmls"].items():

        df = pd.read_csv(isurveys[int(proto_id)]["csv"], encoding="utf-8", low_memory=False)
        df.set_index(['InstanceID'], inplace=True)

        #Allow inserting dummy data
        if config["source_initialization"]["dummy_data_required"]:
            df = df.loc[df["System_LocationID"] == "_DefaultSP"]
        else:
            df = df.loc[df["System_LocationID"] != "_DefaultSP"]
            
        m = Metadata(mdd_file=current_mdd_file, ddf_file=current_ddf_file, sql_query="SELECT InstanceID FROM VDATA")
        df_instancedis = m.convertToDataFrame(questions=["InstanceID"])
        
        if not df_instancedis.empty:
            df_instancedis.set_index(['InstanceID'], inplace=True)
        
        ids = [id for id in list(df.index) if str(id) not in list(df_instancedis.index)]

        if len(ids) > 0:
            df_data = df.loc[ids]

            #Allow inserting dummy data
            if config["source_initialization"]["dummy_data_required"]:
                df_data = df_data.loc[df_data["System_LocationID"] == "_DefaultSP"]
            else:
                df_data = df_data.loc[df_data["System_LocationID"] != "_DefaultSP"]

            if not df_data.empty:
                adoConn.Open(conn)

                for i, row in df_data[list(df_data.columns)].iterrows():
                    try:
                        isurvey = isurveys[int(row["ProtoSurveyID"])]["survey"]
                    except Exception as ex:
                        raise Exception("Config Error", "ProtoID {} should be declare in the config file.".format(str(row["ProtoSurveyID"])))
                        
                    sql_insert = "INSERT INTO VDATA(InstanceID) VALUES(%s)" % (row.name)
                    adoConn.Execute(sql_insert)
                    
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
                                                    v.append("'{}'".format(". ".join(re.split('\n|\r', str(row[csv_obj["csv"][0]])))))
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

                    try:
                        df_update_data_by_id = df_update_data[["Question Name","Current Value"]].loc[row.name]

                        if not df_update_data_by_id.empty:
                            sql_update = "UPDATE VDATA SET %s WHERE InstanceID = %s" % (
                                "%s = %s" % (
                                    df_update_data_by_id[0], 
                                    'NULL' if pd.isnull(df_update_data_by_id[1]) else (df_update_data_by_id[1] if str(df_update_data_by_id[1]).isnumeric() else "'%s'" % (df_update_data_by_id[1]))
                                ) if isinstance(df_update_data_by_id,pd.Series) else ",".join(["%s = %s" % (
                                    x[0],
                                    'NULL' if pd.isnull(x[1]) else (x[1] if str(x[1]).isnumeric() else "'%s'" % (x[1]))
                                ) for x in [tuple(x) for x in df_update_data[["Question Name","Current Value"]].loc[row.name].to_numpy()]]),
                                row.name
                            )
                            adoConn.Execute(sql_update)
                    except:
                        continue

                adoConn.Close()
    
    #Delete all data before inserting new data (default is FALSE)
    if config["source_initialization"]["remove_all_ids"]:
        adoConn.Open(conn)

        sql_delete = "DELETE FROM VDATA WHERE Not _LoaiPhieu.ContainsAny({_1,_5})"
        adoConn.Execute(sql_delete)
        adoConn.Close()

except Exception as error:
    print(repr(error))
    #sys.exit(repr(error))