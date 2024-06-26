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


try:
    isurveys = {}

    csv_files = glob.glob(os.path.join("source\\csv", "*.csv"))
    csv_files = sorted(csv_files, key=lambda x: os.path.getctime(x), reverse=False)

    for csv_file in tqdm(csv_files, desc="Read the csv file"):
        df = pd.read_csv(csv_file, encoding="utf-8", low_memory=False)
        
        for proto_id in list(np.unique(list(df["ProtoSurveyID"]))):
            if proto_id not in isurveys.keys():
                isurveys[proto_id] = {
                    "csv_files" : [csv_file],
                    "survey" : None
                }
            else:
                isurveys[proto_id]["csv_files"].append(csv_file)
    
    if int(config["main"]["protoid_final"]) not in isurveys.keys():
        raise Exception("Config Error: " "ProtoID {} has no data in the CSV file".format(int(config["main"]["protoid_final"])))

    main_protoid_final = int(config["main"]["protoid_final"])

    #Read the xml file for the main section
    try:
        for proto_id, xml_file in tqdm(config["main"]["xmls"].items(), desc="Convet the xml file for the main section"):
            if os.path.exists(f'source\\xml\\{xml_file}'):
                isurveys[int(proto_id)]["survey"] = iSurvey(f'source\\xml\\{xml_file}') 
            else:
                print("Config Warning: ", "ProtoID {} should be declare in the config file.".format(proto_id))
    except:
        raise Exception("Config Error: ", "ProtoID {} has no data in the CSV file.".format(proto_id))
    
    #Read the xml file for the placement + recall section
    follow_up_questions = []

    for stage_id, stage_obj in tqdm(config["stages"].items(), desc="Convet the xml file for the placement + recall section"):
        try:
            for proto_id, xml_file in stage_obj["xmls"] .items():
                if os.path.exists(f'source\\xml\\{xml_file}'):
                    isurveys[int(proto_id)]["survey"] = iSurvey(f'source\\xml\\{xml_file}') 
                else:
                    print("Config Warning: ", "ProtoID {} should be declare in the config file.".format(proto_id))
        except:
            raise Exception("Config Error: ", "ProtoID {} has no data in the CSV file.".format(proto_id))
        
        if len(follow_up_questions) > 0:
            follow_up_questions.extend([{a : 0}  for a in list(isurveys[int(proto_id)]["survey"]["questions"].keys()) if a not in follow_up_questions]) 

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

        mdd_source = Metadata(mdd_file=current_mdd_file, dms_file=source_dms_file, default_language=config["source_initialization"]["default_language"])

        mdd_source.addScript("InstanceID", "InstanceID \"InstanceID\" text;")

        for question_name, question in tqdm(isurveys[main_protoid_final]["survey"]["questions"].items(), desc="Convert the mdd/ddf file for the main question"):
            if "syntax" in question.keys():
                
                if question["attributes"]["objectName"] in ["SHELL_BLOCK"]:
                    a = ""

                mdd_source.addScript(question["attributes"]["objectName"], question["syntax"], is_defined_list=question["is_defined_list"], parent_nodes=list() if "parents" not in question.keys() else [q["attributes"]["objectName"] for q in question["parents"]])

                if "comment_syntax" in question.keys():
                    mdd_source.addScript(f'{question["attributes"]["objectName"]}{question["comment"]["objectName"]}', question["comment_syntax"], parent_nodes=list() if "parents" not in question.keys() else [q["attributes"]["objectName"] for q in question["parents"]])

        if len(follow_up_questions):
            for question_id in tqdm(follow_up_questions, desc="Convert the mdd/ddf file for the follow-up question"):
                a = ""

        mdd_source.runDMS()

    ###################################################################################

    current_mdd_file = "data\\{}_EXPORT.mdd".format(config["project_name"])
    current_ddf_file = "data\\{}_EXPORT.ddf".format(config["project_name"])

    if not os.path.isfile(current_mdd_file) or not os.path.isfile(current_ddf_file):
        raise Exception("File Error", "File mdd/ddf is not exist.")

    #Open the update data file
    df_update_data = pd.read_csv("source\\update_data.csv", encoding="utf-8")

    if not df_update_data.empty:
        df_update_data.set_index(["InstanceID"], inplace=True)

    adoConn = w32.Dispatch('ADODB.Connection')
    conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(current_ddf_file, current_mdd_file)
    
    #Delete all data before inserting new data (default is FALSE)
    if config["source_initialization"]["delete_all"]:
        adoConn.Open(conn)
        sql_delete = "DELETE FROM VDATA"
        adoConn.Execute(sql_delete)
        adoConn.Close()

    for proto_id, xml_file in config["main"]["xmls"].items():
        for csv_file in tqdm(isurveys[int(proto_id)]["csv_files"], desc="During the data insertion process"):
            df = pd.read_csv(csv_file, encoding="utf-8", low_memory=False)
            df.set_index(['InstanceID'], inplace=True)

            #Allow inserting dummy data
            if config["source_initialization"]["dummy_data_required"]:
                df = df.loc[df["System_LocationID"] == "_DefaultSP"]
            else:
                df = df.loc[df["System_LocationID"] != "_DefaultSP"]
                
            m = Metadata(mdd_file=current_mdd_file, ddf_file=current_ddf_file, sql_query="SELECT InstanceID FROM VDATA")
            df_instanceids = m.convertToDataFrame(questions=["InstanceID"])
            
            if not df_instanceids.empty:
                df_instanceids.set_index(['InstanceID'], inplace=True)
            
            ids = [id for id in list(df.index) if str(id) not in list(df_instanceids.index)]

            if len(ids) > 0:
                df_data = df.loc[ids]
                
                #Allow inserting dummy data
                if config["source_initialization"]["dummy_data_required"]:
                    df_data = df_data.loc[df_data["System_LocationID"] == "_DefaultSP"]
                else:
                    df_data = df_data.loc[df_data["System_LocationID"] != "_DefaultSP"]
                
                if not df_data.empty:
                    adoConn.Open(conn)

                    for i, row in tqdm(df_data[list(df_data.columns)].iterrows(), desc="Insert {} instanceids into the mdd/ddf file".format(len(ids))):
                        try:
                            isurvey = isurveys[int(row["ProtoSurveyID"])]["survey"]
                        except Exception as ex:
                            raise Exception("Config Error", "ProtoID {} should be declare in the config file.".format(str(row["ProtoSurveyID"])))
                            
                        sql_insert = "INSERT INTO VDATA(InstanceID) VALUES(%s)" % (row.name)
                        adoConn.Execute(sql_insert)
                        
                        c = list()
                        v = list()

                        for key, question in isurvey["questions"].items():
                            if question["attributes"]["objectName"] in ["_Q22a"]:
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
                                                            v.append("'{}'".format(re.sub(pattern="\r|\n", repl=",", string=str(row[csv_other_obj["csv"][0]]))))
                                                        case 4:
                                                            v.append("'{}'".format(row[csv_other_obj["csv"][0]]))
                                        else:
                                            if not pd.isnull(row[csv_obj["csv"][0]]):
                                                match question["datatype"].value:
                                                    case dataTypeConstants.mtText.value:
                                                        c.append(mdd_col)
                                                        v.append("'{}'".format(". ".join(re.split('\n|\r|\'|\"', str(row[csv_obj["csv"][0]])))))
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
                    #End insert data csv to mdd
    
    #Delete all data before inserting new data (default is FALSE)
    if config["source_initialization"]["remove_all_ids"]:
        adoConn.Open(conn)

        sql_delete = "DELETE FROM VDATA WHERE Not _LoaiPhieu.ContainsAny({_1,_5})"
        adoConn.Execute(sql_delete)
        adoConn.Close()

    col_to_export = ["InstanceID","_ResName","_ResAddress","_ResHouseNo","_ResStreet","_ResProvinces","_ResDistrictSelected","_ResWardSelected","_ResPhone","_ResCellPhone","_Email","_IntID","_IntName","_LoaiPhieu"]
    col_to_remove = ["_ResName","_ResAddress","_ResHouseNo","_ResStreet","_ResPhone","_ResCellPhone","_Email","_IntID","_IntName","SHELL_NAME","SHELL_BLOCK_TEL","SHELL_TEL","SHELL_BLOCK_ADDRESS","SHELL_ADDRESS"]

    respondent_file = './data/{}_Respondent.xlsx'.format(config["project_name"])

    if not os.path.isfile(respondent_file) or config["source_initialization"]["delete_all"]:
        df_respondent_info = pd.DataFrame(columns=col_to_export)
    else:
        if os.path.getsize(respondent_file) == 0:
            df_respondent_info = pd.DataFrame(columns=col_to_export)
        else:
            df_respondent_info = pd.read_excel(respondent_file)
            os.remove(respondent_file)

    writer = pd.ExcelWriter(respondent_file, engine='xlsxwriter',mode='w')

    m = Metadata(mdd_file=current_mdd_file, ddf_file=current_ddf_file,sql_query="SELECT * FROM VDATA")
    df_respondent_info = m.convertToDataFrame(questions=col_to_export)

    df_respondent_info.reset_index()
    df_respondent_info.to_excel(writer,sheet_name="Respondent",index=False)
    writer.close()

    a = "{}".format(current_mdd_file.replace("_EXPORT","_CLEAN_EXPORT"))
    b = "{}".format(current_ddf_file.replace("_EXPORT","_CLEAN_EXPORT"))

    if os.path.isfile(a):
        os.remove(a)
    if os.path.isfile(b):
        os.remove(b)

    #shutil.copy(current_mdd_file, a)
    mdd_source = Metadata(mdd_file=current_mdd_file.replace("_EXPORT","_CLEAN"), dms_file=r"dms\OutputDDFFile.dms", default_language=config["source_initialization"]["default_language"])
    mdd_source.runDMS()
    
    os.remove(b)
    shutil.copy(current_ddf_file, b)
    m = Metadata(mdd_file=a, ddf_file=b,sql_query="SELECT * FROM VDATA")
    m.delVariables(questions=col_to_remove)
    
except Exception as error:
    print(repr(error))
    #sys.exit(repr(error))