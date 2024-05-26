import os
import shutil
import pandas as pd
import numpy as np
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

    #parent_nodes = list()

    for question_name, question in isurveys[main_protoid_final]["questions"].items():
        if "syntax" in question.keys():
            #if pd.isnull(question["parent"]):
            #    parent_nodes = list()
            #else:
            #    if question["parent"] not in parent_nodes:
            #        parent_nodes.append(question["parent"])

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
            if question["datatype"] not in [dataTypeConstants.mtNone, dataTypeConstants.mtLevel]:
                for i in range(len(row[question["columns"]["mdd"]])):
                    col_mdd = question["columns"]["mdd"][i]
                    col_csv = question["columns"]["csv"][i] 

                    if not pd.isnull(row[col_csv]):
                        match question["datatype"].value:
                            case dataTypeConstants.mtText.value:
                                c.append(col_mdd)
                                v.append("'{}'".format(row[col_csv]))
                            case dataTypeConstants.mtDate.value:
                                c.append(col_mdd)
                                v.append("'{}'".format(row[col_csv]))
                            case dataTypeConstants.mtDouble.value:
                                c.append(col_mdd)
                                v.append("{}".format(row[col_csv]))
                                    
                        
        
        sql_update = "UPDATE VDATA SET " + ','.join([cx + str(r" = %s") for cx in c]) % tuple(v) + " WHERE record = {}".format(row.name)
        adoConn.Execute(sql_update)    
    except Exception as ex:
        #print(sql_insert, ex, sep="-")
        sys.exit(1)

a = ""



for i, row in df_datasource[list(df_datasource.columns)].iterrows():
    try:
        start = isurvey["variables"]['record']['position']['start']
        length = isurvey["variables"]['record']['position']['finish']
        record_id = re.sub(pattern="\s", repl="", string=row[0][start:length])

        sql_insert = "INSERT INTO VDATA(record) VALUES(%s)" % (record_id)
        adoConn.Execute(sql_insert)
        
        c = list()
        v = list()

        for variable_name, variable in isurvey["variables"].items():
            if variable_name not in ['date','record']:
                start = isurvey["variables"][variable_name]['position']['start']
                length = isurvey["variables"][variable_name]['position']['finish']
                
                if len(re.sub(pattern="\s", repl="", string=row[0][start:length])) > 0:
                    if variable_name == "Q1":
                        s = ""
                    
                    c.append(variable_name)

                    match variable['type']:
                        case 'quantity':
                            value = re.sub(pattern="\s", repl="", string=row[0][start:length])
                            v.append(value)
                        case 'character':
                            value = re.sub(pattern="\s", repl="", string=row[0][start:length])
                            v.append("'{}'".format(value))
                        case 'single':
                            value = re.sub(pattern="\s", repl="", string=row[0][start:length])
                            v.append("{_%s}" % (value))

                            for code, helperfield in variable['helperfields'].items():
                                if int(record_id) in list(df_oe_datasource.index):
                                    if pd.notnull(df_oe_datasource.loc[int(record_id), helperfield['name']]):
                                        c.append("{}.{}".format(variable_name, helperfield['name']))
                                        v.append("'{}'".format(df_oe_datasource.loc[int(record_id), helperfield['name']]))
                        case 'multiple':
                            value = row[0][start:length]
                            
                            arr = list()

                            for i in range(len(value)):
                                if len(re.sub(pattern="\s", repl="", string=value[i])) > 0:
                                    if int(value[i]) == 1:
                                        arr.append("_{}".format(list(variable['values'].keys())[int(i)]))
                            
                            v.append("{%s}" % (",".join(arr)))

                            for code, helperfield in variable['helperfields'].items():
                                if int(record_id) in list(df_oe_datasource.index):
                                    if pd.notnull(df_oe_datasource.loc[int(record_id), helperfield['name']]):
                                        c.append("{}.{}".format(variable_name, helperfield['name']))
                                        v.append("'{}'".format(df_oe_datasource.loc[int(record_id), helperfield['name']]))
        
        """
        idx = 0

        for s in tuple(row):
            if not pd.isna(s):
                c.append(row.index[idx])

                if df_datasource[row.index[idx]].dtype.name in ["object","int"]:
                    s = s.replace("\n", "")
                    v.append("'{}'".format(s) if len(s) > 0 else "NULL")
                else:
                    v.append(s) 
            idx += 1    
        """

        sql_update = "UPDATE VDATA SET " + ','.join([cx + str(r" = %s") for cx in c]) % tuple(v) + " WHERE record = {}".format(record_id)
        adoConn.Execute(sql_update)    
    except Exception as ex:
        print(sql_insert, ex, sep="-")
        sys.exit(1)















df_datasource = pd.read_csv(r'Syncopa project\230951.dat', delimiter='\t', header=None)
df_oe_datasource = pd.read_excel(r'Syncopa project\230951.xlsx', engine="openpyxl")
df_oe_datasource.set_index(['record'], inplace=True)

adoConn = w32.Dispatch('ADODB.Connection')
conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(mdd_source.mdd_file.replace('.mdd', '_EXPORT.ddf'), mdd_source.mdd_file.replace('.mdd', '_EXPORT.mdd'))
adoConn.Open(conn)

sql_delete = "DELETE FROM VDATA"
adoConn.Execute(sql_delete)

for i, row in df_datasource[list(df_datasource.columns)].iterrows():
    try:
        start = isurvey["variables"]['record']['position']['start']
        length = isurvey["variables"]['record']['position']['finish']
        record_id = re.sub(pattern="\s", repl="", string=row[0][start:length])

        sql_insert = "INSERT INTO VDATA(record) VALUES(%s)" % (record_id)
        adoConn.Execute(sql_insert)
        
        c = list()
        v = list()

        for variable_name, variable in isurvey["variables"].items():
            if variable_name not in ['date','record']:
                start = isurvey["variables"][variable_name]['position']['start']
                length = isurvey["variables"][variable_name]['position']['finish']
                
                if len(re.sub(pattern="\s", repl="", string=row[0][start:length])) > 0:
                    if variable_name == "Q1":
                        s = ""
                    
                    c.append(variable_name)

                    match variable['type']:
                        case 'quantity':
                            value = re.sub(pattern="\s", repl="", string=row[0][start:length])
                            v.append(value)
                        case 'character':
                            value = re.sub(pattern="\s", repl="", string=row[0][start:length])
                            v.append("'{}'".format(value))
                        case 'single':
                            value = re.sub(pattern="\s", repl="", string=row[0][start:length])
                            v.append("{_%s}" % (value))

                            for code, helperfield in variable['helperfields'].items():
                                if int(record_id) in list(df_oe_datasource.index):
                                    if pd.notnull(df_oe_datasource.loc[int(record_id), helperfield['name']]):
                                        c.append("{}.{}".format(variable_name, helperfield['name']))
                                        v.append("'{}'".format(df_oe_datasource.loc[int(record_id), helperfield['name']]))
                        case 'multiple':
                            value = row[0][start:length]
                            
                            arr = list()

                            for i in range(len(value)):
                                if len(re.sub(pattern="\s", repl="", string=value[i])) > 0:
                                    if int(value[i]) == 1:
                                        arr.append("_{}".format(list(variable['values'].keys())[int(i)]))
                            
                            v.append("{%s}" % (",".join(arr)))

                            for code, helperfield in variable['helperfields'].items():
                                if int(record_id) in list(df_oe_datasource.index):
                                    if pd.notnull(df_oe_datasource.loc[int(record_id), helperfield['name']]):
                                        c.append("{}.{}".format(variable_name, helperfield['name']))
                                        v.append("'{}'".format(df_oe_datasource.loc[int(record_id), helperfield['name']]))
        
        """
        idx = 0

        for s in tuple(row):
            if not pd.isna(s):
                c.append(row.index[idx])

                if df_datasource[row.index[idx]].dtype.name in ["object","int"]:
                    s = s.replace("\n", "")
                    v.append("'{}'".format(s) if len(s) > 0 else "NULL")
                else:
                    v.append(s) 
            idx += 1    
        """

        sql_update = "UPDATE VDATA SET " + ','.join([cx + str(r" = %s") for cx in c]) % tuple(v) + " WHERE record = {}".format(record_id)
        adoConn.Execute(sql_update)    
    except Exception as ex:
        print(sql_insert, ex, sep="-")
        sys.exit(1)

"""
for objectname, question in isurvey["questions"].items():
    if "_LST" in objectname:
        a = ""
    if "syntax" in list(question.keys()):   
        mdd_source.addScript(objectname, question["syntax"], is_defined_list=question["attributes"]["surveyBuilderV3CMSObjGUID"] == "F620C65C-1072-4CF0-B293-A9C9012F5BE8")

df_survey_structure = pd.read_excel(excel_path, sheet_name='Survey Structure', index_col=[0,1])
df_datasource = pd.read_excel(excel_path, sheet_name='Sheet1')

if not os.path.exists(current_mdd_file):
    shutil.copy(source_mdd_file, current_mdd_file)

mdd_source = Metadata(mdd_file=current_mdd_file, dms_file=source_dms_file)

questions = {}

clean = re.compile('<.*?>')
others = ["ghi rõ","please specify","khác","làm rõ"]

main_columns = [c for c in list(df_datasource.columns) if re.match(pattern="((?=(^(\[COMMENT\])))|(?!(^(\[COMMENT\]))))(.*)(({([0-9]*)})$)", string=c, flags=re.S) is None]

for c in main_columns:
    if re.match(pattern="^(Unnamed:|Date|Time)(.*)", string=c):
        df_datasource.drop(columns=c, inplace=True)
    else:
        question_name = re.sub(pattern="^(Form:)(.*)\/", repl="", string=c)
        question_name = question_name.strip().replace(" ", "_")

        if question_name not in list(questions.keys()):
            questions[question_name] = dict()

        questions[question_name]["question_name"] = question_name
        questions[question_name]["question_text"] = c

        df_datasource.rename(columns={ c : questions[question_name]["question_name"] }, inplace=True)

        mdd_source.addScript(question_name, '%s "%s" %s;\n\n' % (questions[question_name]["question_name"], questions[question_name]["question_text"], "text" if df_datasource[questions[question_name]["question_name"]].dtype.name in ["object","str"] else "double"))

for i, row in df_survey_structure[list(df_survey_structure.columns)].iterrows():
    if str(i[0]) not in list(questions.keys()):
        questions[str(i[0])] = dict()
    
    m = re.match(pattern="^(\w*)\.", string=i[1])
    question_name = "Q{0}".format(i[0]) if m is None else  i[1][m.span()[0]:m.span()[1] - 1]
    
    m = re.match(pattern="^(.*)", string=i[1])
    question_text = i[1][m.span()[0] : m.span()[1]]
    
    questions[str(i[0])]["question_name"] = question_name
    questions[str(i[0])]["question_text"] = question_text

    if "categories" not in questions[str(i[0])].keys():
        questions[str(i[0])]["categories"] = dict()
    
    category_name = "_{0}".format(row["Answer Position"]) 
    category_label = "No label" if pd.isnull(row["Answer Text"]) else re.sub(clean, '', row["Answer Text"].replace('"', '\''))

    if category_name not in questions[str(i[0])]["categories"].keys():
        if category_name not in questions[str(i[0])]["categories"].keys():
            questions[str(i[0])]["categories"][category_name] = dict()
        
        questions[str(i[0])]["categories"][category_name]["label"] = category_label

        if not np.isnan(row['Answer Score']):
            questions[str(i[0])]["categories"][category_name]["factor"] = row['Answer Score']
        
        questions[str(i[0])]["categories"][category_name]["isother"] = False

        if(re.match(pattern="(.*[\(\[]*)({})([\)\]]*.*:*)".format("|".join(others)), string=category_label, flags=re.I)):
            questions[str(i[0])]["categories"][category_name]["isother"] = True

for qid, question in questions.items():
    if re.match(pattern="^\{(\d+)\}$", string="{%s}" % (qid)):
        main_cols = [ col for col in list(df_datasource.columns) if "[COMMENT]" not in col and re.match(pattern="\{%s\}" % (qid), string=re.sub(pattern="(.*)(?=({([0-9]*)})$)", repl="", string=col, flags=re.S)) ]

        if len(main_cols) == 1:
            df_datasource.rename(columns={ main_cols[0] : question["question_name"] }, inplace=True)

            if question["question_name"] == "Q0b":
                a = ""

            if "float" in df_datasource[question["question_name"]].dtype.name:
                df_datasource[question["question_name"]] = df_datasource[question["question_name"]].fillna(0).astype(np.int64)

                df_datasource.loc[df_datasource[question["question_name"]].notnull(), question["question_name"]] = "{_" + df_datasource[question["question_name"]].astype(str) + "}"

                df_datasource[question["question_name"]] = df_datasource[question["question_name"]].replace(to_replace="{_0}", value=np.nan)
            else:
                df_datasource.loc[df_datasource[question["question_name"]].notnull(), question["question_name"]] = "{_" + df_datasource[question["question_name"]].astype(str) + "}"
                df_datasource[question["question_name"]] = df_datasource[question["question_name"]].replace(regex=";", value=",_")

        question_syntax = '%s "%s"\ncategorical%s\n{\n' % (question["question_name"], question["question_text"], "[1..]" if df_datasource[question["question_name"]].dtype.name in ['object','str'] else "[1..1]")

        for cid, category in question["categories"].items():
            category_syntax = '\t%s "%s"' % (cid, category["label"])
            
            if "factor" in category.keys():
                category_syntax += " factor(%s)" % category["factor"]

            category_syntax += " other" if category["isother"] is True else ""
            category_syntax += ",\n" if list(question["categories"].keys()).index(cid) < len(list(question["categories"].keys())) - 1 else "\n"

            question_syntax += category_syntax 

        question_syntax += '};\n\n'

        mdd_source.addScript(question["question_name"], question_syntax)
        
        #COMMENT
        comment_cols = [ col for col in list(df_datasource.columns) if "[COMMENT]" in col and re.match(pattern="\{%s\}" % (qid), string=re.sub(pattern="(.*)(?=({([0-9]*)})$)", repl="", string=col, flags=re.S)) ]

        if len(comment_cols) == 1:
            df_datasource.rename(columns={ comment_cols[0] : "{}_Text".format(question["question_name"]) }, inplace=True)

            question_syntax = '%s "%s" text;\n\n' % ("{}_Text".format(question["question_name"]), comment_cols[0]) 

            mdd_source.addScript("{}_Text".format(question["question_name"]), question_syntax)
        

other_columns = [c for c in list(df_datasource.columns) if re.match(pattern="^(\[COMMENT\])(.*)(({([0-9]*)})$)", string=c, flags=re.S)] 

for c in other_columns:
    question_name = re.sub(pattern="^(\[COMMENT\])", repl="", string=c)
    question_name = question_name.strip()

    m = re.match(pattern="^(\w*)\.", string=question_name)
    question_name = "Q{0}".format(i[0]) if m is None else  question_name[m.span()[0]:m.span()[1] - 1]
    
    if question_name not in list(questions.keys()):
        questions[question_name] = dict()

    questions[question_name]["question_name"] = question_name
    questions[question_name]["question_text"] = c

    mdd_source.addScript(question_name, '%s "%s" text;\n\n' % (questions[question_name]["question_name"], questions[question_name]["question_text"]))

    df_datasource.rename(columns={ c : questions[question_name]["question_name"] }, inplace=True)
        

mdd_source.runDMS()


df_datasource.set_index(["Survey_ID","Shopper"], inplace=True)

adoConn = w32.Dispatch('ADODB.Connection')
conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(mdd_source.mdd_file.replace('.mdd', '_EXPORT.ddf'), mdd_source.mdd_file.replace('.mdd', '_EXPORT.mdd'))
adoConn.Open(conn)

sql_delete = "DELETE FROM VDATA"
adoConn.Execute(sql_delete)

for i, row in df_datasource[list(df_datasource.columns)].iterrows():
    try:
        sql_insert = "INSERT INTO VDATA(Survey_ID,Shopper) VALUES('%s','%s')" % (i[0], i[1])
        adoConn.Execute(sql_insert)
        
        c = list()
        v = list()

        idx = 0

        for s in tuple(row):
            if not pd.isna(s):
                c.append(row.index[idx])

                if df_datasource[row.index[idx]].dtype.name in ["object","int"]:
                    s = s.replace("\n", "")
                    v.append("'{}'".format(s) if len(s) > 0 else "NULL")
                else:
                    v.append(s) 
            idx += 1    
        
        sql_update = "UPDATE VDATA SET " + ','.join([cx + str(r" = %s") for cx in c]) % tuple(v) + " WHERE Survey_ID = {}".format(int(i[0]))
        adoConn.Execute(sql_update)    
    except Exception as ex:
        print(sql_insert, ex, sep="-")
        sys.exit(1)

"""



