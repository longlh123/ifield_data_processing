from asyncio import exceptions
from datetime import datetime
from selectors import SelectorKey
from typing import Iterator
from xml.dom.pulldom import ErrorHandler
import win32com.client as w32
import pandas as pd
import re
import numpy as np
from object.enumerations import objectTypeConstants, dataTypeConstants

import collections.abc
#hyper needs the four following aliases to be done manually.
collections.Iterable = collections.abc.Iterable
collections.Mapping = collections.abc.Mapping
collections.MutableSet = collections.abc.MutableSet
collections.MutableMapping = collections.abc.MutableMapping

import savReaderWriter

class mrDataFileDsc:
    def __init__(self, **kwargs):
        if "mdd_file" in list(kwargs.keys()):
            self.mdd_file = kwargs.get("mdd_file")
        if "ddf_file" in list(kwargs.keys()):
            self.ddf_file = kwargs.get("ddf_file")
        if "sql_query" in list(kwargs.keys()):
            self.sql_query = kwargs.get("sql_query")
        if "dms_file" in list(kwargs.keys()):
            self.dms_file = kwargs.get("dms_file")
        
        self.MDM = w32.Dispatch(r'MDM.Document')
        self.adoConn = w32.Dispatch(r'ADODB.Connection')
        self.adoRS = w32.Dispatch(r'ADODB.Recordset')
        self.DMOMJob = w32.Dispatch(r'DMOM.Job')
        self.Directives = w32.Dispatch(r'DMOM.StringCollection')

    def openMDM(self):
        self.MDM.Open(self.mdd_file)

    def saveMDM(self):
        self.MDM.Save(self.mdd_file)

    def closeMDM(self):
        self.MDM.Close()

    def openDataSource(self):
        conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(self.ddf_file, self.mdd_file)

        self.adoConn.Open(conn)

        self.adoRS.ActiveConnection = conn
        self.adoRS.Open(self.sql_query)
        
    def closeDataSource(self):
        #Close and clean up
        if self.adoRS.State == 1:
            self.adoRS.Close()
            self.adoRS = None
        if self.adoConn is not None:
            self.adoConn.Close()
            self.adoConn = None

    def runDMS(self):
        self.Directives.Clear()
        self.Directives.add('#define InputDataFile ".\{}"'.format(self.mdd_file))

        self.Directives.add('#define OutputDataMDD ".\{}"'.format(self.mdd_file.replace('.mdd', '_EXPORT.mdd')))
        self.Directives.add('#define OutputDataDDF ".\{}"'.format(self.mdd_file.replace('.mdd', '_EXPORT.ddf')))

        self.DMOMJob.Load(self.dms_file, self.Directives)
        self.DMOMJob.Run()

class Metadata(mrDataFileDsc):
    def __init__(self, **kwargs):
        try:
            match len(kwargs.keys()):
                case 1:
                    self.mdd_file = kwargs.get("mdd_file")
                    mrDataFileDsc.__init__(self, mdd_file=self.mdd_file)
                case 2:
                    self.mdd_file = kwargs.get("mdd_file")
                    self.dms_file = kwargs.get("dms_file")
                    mrDataFileDsc.__init__(self, mdd_file=self.mdd_file, dms_file=self.dms_file)
                case 3:
                    self.mdd_file = kwargs.get("mdd_file")
                    self.ddf_file = kwargs.get("ddf_file")
                    self.sql_query = kwargs.get("sql_query")
                    mrDataFileDsc.__init__(self, mdd_file=self.mdd_file, ddf_file=self.ddf_file, sql_query=self.sql_query)
                case 4:
                    self.mdd_file = kwargs.get("mdd_file")
                    self.ddf_file = kwargs.get("ddf_file")
                    self.sql_query = kwargs.get("sql_query")
                    self.dms_file = kwargs.get("dms_file")
                    mrDataFileDsc.__init__(self, mdd_file=self.mdd_file, ddf_file=self.ddf_file, sql_query=self.sql_query, dms_file=self.dms_file)
        except ValueError as ex:
            print("Error")
    
    def convertToDataFrame(self, questions):
        self.openMDM()
        self.openDataSource()
        
        d = { 'columns' : list(), 'values' : list() }

        i = 0
        
        while not self.adoRS.EOF:
            r = self.getRows(questions, i)

            d['values'].append(r['values'])
            
            if i == 0: 
                d['columns'].append(r['columns'])

            i += 1
            self.adoRS.MoveNext()

        self.closeMDM()
        self.closeDataSource()
        
        if len(d['values']) > 0:
            return pd.DataFrame(data=d['values'], columns=d['columns'][0])
        else:
            return pd.DataFrame()
        
    def getRows(self, questions, row_index):
        r = {
            'columns' : list(),
            'values' : list()  
        }

        for question in questions:
            q = self.getRow(self.MDM.Fields[question], row_index)

            r['values'].extend(q['values'])
            r['columns'].extend(q['columns'])

        return r

    def getRow(self, field, row_index):
        r = {
            'columns' : list(),
            'values' : list()  
        }

        match str(field.ObjectTypeValue):
            case objectTypeConstants.mtVariable.value:
                q = self.getValue(field)
                        
                r['values'].extend(q['values'])
                r['columns'].extend(q['columns'])
            case objectTypeConstants.mtRoutingItems.value:
                if field.UsageType != 1048:
                    q = self.getValue(field)
                    
                    r['values'].extend(q['values'])
                    r['columns'].extend(q['columns'])
            case objectTypeConstants.mtClass.value: #Block Fields
                for f in field.Fields:
                    if f.Properties["py_isHidden"] is None or f.Properties["py_isHidden"] == False:
                        q = self.getRow(f, row_index)
                        
                        r['values'].extend(q['values'])
                        r['columns'].extend(q['columns'])
            case objectTypeConstants.mtArray.value: #Loop
                a = field.Name

                for variable in field.Variables:
                    if variable.Properties["py_isHidden"] is None or variable.Properties["py_isHidden"] == False:
                        q = self.getRow(variable, row_index)
                        
                        r['values'].extend(q['values'])
                        r['columns'].extend(q['columns'])
        return r

    def getValue(self, question): 
        q = {
            'columns' : list(),
            'values' : list()  
        }
        
        max_range = 0
        
        column_name = question.FullName if str(question.ObjectTypeValue) != objectTypeConstants.mtVariable.value else question.Variables[0].FullName

        if str(question.ObjectTypeValue) == objectTypeConstants.mtRoutingItems.value:
            if question.Properties["py_setColumnName"] is not None:
                s = ""

                for i in range(question.Indices.Count):
                    s = s + question.Indices[i].FullName.replace("_","_R")
                    
                alias_name = "{}{}".format(question.Properties["py_setColumnName"], s)
                #alias_name = "{}{}".format(question.Properties["py_setColumnName"], question.Indices[0].FullName.replace("_","_R"))
            else:
                alias_name = column_name
        else:
            if question.LevelDepth == 1:
                if question.UsageTypeName == "OtherSpecify":
                    alias_name = "{}{}".format(column_name if question.Parent.Properties["py_setColumnName"] is None else question.Parent.Properties["py_setColumnName"], question.Name)
                else:
                    alias_name = column_name if question.Properties["py_setColumnName"] is None else question.Properties["py_setColumnName"]
            elif question.LevelDepth == 2:
                current_index_path = re.sub(pattern="{Recall_|}", repl="", string=question.CurrentIndexPath) 

                if question.UsageTypeName == "OtherSpecify":
                    alias_name = "{}{}".format(column_name if question.Parent.Properties["py_setColumnName"] is None else question.Parent.Properties["py_setColumnName"], question.Name)
                else:
                    alias_name = column_name if question.Properties["py_setColumnName"] is None else question.Properties["py_setColumnName"]

                alias_name = "p{}_{}".format(current_index_path, alias_name)

        if question.DataType == dataTypeConstants.mtCategorical.value:    
            show_helperfields = False if question.Properties["py_showHelperFields"] is False else True

            cats_resp = str(self.adoRS.Fields[column_name].Value)[1:(len(str(self.adoRS.Fields[column_name].Value))-1)].split(",")

            if question.Properties["py_showPunchingData"]:
                for category in question.Categories:
                    if not category.IsOtherLocal:
                        q['columns'].append("{}{}".format(alias_name, category.Name.replace("_", "_C")))
                        
                        if question.Properties["py_showVariableValues"] is None:
                            if category.Name in cats_resp:
                                q['values'].append(1)
                            else:
                                q['values'].append(0 if self.adoRS.Fields[column_name].Value is not None else np.nan)
                        else:
                            if category.Name in cats_resp:
                                q['values'].append(category.Label) 
                            else:
                                q['values'].append(np.nan)
                
                if question.HelperFields.Count > 0:
                    if question.Properties["py_combibeHelperFields"]:
                        q['columns'].append("{}{}".format(alias_name, category.Name.replace(category.Name, "_C97")))
                            
                        str_others = ""

                        for helperfield in question.HelperFields:
                            if helperfield.Name in cats_resp:
                                str_others = str_others + (", " if len(str_others) > 0 else "") + self.adoRS.Fields["{}.{}".format(column_name, helperfield.Name)].Value
                        
                        if len(str_others) > 0:
                            match question.Properties["py_showVariableValues"]:
                                case "Names":
                                    q['values'].append(question.Categories[helperfield.Name].Name.replace('_',''))
                                case "Labels":
                                    q['values'].append(question.Categories[helperfield.Name].Label)
                                case _:
                                    q['values'].append(1)
                        else:
                            q['values'].append(np.nan)
                        
                        if show_helperfields:
                            q['columns'].append("{}{}".format(alias_name, category.Name.replace(category.Name, "_C97_Other")))

                            if len(str_others) > 0:
                                q['values'].append(str_others) 
                            else:
                                q['values'].append(np.nan)
                    else:
                        for helperfield in question.HelperFields:
                            q['columns'].append("{}{}".format(alias_name, helperfield.Name.replace("_", "_C")))
                            
                            if question.Properties["py_showVariableValues"] is None:
                                if helperfield.Name in cats_resp:
                                    q['values'].append(1)
                                else:
                                    q['values'].append(0 if self.adoRS.Fields[column_name].Value is not None else np.nan)
                            else:
                                if helperfield.Name in cats_resp:
                                    q['values'].append(helperfield.Label) 
                                else:
                                    q['values'].append(np.nan)

                            if show_helperfields:
                                q['columns'].append("{}{}_Other".format(alias_name, helperfield.Name.replace("_", "_C")))

                                if helperfield.Name in cats_resp:
                                    q['values'].append(self.adoRS.Fields["{}.{}".format(column_name, helperfield.Name)].Value)
                                else: 
                                    q['values'].append(np.nan)
            elif question.Properties["py_showVariableValues"] == "Names":
                q['columns'].append(alias_name)
                q['values'].append(np.nan if self.adoRS.Fields[column_name].Value == None else self.adoRS.Fields[column_name].Value)
            else:
                max_range = question.MaxValue if question.MaxValue is not None else question.Categories.Count
                
                for i in range(max_range):
                    col_name = alias_name if question.MinValue == 1 and question.MaxValue == 1 else "{}_{}".format(alias_name, i + 1)
                    q['columns'].append(col_name)

                    #Generate a column which contain a factor of a category variable (only for single answer question)
                    if question.MinValue == 1 and question.MaxValue == 1:
                        if question.Properties["py_showVariableFactor"] is not None:
                            col_name = "FactorOf{}".format(alias_name)
                            q['columns'].append(col_name)

                    if i < len(cats_resp):
                        category = cats_resp[i]

                        match question.Properties["py_showVariableValues"]:
                            case "Names":
                                q['values'].append(np.nan if self.adoRS.Fields[column_name].Value == None else question.Categories[category].Name)
                            case "Labels":
                                q['values'].append(np.nan if self.adoRS.Fields[column_name].Value == None else question.Categories[category].Label)
                            case _:
                                if type(category[1:len(category)]) is str:
                                    q['values'].append(np.nan if self.adoRS.Fields[column_name].Value == None else str(category[1:len(category)]))
                                else:
                                    q['values'].append(np.nan if self.adoRS.Fields[column_name].Value == None else int(category[1:len(category)]))
                        
                        #Get factor value of a category variable
                        if question.MinValue == 1 and question.MaxValue == 1:
                            if question.Properties["py_showVariableFactor"] is not None:
                                q['values'].append(np.nan if self.adoRS.Fields[column_name].Value == None else question.Categories[category].Factor)
                    else:
                        q['values'].append(np.nan)

                        #Get factor value of a category variable
                        if question.MinValue == 1 and question.MaxValue == 1:
                            if question.Properties["py_showVariableFactor"] is not None:
                                q['values'].append(np.nan)
                
                if show_helperfields:
                    if question.HelperFields.Count > 0:
                        for helperfield in question.HelperFields:
                            col_name = "{}{}_Other".format(alias_name, helperfield.Name.replace("_", "_C"))
                            q['columns'].append(col_name)

                            if helperfield.Name in cats_resp:
                                q['values'].append(np.nan if self.adoRS.Fields["{}.{}".format(column_name, helperfield.Name)].Value == None else self.adoRS.Fields["{}.{}".format(column_name, helperfield.Name)].Value)
                            else:
                                q['values'].append(np.nan) 

        elif question.DataType == dataTypeConstants.mtDate.value:
            q['columns'].append(alias_name)
            q['values'].append(np.nan if self.adoRS.Fields[column_name].Value is None else datetime.strftime(self.adoRS.Fields[column_name].Value, "%d/%m/%Y"))
        elif question.DataType == dataTypeConstants.mtLong.value or question.DataType == dataTypeConstants.mtDouble.value:
            q['columns'].append(alias_name)
            q['values'].append(self.adoRS.Fields[column_name].Value)
        else:
            q['columns'].append(alias_name)
            q['values'].append('' if self.adoRS.Fields[column_name].Value is None else self.adoRS.Fields[column_name].Value)

        if len(q['columns']) != len(q['values']):
            print("A length mismatch error between 'columns': {} and 'values': {}".format(','.join(q['columns']), ','.join(q['values']))) 

        return q

    def addScript(self, question_name, syntax, is_defined_list=False, childnodes=list(), parent_nodes=list()):
        self.openMDM()
        
        if is_defined_list:
            if self.MDM.Types.Exist(question_name):
                self.MDM.Types.Remove(question_name)

            self.MDM.Types.addScript(syntax)
        else:
            if len(parent_nodes) > 0:
                if len(parent_nodes) == 1:
                    if self.MDM.Fields[parent_nodes[0]].Fields.Exist(question_name):
                        self.MDM.Fields[parent_nodes[0]].Fields.Remove(question_name)
                    
                    self.MDM.Fields[parent_nodes[0]].Fields.addScript(syntax)
                elif len(parent_nodes) == 2:
                    if self.MDM.Fields[parent_nodes[0]].Fields[parent_nodes[1]].Fields.Exist(question_name):
                        self.MDM.Fields[parent_nodes[0]].Fields[parent_nodes[1]].Fields.Remove(question_name)
                    
                    self.MDM.Fields[parent_nodes[0]].Fields[parent_nodes[1]].Fields.addScript(syntax)
            else:
                if self.MDM.Fields.Exist(question_name):
                    self.MDM.Fields.Remove(question_name)

                self.MDM.Fields.addScript(syntax)

                for node in childnodes:
                    self.MDM.Fields[question_name].Fields.addScript(node)

        self.saveMDM()
        self.closeMDM()
    
    def getVariables(self):
        self.openMDM()
        arr = list()
        for v in self.MDM.Variables:
            arr.append(v.FullName)
        self.closeMDM()
        return arr

    def addField(self, field):
        self.openMDM()
        self.MDM.Fields.Add(field)
        self.closeMDM()
    
    def getField(self, name):
        self.openMDM()
        field = self.MDM.Fields[name]
        self.closeMDM()
        return field

class BVCObject(mrDataFileDsc):
    def __init__(self, mdd_path, ddf_path, sql_query):
        super().__init__(mdd_path, ddf_path, sql_query)

        self.addScript()
    
    def addScript(self):
        self.openMDM()

        if self.MDM.Fields.Exist("BVC"):
            self.MDM.Fields.Remove("BVC")

        self.MDM.Fields.addScript("""
            BVC "BVC" loop
            {
                _1 "C2" [ childs = "{_1,_2,_3,_4,_34,_5,_6,_29,_30,_33}" ],
                _7 "Không Độ" [ childs = "{_7,_8,_9,_10}" ],
                _11 "Olong Tea Plus" [ childs = "{_11,_12,_13,_14}" ],
                _15 "Dr. Thanh" [ childs = "{_15,_16,_17,_18}" ],
                _19 "Wonderfarm" [ childs = "{_19,_20,_21,_22}" ],
                _23 "Trà Hạt Chia Fuze Tea+ - NET" [ childs = "{_23,_24,_25}" ],
                _26 "Trà sữa không độ/ Trà sữa Macchiato" [ childs = "{_26}" ],
                _27 "Trà sữa C2" [ childs = "{_27}" ],
                _2000 "Trà lá vối Seventy" [ childs = "{_2000}" ],
                _3000 "Trà sữa ít đường Vinamilk Happy Milktea" [ childs = "{_3000}" ],
                _4000 "Trà Tea Go" [ childs = "{_4000,_4001,_4002,_4003}" ],
                _5000 "Trà mật ong Boncha" [ childs = "{_5000,_5001,_5002}" ],
                _6000 "Trà TH True Tea" [ childs = "{_6000,_6001,_6002}" ],
                _7000 "Trà Jokky" [ childs = "{_7000,_7001,_7002,_7003}" ],
                _8000 "Trà thảo mộc VietFuji" [ childs = "{_8000}" ],
                _9000 "Trà Cozy đóng chai" [ childs = "{_9001,_9002,_9003,_9004}" ]
            }fields
            (
                TUBA "Top of mind and total mentions unaided brand aware"
                categorical[1..1]
                {
                    _0 "No" [value=0],
                    _1 "Yes" [value=1]
                };
                
                TMBA "Top of mind unaided brand aware"
                categorical[1..1]
                {
                    _0 "No" [value=0],
                    _1 "Yes" [value=1]
                };

                AWARE "Awareness"
                categorical[1..1]
                {
                    _0 "No" [value=0],
                    _1 "Yes" [value=1]
                };

                P3MUSAGE "P3M Usage"
                categorical[1..1]
                {
                    _0 "No" [value=0],
                    _1 "Yes" [value=1]
                };

                P1MUSAGE "P1M Usage"
                categorical[1..1]
                {
                    _0 "No" [value=0],
                    _1 "Yes" [value=1]
                };

                BUMO "BUMO"
                categorical[1..1]
                {
                    _0 "No" [value=0],
                    _1 "Yes" [value=1]
                };

                SOW "Claimed Share"
                double;

                CONSIDER "Would consider"
                categorical[1..1]
                {
                    _0 "No" [value=0],
                    _1 "Yes" [value=1]
                };

                PBVC "Brand Performance"
                categorical[1..1]
                {
                    _1 "1" [value=1],
                    _2 "2" [value=1],
                    _3 "3" [value=1],
                    _4 "4" [value=1],
                    _5 "5" [value=1],
                    _6 "6" [value=1],
                    _7 "7" [value=1],
                    _8 "8" [value=1],
                    _9 "9" [value=1],
                    _10 "10" [value=1]
                };

                CLBVC "Closeness"
                categorical[1..1]
                {
                    _1 "1" [value=1],
                    _2 "2" [value=1],
                    _3 "3" [value=1],
                    _4 "4" [value=1],
                    _5 "5" [value=1],
                    _6 "6" [value=1],
                    _7 "7" [value=1],
                    _8 "8" [value=1],
                    _9 "9" [value=1],
                    _10 "10" [value=1]
                };

                BIA "" loop
                {
                    _1 "Popular brand or drank by everyone",
                    _2 "Good packaging",
                    _3 "Reliable or trustworthy brand",
                    _4 "Reasonable price or value for money",
                    _5 "Can drink everyday",
                    _6 "Thirst-quenching",
                    _7 "Good taste",
                    _8 "Modern, active and energetic brand",
                    _10 "A drink made from natural ingredients",
                    _12 "No worry about inner heat",
                    _13 "Made from pure green tea",
                    _14 "Bring joyful feelings",
                    _15 "A healthy drink",
                    _16 "A drink can be shared with friends",
                    _17 "Stress of Fatigue release",
                    _18 "Bring feelings of lightness",
                    _19 "Helps absorb less fat",
                    _20 "Cool down my life",
                    _21 "Bring refreshment",
                    _22 "Contain antioxidant EGCG",
                    _23 "Irresistibly delicious",
                    _24 "Bring a sense of coolness",
                    _25 "Recharge energy",
                    _26 "Purifying the body"
                }
                fields
                (
                    BIA_Codes ""
                    categorical[1..1]
                    {
                        _0 "No" [value=0],
                        _1 "Yes" [value=1]
                    };
                )expand grid;

                ME "" loop
                {
                    _2 "Asked but the store does not sell it",
                    _3 "Not preserved in cool place",
                    _4 "Hard to find in stores - Product of other brands display dorminantly",
                    _5 "Out of stock",
                    _7 "Too expensive compare to other RTDT brands",
                    _9 "Does not have a packaging I need (can-pack) ",
                    _10 "Does not have large size",
                    _11 "Does not have smaller size",
                    _12 "Does not have many flavors to choose",
                    _14 "Not have sale promotion while other brands have",
                    _15 "Other brands sale promotion is more attractive",
                    _17 "Not recommended by family-friend",
                    _18 "Not recommended by sellers",
                    _19 "I see another brands’ Point Of Sale activities is more attractive"
                }
                fields
                (
                    ME_Codes ""
                    categorical[1..1]
                    {
                        _0 "No" [value=0],
                        _1 "Yes" [value=1]
                    };
                )expand grid;

                BARCON "" loop
                {
                    _1 "Do not like the flavor",
                    _2 "Unhealthy",
                    _3 "Do not trust in product quality",
                    _4 "Expensive than I expected ",
                    _5 "Too cheap",
                    _6 "Does not have the flavors that I like",
                    _7 "Does not have a packaging that I need",
                    _8 "Has higher sugar than I needed",
                    _9 "Do not have sale promotion while other brands have",
                    _10 " Do not want other people see when I drink/buy",
                    _11 " Unpopular brands",
                    _12 " Do not like recent packaging designs",
                    _13 " Does not elegant, premium ",
                    _14 " Asked but the store does not sell it",
                    _15 " Not preserved in cool place"
                }
                fields
                (
                    BARCON_Codes ""
                    categorical[1..1]
                    {
                        _0 "No" [value=0],
                        _1 "Yes" [value=1]
                    };
                )expand grid;
            )expand grid;
            """)

        self.saveMDM()
        self.closeMDM()

class SPSSObject_Dataframe(mrDataFileDsc):
    def __init__(self, mdd_path, ddf_path, sql_query, questions, group_name=None):
        mrDataFileDsc.__init__(self, mdd_file=mdd_path, ddf_file=ddf_path, sql_query=sql_query)

        self.df = pd.DataFrame()
        self.df_spss = pd.DataFrame()

        self.records = list()
        self.varNames = list()
        self.varMissingValues = dict() 
        self.varLabels = dict()
        self.varTypes = dict()
        
        self.valueLabels = dict()
        self.measureLevels = dict()
        self.var_date_formats = dict()
        self.var_dates = list()

        #Group of columns with format A_suffix1, A_suffix2...B_suffix1, B_suffix2
        self.group_name = group_name
        #Columns to use as id variable
        self.id_vars = list()
        self.value_groupnames = list()
        self.groupname_labels = list()

        self.init(questions)

    def init(self, questions):
        self.openMDM()
        self.openDataSource()
        
        if self.group_name is not None:
            self.value_groupnames.extend(self.getValueLabels(self.MDM.Fields[self.group_name]).keys())

        data = self.adoRS.GetRows()

        columns = [f.Name for f in self.adoRS.Fields]
        
        self.df = pd.DataFrame(data=np.array(data).T, columns=columns)
        self.transform(questions)

        self.closeMDM()
        self.closeDataSource()

    def transform(self, questions):
        
        for question in questions:
            field = self.MDM.Fields[question]

            if field.Properties["py_isHidden"] is None or field.Properties["py_isHidden"] == False:
                self.generate_variable(field)

    def get_iterations(self, field, indexes):
        indexes = indexes.split(',')
        iterations = list()

        def findIteration(f, index):
            if index < 0:
                return
            
            iterations.append(f.Categories[re.sub(pattern="[\{\}]", repl="", string=indexes[index])])
            findIteration(f.Parent.Parent, index - 1)

        findIteration(field, len(indexes) - 1)
        return iterations
        
    def generate_variable(self, field, iterations_list=list(), iterations=list()):

        if str(field.ObjectTypeValue) == objectTypeConstants.mtVariable.value:
            if field.LevelDepth > 1:
                if field.Parent.Parent.Name == "@class":
                    iterations = self.get_iterations(field.Parent.Parent.Parent, field.CurrentIndexPath) 
                else:
                    iterations = self.get_iterations(field.Parent.Parent, field.CurrentIndexPath) 

            if field.DataType in [ dataTypeConstants.mtLong.value, dataTypeConstants.mtDouble.value, dataTypeConstants.mtText.value, dataTypeConstants.mtBoolean.value ]:
                self.transform_simple_data(field, iterations=iterations)
            elif field.DataType == dataTypeConstants.mtCategorical.value:
                self.transform_categorical_data(field, iterations=iterations)
        else:
            iterations_list = list()

            if field.Properties["py_setCategoriesList"]:
                iterations_list = field.Properties["py_setCategoriesList"].split(',')
            else:
                iterations_list = [cat.Name for cat in field.Categories]

            for f in field.Fields:
                if f.Properties["py_isHidden"] is None or f.Properties["py_isHidden"] == False:
                    full_name = f.FullName.replace("..","%s") % tuple(field.CurrentIndexPath.split(','))
                    for iteration in iterations_list:
                        iteration_full_name = full_name.replace("..", "{%s}" % iteration)
                        self.generate_variable(self.MDM.Fields[iteration_full_name], iterations=field.Categories[iteration])
                    
    def transform_simple_data(self, field, iterations=list()):
        var_name = self.get_variable_name(field, iterations=iterations)
        var_label = self.replaceLabel(self.get_variable_label(field, iterations=iterations))

        self.varNames.append(var_name)
        self.varLabels[var_name] = var_label.encode('utf-8')
        self.varTypes[var_name] = self.get_datatype(field)
        self.measureLevels[var_name] = self.get_measure_level(field)

        if field.LevelDepth == 1:
            self.df_spss[var_name] = self.df[field.FullName]
        else:
            self.df_spss[var_name] = self.df[field.FullName.replace("..", field.CurrentIndexPath)]
    
    def transform_categorical_data(self, field, iterations=list()):
        var_name = self.get_variable_name(field, iterations=iterations)
        var_label = self.replaceLabel(self.get_variable_label(field, iterations=iterations))

        value_labels = self.get_categories_list(field, show_punching_data=field.Properties["py_showPunchingData"])
        
        categories_list = list()
        remaining_categories_list = list()
        other_categories_list = list()

        if field.Properties["py_setCategoriesList"]:
            categories_list = field.Properties["py_setCategoriesList"].split(',')
        else:
            categories_list = [cat.Name for cat in field.Categories if not cat.IsOtherLocal]

        remaining_categories_list = [cat.Name for cat in field.Categories if cat.Name not in categories_list and not cat.IsOtherLocal]
        
        try:
            if field.OtherCategories.Count > 0:
                other_categories_list = [cat.Name for cat in field.OtherCategories]
        except AttributeError as e:
            pass

        if field.Properties["py_showPunchingData"]:
            var_fullname = field.FullName if len(iterations) == 0 else field.FullName.replace("..", "%s") % tuple(field.CurrentIndexPath.split(","))

            df_updated = self.df[var_fullname].replace(to_replace="[\{\}]", value="", regex=True)

            for cat_name in categories_list:
                category = field.Categories[cat_name]
                
                attr_name = re.search(pattern="_\d+$", string=var_name)[0]
                q_name = var_name.replace(attr_name, "")

                var_name_temp = "{}_C{}{}".format(q_name, re.sub(pattern="^_", repl="", string=category.Name), attr_name)

                #var_name_temp = "{}_C{}".format(var_name, re.sub(pattern="^_", repl="", string=category.Name))

                self.varNames.append(var_name_temp)
                self.varLabels[var_name_temp] = "{}_{}".format(var_label, category.Label).encode('utf-8')     
                self.varTypes[var_name_temp] = self.get_datatype(field)
                self.measureLevels[var_name_temp] = self.get_measure_level(field)
                self.valueLabels[var_name_temp] = value_labels
                
                self.df_spss = pd.concat([self.df_spss, df_updated], ignore_index=False, axis=1)
                
                match field.Properties["py_setVariableValues"]:
                    case "Values":
                        self.df_spss[var_name_temp] = self.df_spss[df_updated.name].apply(lambda x: None if x is None else 1 if category.Name.lower() in x.split(',') else 0)
                    case "Labels":
                        self.df_spss[var_name_temp] = self.df_spss[df_updated.name].apply(lambda x: None if x is None else "Yes" if category.Name.lower() in x.split(',') else "No")

                self.df_spss.drop(columns=[df_updated.name], inplace=True)

            if len(remaining_categories_list) > 0:
                var_name_temp = "{}_C{}".format(var_name, re.sub(pattern="^_", repl="", string='_97'))

                self.varNames.append(var_name_temp)
                self.varLabels[var_name_temp] = "{}_{}".format(var_label, category.Label).encode('utf-8')     
                self.varTypes[var_name_temp] = self.get_datatype(field)
                self.measureLevels[var_name_temp] = self.get_measure_level(field)
                self.valueLabels[var_name_temp] = value_labels

                self.df_spss = pd.concat([self.df_spss, df_updated], ignore_index=False, axis=1)
                
                match field.Properties["py_setVariableValues"]:
                    case "Values":
                        self.df_spss[var_name_temp] = self.df_spss[df_updated.name].apply(lambda x: None if x is None else 1 if any([y in remaining_categories_list for y in x.split(',')]) else 0)
                    case "Labels":
                        self.df_spss[var_name_temp] = self.df_spss[df_updated.name].apply(lambda x: None if x is None else "Yes" if any([y in remaining_categories_list for y in x.split(',')]) else "No")

                self.df_spss.drop(columns=[df_updated.name], inplace=True)

            if field.Properties["py_showHelperFields"]:
                if field.HelperFields.Count > 0:
                    df_others = pd.DataFrame()
                    df_others["Other"] = self.df[[helperfield.FullName for helperfield in field.HelperFields]].fillna('').sum(1).replace('', np.nan)
                    
                    if field.Properties["py_combibeHelperFields"]:
                        var_name_temp = "{}_C{}".format(var_name, re.sub(pattern="^_", repl="", string='_997'))
                        
                        self.varNames.append(var_name_temp)
                        self.varLabels[var_name_temp] = "{}_{}".format(var_label, category.Label).encode('utf-8')     
                        self.varTypes[var_name_temp] = self.get_datatype(field)
                        self.measureLevels[var_name_temp] = self.get_measure_level(field)
                        self.valueLabels[var_name_temp] = value_labels

                        self.df_spss[var_name_temp] = df_others["Other"].apply(lambda x: 0 if pd.isna(x) else 1)

                        var_name_temp = "{}_{}_Verbatim".format(var_name, re.sub(pattern="^_", repl="", string='_997'))
                        self.varNames.append(var_name_temp)
                        self.varLabels[var_name_temp] = var_label.encode('utf-8')
                        self.varTypes[var_name_temp] = self.get_datatype(field.HelperFields[0])

                        if self.varTypes[var_name_temp] == 0:
                            self.varMissingValues[var_name_temp] = {"values": [999, -1, -2]}

                        self.measureLevels[var_name_temp] = self.get_measure_level(field.HelperFields[0])

                        self.df_spss[var_name_temp] = df_others["Other"].apply(lambda x: '' if pd.isna(x) else x)
                    else:
                        for helperfield in field.HelperFields:
                            var_name_temp = "{}_C{}".format(var_name, re.sub(pattern="^_", repl="", string=helperfield.Name))
                            self.varNames.append(var_name_temp)
                            self.varLabels[var_name_temp] = "{}_{}".format(var_label, category.Label).encode('utf-8')     
                            self.varTypes[var_name_temp] = self.get_datatype(field)
                            self.measureLevels[var_name_temp] = self.get_measure_level(field)
                            self.valueLabels[var_name_temp] = value_labels

                            self.df_spss[var_name_temp] = self.df[helperfield.FullName].apply(lambda x: 0 if pd.isna(x) else 1)

                            var_name_temp = "{}_C{}_Verbatim".format(var_name, re.sub(pattern="^_", repl="", string=helperfield.Name))
                            self.varNames.append(var_name_temp)
                            self.varLabels[var_name_temp] = var_label.encode('utf-8')
                            self.varTypes[var_name_temp] = self.get_datatype(helperfield)
                            self.measureLevels[var_name_temp] = self.get_measure_level(helperfield)

                            self.df_spss[var_name_temp] = self.df[helperfield.FullName].apply(lambda x: '' if pd.isna(x) else x)
        elif field.Properties["py_showVariableNames"] and field.Properties["py_showVariableLabels"]:
            replaced_categories_list = {}
            replaced_values_list = {}

            for cat in field.Categories:
                replaced_categories_list[cat.Name] = cat.Label
                replaced_values_list[cat.Name] = cat.Name if cat.Properties["value"] is None else cat.Properties["value"]

            if field.LevelDepth == 1:
                match field.Properties["py_setVariableValues"]:
                    case "Values":
                        self.df_spss[var_name] = self.df[field.FullName].apply(lambda x: np.nan if x is None else ';'.join(map(replaced_values_list.get, x[1:len(x)-1].split(','))))
                    case _:
                        self.df_spss[var_name] = self.df[field.FullName].replace(to_replace=",_", value=";", regex=True)
                        self.df_spss[var_name] = self.df_spss[var_name].replace(to_replace="[\{\}_]", value="", regex=True)

                self.df_spss[f"{var_name}_Label"] = self.df[field.FullName].apply(lambda x: np.nan if x is None else ';'.join(map(replaced_categories_list.get, x[1:len(x)-1].split(','))))
            else:
                #var_name = "{} - {}".format(var_name, field.CurrentIndexPath)
                var_fullname = field.FullName.replace("..", "%s") % tuple(field.CurrentIndexPath.split(","))
                
                match field.Properties["py_setVariableValues"]:
                    case "Values":
                        self.df_spss[var_name] = self.df[var_fullname].apply(lambda x: np.nan if x is None else ';'.join(map(replaced_values_list.get, x[1:len(x)-1].split(','))))
                    case _:
                        self.df_spss[var_name] = self.df[var_fullname].replace(to_replace=",_", value=";", regex=True)
                        self.df_spss[var_name] = self.df_spss[var_name].replace(to_replace="[\{\}_]", value="", regex=True)
                
                self.df_spss[f"{var_name}_Label"] = self.df[var_fullname].apply(lambda x: np.nan if x is None else ';'.join(map(replaced_categories_list.get, x[1:len(x)-1].split(','))))
        else:
            replaced_categories_list = {}

            for cat in field.Categories:
                cat_name = cat.Name[1:len(cat.Name)] if cat.Name[0:1] in ["_"] else cat.Name

                match field.Properties["py_setVariableValues"]:
                    case "Values":
                        replaced_cat_name = int(cat_name if cat_name.isnumeric() else field.Categories[cat.Name].Properties["value"])
                        replaced_categories_list[cat_name.lower()] = replaced_cat_name
                    case "Labels":
                        replaced_categories_list[cat_name.lower()] = cat.Label
            
            var_fullname = field.FullName if len(iterations) == 0 else field.FullName.replace("..", "%s") % tuple(field.CurrentIndexPath.split(","))

            df_updated = self.df[var_fullname].replace(to_replace="[\{\}]", value="", regex=True)
            df_updated = df_updated.str.split(',', expand=True)
            df_updated = df_updated[list(df_updated.columns)].replace(to_replace="^_", value="", regex=True)

            column_renamed = dict()

            for i in list(df_updated.columns):
                var_name_temp = var_name if field.MinValue == 1 and field.MaxValue == 1 else "{}_{}".format(var_name, i + 1)
                var_name_temp = re.sub(pattern="^_", repl="", string=var_name_temp)
                self.varNames.append(var_name_temp)
                self.varLabels[var_name_temp] = var_label.encode('utf-8')
                self.varTypes[var_name_temp] = self.get_datatype(field)
                self.measureLevels[var_name_temp] = self.get_measure_level(field)
                self.valueLabels[var_name_temp] = value_labels

                column_renamed[i] = var_name_temp
            
            df_updated.rename(columns=column_renamed, inplace=True)
            df_updated.replace(replaced_categories_list, inplace=True)

            self.df_spss = pd.concat([self.df_spss, df_updated], ignore_index=False, axis=1)

            if field.Properties["py_showHelperFields"]:
                if field.HelperFields.Count > 0:
                    df_others = pd.DataFrame()
                    df_others["Other"] = self.df[[helperfield.FullName for helperfield in field.HelperFields]].fillna('').sum(1).replace('', np.nan)

                    if field.Properties["py_combibeHelperFields"]:
                        var_name_temp = "{}_C{}_Verbatim".format(var_name, re.sub(pattern="^_", repl="", string=helperfield.Name))
                        self.varNames.append(var_name_temp)
                        self.varLabels[var_name_temp] = var_label.encode('utf-8')
                        self.varTypes[var_name_temp] = self.get_datatype(helperfield)
                        self.measureLevels[var_name_temp] = self.get_measure_level(helperfield)

                        self.df_spss[var_name_temp] = self.df[helperfield.FullName].apply(lambda x: '' if pd.isna(x) else x)
                    else:
                        for helperfield in field.HelperFields:
                            var_name_temp = "{}_C{}".format(var_name, re.sub(pattern="^_", repl="", string=helperfield.Name))
                            self.varNames.append(var_name_temp)
                            self.varLabels[var_name_temp] = var_label.encode('utf-8')
                            self.varTypes[var_name_temp] = self.get_datatype(helperfield)
                            self.measureLevels[var_name_temp] = self.get_measure_level(helperfield)

                            self.df_spss[var_name_temp] = self.df[helperfield.FullName]

    def get_categories_list(self, field, show_punching_data=False):
        categories_list = dict()

        if show_punching_data:
            categories_list[0] = "No".encode('utf-8')
            categories_list[1] = "Yes".encode('utf-8')            
        else: 
            for cat in field.Categories:
                if not cat.IsOtherLocal:
                    cat_name = cat.Name[1:len(cat.Name)] if cat.Name[0:1] in ["_"] else cat.Name
                    cat_name = int(cat_name if cat_name.isnumeric() else field.Categories[cat.Name].Properties["value"])

                    #cat_name = int(cat.Name[1:len(cat.Name)] if cat.Name[0:1] in ["_"] else cat.Name)
                    categories_list[cat_name] = cat.Label.encode('utf-8')
        
        return categories_list

    def get_measure_level(self, field):
        match field.DataType:
            case dataTypeConstants.mtLong.value | dataTypeConstants.mtDouble.value:
                return "scale"
            case dataTypeConstants.mtText.value:
                return "nominal"
            case _:
                return "nominal"
            
    def get_datatype(self, field):
        match field.DataType:
            case dataTypeConstants.mtLong.value | dataTypeConstants.mtDouble.value:
                return 0
            case dataTypeConstants.mtText.value:
                if field.LevelDepth == 1:
                    return int(self.df[field.FullName].str.encode('utf-8').replace(r'\xcd',r'\xcc\xa6').str.decode('utf-8').str.len().max())
                else:
                    return int(self.df[field.FullName.replace("..", field.CurrentIndexPath)].fillna("").str.len().max())
            case dataTypeConstants.mtDate.value:
                return 0
            case _:
                return 0

    def get_iteration_name(self, iterations=list()):
        labels = list()

        def get_label(index):
            if index < 0:
                return

            labels.append(iterations[index].Name)
            get_label(index - 1) 
        
        get_label(len(iterations) - 1)

        ls = labels[1:len(labels)]
        ls.append(labels[0])
        return "_".join([l if l[0:1] not in ["_"] else l[1:len(l)] for l in ls])
          
    def get_variable_name(self, field, iterations=list()):
        if len(iterations) == 0:
            var_name = field.FullName if field.Properties["py_setColumnName"] is None else field.Properties["py_setColumnName"]
        else:
            if field.Properties["py_setColumnName"] is None:
                var_name = field.FullName.replace("..","%s") % tuple(field.CurrentIndexPath.split(','))
            else:
                var_name = field.Properties["py_setColumnName"] % tuple(re.sub(pattern="[\{\}]",repl='',string=field.CurrentIndexPath).split(','))
            
            #var_name = "{}_{}".format(var_name, self.get_iteration_name(iterations))

        #var_name = var_name if var_name[0:1] not in ["_"] else var_name[1:len(var_name)]
        
        #if field.UsageType == categoryUsageConstants.vtVariable.value:
        #    var_name = var_name.replace(".", "_")
                
        return var_name
        
    def replaceLabel(self, label):
        blacklist = ["SHOWTABLET", "SHOW TABLET", "SHOWTABLET THANG ĐIỂM", "SHOWPHOTO"]
        s = re.sub(pattern=".*(?=({}))".format("|".join(blacklist)), repl="", string=label)
        s = re.sub(pattern="({})".format("|".join(blacklist)), repl="", string=s)
        return s 

    def get_iteration_label(self, iterations=list()):
        labels = list()

        def get_label(index):
            if index < 0:
                return

            labels.append(iterations[index].Label)
            get_label(index - 1) 
        
        get_label(len(iterations) - 1)
        return "_".join(labels)

    def get_variable_label(self, field, iterations=list()):
        m = re.match(pattern="^(.*)[\.](.*)\?", string= field.Label)
        s = field.Label if m is None else field.Label[m.regs[0][0]:m.regs[0][1]]
        
        if len(iterations) > 0:
            s = "{} - {}".format(s, self.get_iteration_label(iterations=iterations))
        return s

    def to_spss(self):
        with savReaderWriter.SavWriter(self.mdd_file.replace(".mdd", "_SPSS.sav"), varNames=self.varNames, missingValues=self.varMissingValues, varTypes=self.varTypes, formats=self.var_date_formats, varLabels=self.varLabels, measureLevels=self.measureLevels, valueLabels=self.valueLabels, ioUtf8=True) as writer:
            for i, row in self.df_spss.iterrows():
                for v in self.var_dates:    
                    if re.match(pattern="(.*)DATE(.*)", string=v):
                        try:
                            d = datetime.strptime(row[v], "%m/%d/%Y")
                            row[v] = writer.spssDateTime(datetime.strftime(d, "%m/%d/%Y").encode('utf-8'), "%m/%d/%Y")
                        except:
                            row[v] = np.nan
                        
                    if re.match(pattern="(.*)TIME(.*)", string=v):
                        try:
                            d = datetime.strptime(row[v], "%H:%M:%S")
                            row[v] = writer.convertTime(d.day, d.hour, d.minute, d.second)
                        except:
                            row[v] = np.nan

                writer.writerow(list(row))

class SPSSObject(mrDataFileDsc):
    def __init__(self, mdd_path, ddf_path, sql_query, questions, group_name=None):
        mrDataFileDsc.__init__(self, mdd_file=mdd_path, ddf_file=ddf_path, sql_query=sql_query)

        self.records = list()
        self.varNames = list()
        self.varTypes = dict()
        self.valueLabels = dict()
        self.varLabels = dict()
        self.measureLevels = dict()
        self.var_date_formats = dict()
        self.var_dates = list()

        #Group of columns with format A_suffix1, A_suffix2...B_suffix1, B_suffix2
        self.group_name = group_name
        #Columns to use as id variable
        self.id_vars = list()
        self.value_groupnames = list()
        self.groupname_labels = list()

        self.initialize(questions)

    def initialize(self, questions):
        self.openMDM()
        self.openDataSource()
        
        if self.group_name is not None:
            self.value_groupnames.extend(self.getValueLabels(self.MDM.Fields[self.group_name]).keys())
        
        index_record = 0

        while not self.adoRS.EOF:
            self.initializeVariable(questions, index_record)
            index_record += 1
            self.adoRS.MoveNext()
            
        self.closeMDM()
        self.closeDataSource()
        
    def initializeVariable(self, questions, index_record):
        r = list()

        for question in questions:
            field = self.MDM.Fields[question]

            if field.Properties["py_isHidden"] is None or field.Properties["py_isHidden"] == False:
                r.extend(self.createVariable(field, index_record, verify_id_var=(question != self.group_name))) 
        
        self.records.append(r)
    
    #verify_id_var: check whether a column is an id variable.

    def createVariable(self, field, index_record, verify_id_var=False):
        r = list()

        if str(field.ObjectTypeValue) == objectTypeConstants.mtVariable.value:
            if index_record == 0:
                self.createProperties(field, verify_id_var=verify_id_var)

            r.extend(self.transformData(field))
        elif str(field.ObjectTypeValue) == objectTypeConstants.mtRoutingItems.value:
            if index_record == 0:
                self.createProperties(field, verify_id_var=verify_id_var)

            r.extend(self.transformData(field))
        else:
            match str(field.ObjectTypeValue):
                case objectTypeConstants.mtClass.value:
                    for f in field.Fields:
                        if f.Properties["py_isHidden"] is None or f.Properties["py_isHidden"] == False:
                            r.extend(self.createVariable(f, index_record, verify_id_var=verify_id_var))
                case objectTypeConstants.mtArray.value:
                    for variable in field.Variables:
                        if variable.Properties["py_isHidden"] is None or variable.Properties["py_isHidden"] == False:
                            if variable.UsageType != 1048:
                                r.extend(self.createVariable(variable, index_record, verify_id_var=verify_id_var))
        return r
    
    def createProperties(self, field, verify_id_var=False):
        
        var_name = self.getVariableName(field)

        if "QME" in var_name:
            a = ""

        var_label = self.replaceLabel(self.getVariableLabel(field))
        
        if field.DataType == dataTypeConstants.mtCategorical.value:
            
            if field.Properties["py_showPunchingData"]:
                value_labels = { 1 : "Yes".encode('utf-8'), 0 : "No".encode('utf-8') }
                
                categories_list = list()
                other_categories_list = list()

                if field.Properties["py_setCategoriesList"]:
                    categories_list = field.Properties["py_setCategoriesList"].split(',')

                    other_categories_list = [cat.Name for cat in field.Categories if cat.Name not in categories_list]
                else:
                    categories_list = [cat.Name for cat in field.Categories if not cat.IsOtherLocal]

                    other_categories_list = [cat.Name for cat in field.Categories if cat.Name not in categories_list]

                for cat_name in categories_list:
                    category = field.Categories[cat_name]

                    self.setVariable(
                        var_name="{}{}".format(var_name, category.Name), 
                        var_label="{}_{}".format(var_label, category.Label).encode('utf-8'),
                        var_types=self.getDataType(field),
                        measure_levels=self.getMeasureLevel(field),
                        value_labels=value_labels, 
                        verify_id_var=verify_id_var
                    )
                
                if field.Properties["py_showHelperFields"]:
                    if field.Properties["py_combibeHelperFields"]:
                        self.setVariable(
                            var_name="{}{}".format(var_name, "_97"), 
                            var_label="{}_{}".format(var_label, "Others").encode('utf-8'),
                            var_types=self.getDataType(field),
                            measure_levels=self.getMeasureLevel(field),
                            value_labels=value_labels, 
                            verify_id_var=verify_id_var
                        )
                    else:
                        for other_cat_name in other_categories_list:
                            helperfield = field.Categories[other_cat_name]

                            self.setVariable(
                                var_name="{}{}".format(var_name, helperfield.Name), 
                                var_label="{}_{}".format(var_label, helperfield.Label).encode('utf-8'),
                                var_types=self.getDataType(field),
                                measure_levels=self.getMeasureLevel(field),
                                value_labels=value_labels, 
                                verify_id_var=verify_id_var
                            )
                a = ""
                """
                for category in field.Categories:
                    if not category.IsOtherLocal:
                        self.setVariable(
                            var_name="{}{}".format(var_name, category.Name), 
                            var_label="{}_{}".format(var_label, category.Label).encode('utf-8'),
                            var_types=self.getDataType(field),
                            measure_levels=self.getMeasureLevel(field),
                            value_labels=value_labels, 
                            verify_id_var=verify_id_var
                        )
                    
                if field.Properties["py_showHelperFields"]:
                    if field.HelperFields.Count > 0:
                        if field.Properties["py_combibeHelperFields"]:
                            self.setVariable(
                                var_name="{}{}".format(var_name, "_C97"), 
                                var_label="{}_{}".format(var_label, "Others").encode('utf-8'),
                                var_types=self.getDataType(field),
                                measure_levels=self.getMeasureLevel(field),
                                value_labels=value_labels, 
                                verify_id_var=verify_id_var
                            )
                        else: 
                            for helperfield in field.HelperFields:
                                self.setVariable(
                                    var_name="{}{}".format(var_name, helperfield.Name), 
                                    var_label="{}_{}".format(var_label, helperfield.Label).encode('utf-8'),
                                    var_types=self.getDataType(field),
                                    measure_levels=self.getMeasureLevel(field),
                                    value_labels=value_labels, 
                                    verify_id_var=verify_id_var
                               )
                """
            else:
                value_labels = self.getValueLabels(field)

                max_range = field.MaxValue if field.MaxValue is not None else field.Categories.Count

                for i in range(max_range):                    
                    self.setVariable(
                        var_name=(var_name if field.MinValue == 1 and field.MaxValue == 1 else "{}_{}".format(var_name, i + 1)), 
                        var_label=var_label.encode('utf-8'),
                        var_types=self.getDataType(field),
                        measure_levels=self.getMeasureLevel(field),
                        value_labels=value_labels, 
                        verify_id_var=verify_id_var
                    )

                if field.Properties["py_showHelperFields"]:
                    if field.HelperFields.Count > 0:
                        if field.Properties["py_combibeHelperFields"]:
                            self.setVariable(
                                var_name="{}{}".format(var_name.decode('utf-8'), "_C97"), 
                                var_label=var_label,
                                var_types=self.getDataType(field),
                                measure_levels=self.getMeasureLevel(field),
                                value_labels=value_labels, 
                                verify_id_var=verify_id_var
                            )
                        else:
                            for helperfield in field.HelperFields:
                                self.setVariable(
                                    var_name="{}{}_Other".format(var_name, helperfield.Name),
                                    #var_name="{}{}#{}".format(var_name.split('#')[0], helperfield.Name.replace("_", "_C"), var_name.split('#')[1]), 
                                    var_label=var_label.encode('utf-8'),
                                    var_types=self.getDataType(helperfield),
                                    measure_levels=self.getMeasureLevel(helperfield),
                                    verify_id_var=verify_id_var
                                )
        elif field.DataType == dataTypeConstants.mtDate.value:
            self.setVariable(
                var_name=var_name, 
                var_label=var_label.encode('utf-8'),
                var_types=self.getDataType(field),
                measure_levels=self.getMeasureLevel(field),
                var_date_formats="ADATE10" if re.match(pattern="(.*)DATE(.*)", string=var_name) else "DATETIME20", 
                verify_id_var=verify_id_var
            )
        else:
            self.setVariable(
                var_name=var_name, 
                var_label=var_label.encode('utf-8'),
                var_types=self.getDataType(field),
                measure_levels=self.getMeasureLevel(field), 
                verify_id_var=verify_id_var
            )

    def setVariable(self, var_name="", var_label="", var_types=0, measure_levels="unknown", var_date_formats="", value_labels=dict(), verify_id_var=False):
        self.varNames.append(var_name)
        self.varLabels[var_name] = var_label
        self.varTypes[var_name] = var_types
        self.measureLevels[var_name] = measure_levels
        
        if len(var_date_formats) > 0:
            self.var_date_formats[var_name] = var_date_formats
            self.var_dates.append(var_name)

        if len(value_labels.keys()) > 0: self.valueLabels[var_name] = value_labels

        if verify_id_var is True:
            self.id_vars.append(var_name)
        else:
            g = re.sub(pattern="(#({}))$".format("|".join(self.value_groupnames)), repl="", string=var_name)
            
            if g not in self.groupname_labels:
                self.groupname_labels.append(g)

    def transformData(self, field):
        record = list()

        match field.DataType:
            case dataTypeConstants.mtLong.value | dataTypeConstants.mtDouble.value:
                record.append(np.nan if self.adoRS.Fields[field.FullName].Value is None else self.adoRS.Fields[field.FullName].Value)
            case dataTypeConstants.mtText.value:
                record.append('' if self.adoRS.Fields[field.FullName].Value is None else self.adoRS.Fields[field.FullName].Value)
            case dataTypeConstants.mtDate.value:
                d = self.adoRS.Fields[field.FullName].Value
                
                try:
                    if d.year == 1899:
                        record.append(d.strftime("%H:%M:%S"))
                    else:
                        record.append(d.strftime("%m/%d/%Y"))
                except:
                    record.append(np.nan)
            case dataTypeConstants.mtCategorical.value:
                categories_list = list()
                other_categories_list = list()

                if field.Properties["py_setCategoriesList"]:
                    categories_list = field.Properties["py_setCategoriesList"].split(',')

                    other_categories_list = [cat.Name for cat in field.Categories if cat.Name not in categories_list]
                else:
                    categories_list = [cat.Name for cat in field.Categories if not cat.IsOtherLocal]

                    other_categories_list = [cat.Name for cat in field.Categories if cat.Name not in categories_list]

                cats_resp = list()
                
                if self.adoRS.Fields[field.FullName].Value is not None:
                    if not re.match(pattern="^\{\}$", string=str(self.adoRS.Fields[field.FullName].Value)):
                        cats_resp = str(self.adoRS.Fields[field.FullName].Value)[1:(len(str(self.adoRS.Fields[field.FullName].Value))-1)].split(",")
                
                if field.Properties["py_showPunchingData"]:
                    for cat_name in categories_list:
                        category = field.Categories[cat_name]

                        if len(cats_resp) == 0:
                            record.append(np.nan)
                        else:
                            if category.Name in cats_resp:
                                if field.Properties["py_showVariableValues"] is None:
                                    record.append(1)
                                else:
                                    record.append(category.Label) 
                            else:
                                record.append(0)
                            """
                            if field.Properties["py_showVariableValues"] is None:
                                if category.Name in cats_resp:
                                    record.append(1)
                                else:
                                    record.append(0 if self.adoRS.Fields[field.FullName].Value is not None else np.nan)   
                            else:
                                if category.Name in cats_resp:
                                    record.append(category.Label) 
                                else:
                                    record.append(np.nan)
                            """

                    if field.Properties["py_showHelperFields"]:
                        if field.Properties["py_combibeHelperFields"]:
                            if len(cats_resp) == 0:
                                record.append(np.nan)
                            else:
                                if any(cat in other_categories_list for cat in cats_resp):
                                    record.append(1)
                                else:
                                    record.append(0)

                    """
                    for category in field.Categories:
                        if not category.IsOtherLocal:
                            if field.Properties["py_showVariableValues"] is None:
                                if category.Name in cats_resp:
                                    record.append(1)
                                else:
                                    record.append(0 if self.adoRS.Fields[field.FullName].Value is not None else np.nan)
                            else:
                                if category.Name in cats_resp:
                                    record.append(category.Label) 
                                else:
                                    record.append(np.nan)

                    if field.Properties["py_showHelperFields"]:
                        if field.HelperFields.Count > 0:
                            if field.Properties["py_combibeHelperFields"]:
                                str_others = ""

                                for helperfield in field.HelperFields:
                                    if helperfield.Name in cats_resp:
                                        str_others = str_others + (", " if len(str_others) > 0 else "") + self.adoRS.Fields["{}.{}".format(field.FullName, helperfield.Name)].Value
                                
                                if len(str_others) > 0:
                                    match field.Properties["py_showVariableValues"]:
                                        case "Names":
                                            record.append(field.Categories[helperfield.Name].Name.replace('_',''))
                                        case "Labels":
                                            record.append(field.Categories[helperfield.Name].Label)
                                        case _:
                                            record.append(1)
                                else:
                                    record.append("")
                            else: 
                                for helperfield in field.HelperFields:
                                    if field.Properties["py_showVariableValues"] is None:
                                        if helperfield.Name in cats_resp:
                                            record.append(1)
                                        else:
                                            record.append(0 if self.adoRS.Fields[field.FullName].Value is not None else np.nan)
                                    else:
                                        if helperfield.Name in cats_resp:
                                            record.append(helperfield.Label) 
                                        else:
                                            record.append("")
                    """
                else:
                    max_range = field.MaxValue if field.MaxValue is not None else field.Categories.Count

                    for i in range(max_range):
                        if i < len(cats_resp):
                            category = cats_resp[i]

                            match field.Properties["py_showVariableValues"]:
                                case "Names":
                                    record.append(np.nan if self.adoRS.Fields[field.FullName].Value == None else field.Categories[category].Name)
                                case "Labels":
                                    record.append(np.nan if self.adoRS.Fields[field.FullName].Value == None else field.Categories[category].Label)
                                case _:
                                    record.append(np.nan if self.adoRS.Fields[field.FullName].Value == None else int(category[1:len(category)]))
                        else:
                            record.append(np.nan)
                    
                    if field.Properties["py_showHelperFields"]:
                        if field.HelperFields.Count > 0:
                            if field.Properties["py_combibeHelperFields"]:
                                str_others = ""

                                for helperfield in field.HelperFields:
                                    if helperfield.Name in cats_resp:
                                        str_others = str_others + (", " if len(str_others) > 0 else "") + self.adoRS.Fields["{}.{}".format(field.FullName, helperfield.Name)].Value
                                
                                if len(str_others) > 0:
                                    match field.Properties["py_showVariableValues"]:
                                        case "Names":
                                            record.append(field.Categories[helperfield.Name].Name.replace('_',''))
                                        case "Labels":
                                            record.append(field.Categories[helperfield.Name].Label)
                                        case _:
                                            record.append(1)
                                else:
                                    record.append("")
                            else:
                                for helperfield in field.HelperFields:
                                    record.append("" if self.adoRS.Fields[helperfield.FullName].Value is None else self.adoRS.Fields[helperfield.FullName].Value)
                                    """
                                    if field.Properties["py_showVariableValues"] is None:
                                        if helperfield.Name in cats_resp:
                                            record.append(1)
                                        else:
                                            record.append("" if self.adoRS.Fields[helperfield.FullName].Value is not None else "")
                                    else:
                                        if helperfield.Name in cats_resp:
                                            record.append(helperfield.Label) 
                                        else:
                                            record.append("")
                                    """

        return record

    def getDataType(self, field):
        match field.DataType:
            case dataTypeConstants.mtLong.value | dataTypeConstants.mtDouble.value:
                return 0
            case dataTypeConstants.mtText.value:
                return 1024
            case dataTypeConstants.mtDate.value:
                return 0
            case _:
                return 0
    
    def getMeasureLevel(self, field):
        match field.DataType:
            case dataTypeConstants.mtLong.value | dataTypeConstants.mtDouble.value:
                return "scale"
            case dataTypeConstants.mtText.value:
                return "nominal"
            case _:
                return "nominal"

    def getValueLabels(self, field):
    
        cats = dict()

        for category in field.Categories:
            cat = category.Name if category.Name[0:1] != "_" else category.Name[1:len(category.Name)] 
            
            if cat not in cats:
                if cat.isnumeric():
                    cats[int(cat)] = category.Label.encode('utf-8')
                else:
                    cats[str(cat)] = category.Label.encode('utf-8')
        
        return cats
    
    def getVariableLabel(self, field):
        if str(field.ObjectTypeValue) == objectTypeConstants.mtVariable.value:
            return field.Label
        else:
            indexes = re.sub(pattern="[{}]", repl="", string=field.Indexes)
            indexes = indexes.split(',')

            var_label = "{} - {}".format(field.Label, self.getIterationLabel(field.Parent.Parent, indexes=indexes)) 
            return var_label

    def getIterationLabel(self, field, indexes=list()):
        if field.Parent.Parent is None:
            return field.Categories[indexes[field.LevelDepth - 1]].Label 
        else:
            return "{}_{}".format(field.Categories[indexes[field.LevelDepth - 1]].Label, self.getIterationName(field.Parent.Parent, indexes=indexes))
            #return "{}_{}".format(indexes[field.LevelDepth - 1], self.getIterationName(field.Parent.Parent, indexes=indexes))

    def getVariableName(self, field):
        if str(field.ObjectTypeValue) == objectTypeConstants.mtVariable.value:
            var_name = field.FullName if field.Properties["py_setColumnName"] is None else field.Properties["py_setColumnName"]
            var_name = var_name if var_name[0:1] not in ["_"] else var_name[1:len(var_name)]
            var_name = var_name.replace(".", "_")
            return var_name
        else:
            var_name = field.Name if field.Properties["py_setColumnName"] is None else field.Properties["py_setColumnName"]

            indexes = re.sub(pattern="[{}]", repl="", string=field.Indexes)
            indexes = indexes.split(',')

            iteration_name = self.getIterationName(field.Parent.Parent, indexes=indexes)
            var_name = "{}{}".format(var_name, iteration_name)

            return var_name

    def getIterationName(self, field, indexes=list()):
        if field.Parent.Parent is None:
            prefix = "#" if field.Name == self.group_name else "_"
            iteration_name = indexes[field.LevelDepth - 1]
            iteration_name = re.sub(pattern="^[_#]", repl="", string=iteration_name)
            iteration_name = "{}{}".format(prefix, iteration_name)

            return iteration_name
        else:
            iteration_name = self.getIterationName(field.Parent.Parent, indexes=indexes)
            return "{}{}".format(indexes[field.LevelDepth - 1], iteration_name)
    
    def replaceLabel(self, label):
        blacklist = ["SHOWTABLET", "SHOW TABLET", "SHOWTABLET THANG ĐIỂM", "SHOWPHOTO"]
        s = re.sub(pattern=".*(?=({}))".format("|".join(blacklist)), repl="", string=label)
        s = re.sub(pattern="({})".format("|".join(blacklist)), repl="", string=s)
        return s
    
    def to_spss_wide_to_long(self):
        df = pd.DataFrame(data=self.records, columns=self.varNames)

        df_unpivot = pd.wide_to_long(df, stubnames=self.groupname_labels, i=self.id_vars, j='Rotation', sep="#", suffix="(({}))$".format("|".join(self.value_groupnames)))
        df_unpivot.reset_index(inplace=True)
        
        var_names_unpivot = list()
        var_types_unpivot = dict()
        formats_unpivot = dict()
        var_labels_unpivot = dict()
        measure_levels_unpivot = dict()
        value_labels_unpivot = dict()
        var_dates_unpivot = list()

        for c in list(df_unpivot.columns):
            if c == "Rotation":
                var_names_unpivot.append(c)
                var_types_unpivot[c] = 10
            else:
                for v in self.varNames:
                    variable = re.sub(pattern="(#({}))$".format("|".join(self.value_groupnames)), repl="", string=v)
                    
                    if c == variable and variable not in var_names_unpivot:
                        var_names_unpivot.append(variable)
                        var_types_unpivot[variable] = self.varTypes[v]

                        if v in self.var_date_formats.keys():
                            formats_unpivot[variable] = self.var_date_formats[v]
                        if v in self.varLabels.keys():
                            var_labels_unpivot[variable] = self.varLabels[v]
                        if v in self.measureLevels.keys():
                            measure_levels_unpivot[variable] = self.measureLevels[v]
                        if v in self.valueLabels.keys():
                            value_labels_unpivot[variable] = self.valueLabels[v]
                        if v in self.var_dates:
                            var_dates_unpivot.append(c)
                        break
        
        with savReaderWriter.SavWriter(self.mdd_path.replace(".mdd", "unpivot.sav"), varNames=var_names_unpivot, varTypes=var_types_unpivot, formats=formats_unpivot, varLabels=var_labels_unpivot, measureLevels=measure_levels_unpivot, valueLabels=value_labels_unpivot, ioUtf8=True) as writer:
            for i, row in df_unpivot.iterrows():
                for v in var_dates_unpivot:
                    if re.match(pattern="(.*)DATE(.*)", string=v):
                        try:
                            d = datetime.strptime(row[v], "%m/%d/%Y")
                            row[v] = writer.spssDateTime(datetime.strftime(d, "%m/%d/%Y").encode('utf-8'), "%m/%d/%Y")
                        except:
                            row[v] = np.nan
                        
                    if re.match(pattern="(.*)TIME(.*)", string=v):
                        try:
                            d = datetime.strptime(row[v], "%H:%M:%S")
                            row[v] = writer.convertTime(d.day, d.hour, d.minute, d.second)
                        except:
                            row[v] = np.nan

                writer.writerow(list(row))
        
        df_queue = pd.DataFrame()
        df_queue = df.loc[:, self.id_vars]
        df_queue.set_index(["InstanceID"], inplace=True)
        
        self.groupname_labels.extend(["InstanceID","Rotation"])
        df_main_unpivot = df_unpivot.loc[:, self.groupname_labels]

        products = list(value_labels_unpivot["Product_Selected"].keys())
        
        for p in products:
            df_temp = df_main_unpivot.loc[(df_main_unpivot["Rotation"] != "Recall_1") & (df_main_unpivot["Product_Selected"] == p)]
            df_temp.drop(columns=["Rotation"], inplace=True, axis=1)
            df_temp.set_index(["InstanceID"], inplace=True)
            
            if not df_temp.empty:
                #Rename columns by product
                renamed_columns = dict()

                for c in df_temp.columns:
                    renamed_columns[c] = "PC_{}_{}".format(p, c)

                df_temp.rename(columns=renamed_columns, inplace=True)
                df_temp.reset_index(inplace=True)
                df_temp.set_index("InstanceID", inplace=True)

                df_queue = pd.concat([df_queue, df_temp], ignore_index=False, sort=False, axis=1)
        
        df_queue.reset_index(inplace=True)

        var_names_unpivot = list()
        var_types_unpivot = dict()
        formats_unpivot = dict()
        var_labels_unpivot = dict()
        measure_levels_unpivot = dict()
        value_labels_unpivot = dict()
        var_dates_unpivot = list()

        for c in list(df_queue.columns):
            if c == "Rotation":
                var_names_unpivot.append(c)
                var_types_unpivot[c] = 10
            else:
                for v in self.varNames:
                    v_temp = re.sub(pattern="(#({}))$".format("|".join(self.value_groupnames)), repl="", string=v)
                    c_temp = re.sub(pattern="(^PC_({})_)".format("|".join([str(i) for i in products])), repl="", string=c)
                    
                    if c_temp == v_temp and c not in var_names_unpivot:
                        var_names_unpivot.append(c)
                        var_types_unpivot[c] = self.varTypes[v]

                        if v in self.var_date_formats.keys():
                            formats_unpivot[c] = self.var_date_formats[v]
                        if v in self.varLabels.keys():
                            var_labels_unpivot[c] = self.varLabels[v]
                        if v in self.measureLevels.keys():
                            measure_levels_unpivot[c] = self.measureLevels[v]
                        if v in self.valueLabels.keys():
                            value_labels_unpivot[c] = self.valueLabels[v]
                        if v in self.var_dates:
                            var_dates_unpivot.append(c)
                        break
        
        arr_string = [k for k, v in var_types_unpivot.items() if v == 1024 and k != "InstanceID"]

        for c in arr_string:
            df_queue.loc[df_queue[c].isnull(), c] = ''

        with savReaderWriter.SavWriter(self.mdd_path.replace(".mdd", "unpivot2.sav"), varNames=var_names_unpivot, varTypes=var_types_unpivot, formats=formats_unpivot, varLabels=var_labels_unpivot, measureLevels=measure_levels_unpivot, valueLabels=value_labels_unpivot, ioUtf8=True) as writer:
            for i, row in df_queue.iterrows():
                for v in var_dates_unpivot:
                    if re.match(pattern="(.*)DATE(.*)", string=v):
                        try:
                            d = datetime.strptime(row[v], "%m/%d/%Y")
                            row[v] = writer.spssDateTime(datetime.strftime(d, "%m/%d/%Y").encode('utf-8'), "%m/%d/%Y")
                        except:
                            row[v] = np.nan
                        
                    if re.match(pattern="(.*)TIME(.*)", string=v):
                        try:
                            d = datetime.strptime(row[v], "%H:%M:%S")
                            row[v] = writer.convertTime(d.day, d.hour, d.minute, d.second)
                        except:
                            row[v] = np.nan
                        

                writer.writerow(list(row))
    
    def to_spss(self):
        df = pd.DataFrame(data=self.records, columns=self.varNames)
        
        with savReaderWriter.SavWriter(self.mdd_path.replace(".mdd", ".sav"), varNames=self.varNames, varTypes=self.varTypes, formats=self.var_date_formats, varLabels=self.varLabels, measureLevels=self.measureLevels, valueLabels=self.valueLabels, ioUtf8=True) as writer:
            for i, row in df.iterrows():
                for v in self.var_dates:    
                    if re.match(pattern="(.*)DATE(.*)", string=v):
                        try:
                            d = datetime.strptime(row[v], "%m/%d/%Y")
                            row[v] = writer.spssDateTime(datetime.strftime(d, "%m/%d/%Y").encode('utf-8'), "%m/%d/%Y")
                        except:
                            row[v] = np.nan
                        
                    if re.match(pattern="(.*)TIME(.*)", string=v):
                        try:
                            d = datetime.strptime(row[v], "%H:%M:%S")
                            row[v] = writer.convertTime(d.day, d.hour, d.minute, d.second)
                        except:
                            row[v] = np.nan

                writer.writerow(list(row))
        

    



