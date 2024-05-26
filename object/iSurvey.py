import numpy as np
import re
import xml.etree.ElementTree as ET
from object.enumerations import dataTypeConstants

class iSurvey(dict):
    def __init__(self, xml_file):
        self.__dict__ = dict()
        self.openXML(xml_file)

    def openXML(self, xml_file):
        tree = ET.parse(xml_file)
        root = tree.getroot()

        header = root.find('header')
        self["properties"] = header.find('surveyProperties').attrib
        self["title"] = header.find('title').text
        self["subtitle"] = header.find('subTitle').text 
        self["survey_family"] = header.find('surveyFamily').text 

        answersref = root.find('answersRef')
        self["answersref"] = iAnswersRef(answersref)

        body = root.find('body')
        self["questions"] = iQuestions(body, self["answersref"])
    
class iQuestions(dict):
    def __init__(self, body, answersref):
        self.__dict__ = dict()
        self.generate(body, answersref)

    def generate(self, body, answersref):
        parent_nodes = list()
        
        for question in body.findall('*'):
            
            if question.tag in ["sectionEnd", "loopEnd"]:
                parent_nodes = list()
            else:
                if question.attrib["surveyBuilderV3CMSObjGUID"] not in ["101622D0-8B7C-4DE5-B97B-67D33C2E51D7","0AB35540-8549-42F2-A4C4-EA793334170F"]:
                    object_name = question.attrib["objectName"]
                    
                    if len(parent_nodes) > 0:
                        object_name = "{}.{}".format(".".join([q["attributes"]["objectName"] for q in parent_nodes]), object_name)
                    
                    self[object_name] = iQuestion(question, answersref, parent_nodes=parent_nodes)

                    if question.tag in ["sectionStart", "loopStart"]:
                        parent_nodes.append(self[object_name]) 

class iQuestion(dict):
    def __init__(self, question, answersref, parent_nodes=None):
        self.__dict__ = dict()
        self.generate(question, answersref, parent_nodes=parent_nodes)
    
    def generate(self, question, answersref, parent_nodes=None):
        if question.attrib['pos'] == "56":
            a = ""
        self["text"] = self.get_text(question)
        self["attributes"] = question.attrib
        
        if question.find('answers') is not None:
            self["answers"] = iAnswers(question.find('answers'), answersref)
        
        if len(parent_nodes) > 0:
            self["parents"] = list()
            self["parents"].extend(parent_nodes)
        
        self["is_defined_list"] = False
        
        match question.attrib["surveyBuilderV3CMSObjGUID"]:
            case "8642F4F1-E3E3-480C-89C8-60EDC3DD65FC": #DataType.Text
                self["datatype"] = dataTypeConstants.mtText

                self["comment"] = question.find('comment').attrib
                self["syntax"] = self.syntax_comment()
                
                self["columns"] = self.get_columns()
            case "7AA1B118-B3CA-4112-A4BC-3AFEF497B034": #DataType.Date
                self["datatype"] = dataTypeConstants.mtDate

                self["comment"] = question.find('comment').attrib
                self["syntax"] = self.syntax_comment()
                
                self["columns"] = self.get_columns()
            case "FCE61FC3-99D3-455A-B635-517183475C26": #DataType.Media | DataType.Categorical
                if self["answers"]["attributes"]["answerSetID"] != "8":
                    self["datatype"] = dataTypeConstants.mtCategorical
                    self["syntax"] = self.syntax_categorical()

                    self["columns"] = self.get_columns()
                else:
                    self["datatype"] = dataTypeConstants.mtNone
            case "FA4B8A93-09EC-4E23-B45D-FB848C64B834": #DataType.Categorical
                self["datatype"] = dataTypeConstants.mtCategorical
                self["syntax"] = self.syntax_categorical()

                self["columns"] = self.get_columns()
            case "101622D0-8B7C-4DE5-B97B-67D33C2E51D7": #Display.png
                self["datatype"] = dataTypeConstants.mtNone
            case "F620C65C-1072-4CF0-B293-A9C9012F5BE8": #DataType.Define
                self["datatype"] = dataTypeConstants.mtNone
                self["is_defined_list"] = True
                self["syntax"] = self.syntax_define()
            case "2E46C5F3-AF64-4EB9-99D3-E920455F33B6": #DataType.Long | DataType.Double
                self["datatype"] = dataTypeConstants.mtDouble

                self["comment"] = question.find('comment').attrib
                self["syntax"] = self.syntax_comment()
                
                self["columns"] = self.get_columns()
            case "A7C7BA09-0741-4F80-A99F-24C8F045E0B0": #DataType.Loop
                self["datatype"] = dataTypeConstants.mtLevel
                self["objecttype"] = "Loop"
                self["syntax"] = self.syntax_loop()
            case "59BD961F-E403-4D86-95ED-6A740EEEB16B": #DataType.Loop
                self["datatype"] = dataTypeConstants.mtLevel
                self["objecttype"] = "Loop"
                self["syntax"] = self.syntax_loop()
            case "809CF49C-529D-4336-872A-24BE1C3DC37C": #DataType.BlockFields
                self["datatype"] = dataTypeConstants.mtLevel 
                self["objecttype"] = "BlockFields"
                self["syntax"] = self.syntax_block_fields()
            case "0AB35540-8549-42F2-A4C4-EA793334170F": #Section in iField
                self["datatype"] = dataTypeConstants.mtLevel  
                self["objecttype"] = "BlockFields"
                self["syntax"] = self.syntax_block_fields()
            case "90922453-5C1F-4A6A-BEF2-D4F5A805AD6B": #DataType.Object
                self["datatype"] = dataTypeConstants.mtObject

                self["comment"] = question.find('comment').attrib
                self["comment_syntax"] = self.syntax_general()
                self["syntax"] = self.syntax_categorical()

                self["columns"] = self.get_columns()
                
    def syntax_block_fields(self):
        s = "%s \"%s\" block fields();" % (self["attributes"]["objectName"], self["text"]) 
        return s

    def syntax_loop(self):
        s = "%s \"%s\" loop {%s}fields () expand grid;" % (
                self["attributes"]["objectName"], 
                self["text"], 
                self["answers"]["syntax"]
        )
        return s

    def syntax_define(self):
        s = '%s "" define{%s};' % (
            self["attributes"]["objectName"], 
            self["answers"]["syntax"]
        )
        return s

    def syntax_categorical(self):
        s = '%s "%s" categorical%s{%s};' % (
            self["attributes"]["objectName"], 
            self["text"], 
            "[1..1]" if self["answers"]["answerref"]["attributes"]["isMultipleSelection"] else "[1..]",
            self["answers"]["syntax"]
        )
        return s

    def syntax_comment(self):
        datatype = "text"

        match int(self["comment"]["datatype"]):
            case 2:
                datatype = "long" if self["comment"]["scale"] == 0 else "double"
            case 3:
                datatype = "text"
            case 4:
                datatype = "date"
                
        return f'{self["attributes"]["objectName"]} "{self["text"]}" {datatype};'

    def syntax_general(self):
        datatype = "text"

        match int(self["comment"]["datatype"]):
            case 2:
                datatype = "long" if self["comment"]["scale"] == 0 else "double"
            case 3:
                datatype = "text"
            case 4:
                datatype = "date"
                
        return f'{self["attributes"]["objectName"]}{self["comment"]["objectName"]} "{self["text"]}" {datatype};'

    def get_text(self, question):
        if question.find('text') is None: 
            return ""

        text = np.nan if question.find('text') is None else question.find('text').text

        text = re.sub(pattern='\<(?:\"[^\"]*\"[\'\"]*|\'[^\']*\'[\'\"]*|[^\'\">])+\>', repl="", string=text)
        text = re.sub(pattern='[\"\']', repl="", string=text)
        text = re.sub(pattern='\n', repl="", string=text)

        m = re.match(pattern="{#resource:(.+)}", string=text)

        if m is not None:
            if len(re.sub(pattern="{#resource:(.+)}", repl="", string=text)) == 0:
                text = re.sub(pattern="{#resource:|#}", repl="", string=text)
            else:
                text = re.sub(pattern="{#resource:(.+)}", repl="", string=text)

        m = re.search(pattern="([^.]*\.)(?=(.+))", string=text)

        if m is not None:
            text = text.replace(text[m.span()[0]:m.span()[1]], f'{re.sub(pattern="^_", repl="", string=question.attrib["objectName"]) }.')

        return text

    def get_columns(self):
        columns = []
        csv_columns = []

        def backtrack():
            if "parents" not in self.keys():
                match self["datatype"].value:
                    case dataTypeConstants.mtText.value:
                        columns.append(self["attributes"]["objectName"])
                        csv_columns.append(self["attributes"]["objectName"])
                    case dataTypeConstants.mtDate.value:
                        columns.append(self["attributes"]["objectName"])
                        csv_columns.append(self["attributes"]["objectName"])
                    case dataTypeConstants.mtDouble.value:
                        columns.append(self["attributes"]["objectName"])
                        csv_columns.append(self["attributes"]["objectName"])
                    case dataTypeConstants.mtCategorical.value:
                        columns.append(self["attributes"]["objectName"])

                        if bool(int(self["answers"]["answerref"]["attributes"]["isMultipleSelection"])):
                            for key, option in self["answers"]["options"].items():
                                if not bool(int(option["attributes"]["isDisplayAsHeader"])):
                                    csv_columns.append("%s.%s" % (self["attributes"]["objectName"], option["objectname"]))
                        else:
                            csv_columns.append(self["attributes"]["objectName"])
                        
                        for key, option in self["answers"]["options"].items():
                            if bool(int(option["attributes"]["isOtherSpecify"])):
                                csv_columns.append("%s.%s" % (self["attributes"]["objectName"], option["otherfield"].attrib["objectName"])) 
                    case dataTypeConstants.mtObject.value:
                        columns.append(self["attributes"]["objectName"])
                        columns.append(f'{self["attributes"]["objectName"]}{self["comment"]["objectName"]}')

                        csv_columns.append(self["attributes"]["objectName"])
                        csv_columns.append(f'{self["attributes"]["objectName"]}{self["comment"]["objectName"]}')
                return
            
            columns.extend([f'{p}.{self["attributes"]["objectName"]}' for p in self.get_parents(0, format="mdd")])
            csv_columns.extend([f'{p}.{self["attributes"]["objectName"]}' for p in self.get_parents(0, format="csv")])

        backtrack()
        
        return {"mdd" : columns, "csv" : csv_columns}
    
    def get_parents(self, index, format="mdd"):
        if index == len(self["parents"]):
            return 
        else:
            p1s = self.get_parent_columns(self["parents"][index], format=format)
            p2s = self.get_parents(index + 1)

            if p2s is None:
                return p1s
            else:
                return ["%s.%s" % (p1, p2) for p1 in p1s for p2 in p2s]

    def get_parent_columns(self, parent, format="mdd"):
        parent_columns = list()

        if parent["objecttype"] == "BlockFields":
            parent_columns.append(parent["attributes"]["objectName"]) 
        if parent["objecttype"] == "Loop":
            for key, option in parent["answers"]["options"].items():
                if not bool(int(option["attributes"]["isDisplayAsHeader"])):
                    if format == "mdd":
                        parent_columns.append('%s[{%s}]' % (parent["attributes"]["objectName"], option['objectname']))
                    elif format == "csv":
                        parent_columns.append('%s[%s]' % (parent["attributes"]["objectName"], option['objectname']))

        return parent_columns

class iAnswersRef(dict):
    def __init__(self, answersref):
        self.__dict__ = dict()
        self.generate(answersref)

    def generate(self, answersref):
        for answer in answersref.findall('answer'):
            if answer.attrib["id"] not in self:
                self[answer.attrib["id"]] = dict()
            
            #print(answer.attrib["id"])
            self[answer.attrib["id"]]["attributes"] = answer.attrib
            self[answer.attrib["id"]]["options"] = iOptions(answer.findall('option'))

class iAnswers(dict):
    def __init__(self, answers, answersref):
        self.__dict__ = dict()
        self.generate(answers, answersref)

    def generate(self, answers, answersref):
        self["attributes"] = answers.attrib
        self["answerref"] = answersref[self["attributes"]["answerSetID"]]
    
        if self["attributes"]["answerSetID"] != '8':
            self["options"] = iOptions(answers.find('options').findall('option'), self["answerref"]) 
            self["syntax"] = self.syntax()

    def syntax(self):
        s = [option["syntax"] for key, option in self["options"].items() if len(option["answersetreference"]) == 0]
        return ",".join(s)

class iOptions(dict):
    def __init__(self, *args):
        self.__dict__ = dict()

        if len(args) == 1 and args[0] is not None:
            self.generate(options=args[0])
        if len(args) == 2 and args[0] is not None and args[1] is not None:
            self.generate(options=args[0], answerref=args[1])
        
    def generate(self, options=dict(), answerref=dict()):
        for option in options:
            if option.attrib["pos"] not in self:
                self[option.attrib["pos"]] = iOption(option, answerref)

class iOption(dict):
    def __init__(self, *args):
        self.__dict__ = dict()
        self.generate(option=args[0], answerref=args[1])
        
    def generate(self, option=dict(), answerref=dict()):
        if len(answerref) == 0:
            self["text"] = "" if option.find('text') is None else option.find('text').text
            self["attributes"] = option.attrib
        else:
            self["text"] = self.format_text(answerref["options"][option.attrib["pos"]]["text"])
            self["objectname"] = option.attrib["objectName"]
            self["answersetreference"] = option.attrib["answerSetReference"]
            self["attributes"] = answerref["options"][option.attrib["pos"]]["attributes"]
            self["otherfield"] = option.find('otherField')
            self["syntax"] = self.syntax()
            
    def syntax(self):
        if bool(int(self["attributes"]["isDisplayAsHeader"])):
            s = "use %s" % (self["objectname"])
        else:
            s = '%s "%s" [ pos=%s, value="%s" ]' % (
                                self["objectname"], 
                                self["text"], 
                                self["attributes"]['pos'],
                                self["objectname"] if not re.match(pattern='^_(.*)$', string=self["objectname"]) else self["objectname"][1:len(self["objectname"])])

            if int(self["attributes"]['isOtherSpecify']) == 1:
                otherdisplaytype = "text"

                match int(self["attributes"]['otherDisplayType']):
                    case 2:
                        otherdisplaytype = "double"
                    case 3:
                        otherdisplaytype = "date"

                s = f'{s} other({self["objectname"]} "" {otherdisplaytype})'
                
            if int(self["attributes"]['isExclusive']) == 1:
                s = f'{s} dk'
            if int(self["attributes"]['isExclusive']) == 1:
                s = f'{s} fix'

        return s
    
    def format_text(self, text):
        text = re.sub(pattern='\<(?:\"[^\"]*\"[\'\"]*|\'[^\']*\'[\'\"]*|[^\'\">])+\>', repl="", string=text)
        text = re.sub(pattern='[\"\']', repl="", string=text)
        text = re.sub(pattern='\n', repl="", string=text)

        m = re.match(pattern="{#resource:(.+)}", string=text)

        if m is not None:
            if len(re.sub(pattern="{#resource:(.+)}", repl="", string=text)) == 0:
                text = re.sub(pattern="{#resource:|#}", repl="", string=text)
            else:
                text = re.sub(pattern="{#resource:(.+)}", repl="", string=text)

        return text

        