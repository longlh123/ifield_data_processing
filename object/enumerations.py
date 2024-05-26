from enum import Enum

class objectTypeConstants(Enum):
    mtUnknown = 'ff'
    mtVariable = '0'
    mtArray = '1'
    mtGrid = '2'
    mtClass = '3'
    mtElement = '4'
    mtElements = '5'
    mtLabel = '6'
    mtField = '7'
    mtHelperFields = '8'
    mtFields = '9'
    mtTypes = 'A'
    mtProperties = 'B'
    mtRouting = 'C'
    mtContexts = 'D'
    mtLanguages = 'E'
    mtLevelObject = 'F'
    mtVariableInstance = '10'
    mtRoutingItem = '11'
    mtCompound = '12'
    mtElementInstance = '13'
    mtElementInstances = '14'
    mtLanguage = '15'
    mtRoutingItems = '16'
    mtRanges = '17'
    mtCategories = '18'
    mtCategoryMap = '19'
    mtDataSources = '1A'
    mtDocument = '1B'
    mtVersion = '1D'
    mtVersions = '1E'
    mtVariables = '1F'
    mtDataSource = '20'
    mtAliasMap = '21'
    mtIndexElement = '22'
    mtIndicesElements = '23'
    mtPages = '24'
    mtParameters = '25'
    mtPage = '26'
    mtItems = '27'
    mtContext = '28'
    mtContextAlternatives = '29'
    mtElementList = '2A'
    mtGoto = '2B'
    mtTemplate = '2C'
    mtTemplates = '2D'
    mtStyle = '2E'
    mtNote = '2F'
    mtNotes = '30'
    mtIfBlock = '31'
    mtConditionalRouting = '32'
    mtDBElements = '33'
    mtDBQuestionDataProvider = '34'

class dataTypeConstants(Enum):
    mtNone = 0 
    mtLong = 1 
    mtText = 2 
    mtCategorical = 3 
    mtObject = 4 
    mtDate = 5 
    mtDouble = 6 
    mtBoolean = 7 
    mtLevel = 8 

class categoryFlagConstants(Enum):
    flNone = 0
    flOther = 16
    flExclusive = 4291  

class categoryUsageConstants(Enum):
    vtVariable = 0
    vtOtherSpecify = 1040  


#flNone          = &H0000
#flUser          = &H0001
#flDontknow      = &H0002
#flRefuse        = &H0004
#flNoanswer      = &H0008
#flOther         = &H0010
#flMultiplier    = &H0020
#flExclusive     = &H1000
#flFixedPosition = &H0040
#flNoFilter      = &H0080
#flInline        = &H0100