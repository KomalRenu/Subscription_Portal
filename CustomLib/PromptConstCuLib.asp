<%'** Copyright � 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
'Const OBJECT_BROWSING_FLAG = DssXmlObjectDefn + DssXmlObjectDepnBrowser + DssXmlObjectDepnDates + DssXmlObjectDepnSecurity + DssXmlObjectAncestors
Const OBJECT_BROWSING_FLAG = 268453377
'DssXmlObjectAncestors(268435456)+DssXmlObjectDepnSecurity(16384)+DssXmlObjectDepnDates(1024)+DssXmlObjectDepnBrowser(512)+DssXmlObjectDefn(1)

Const CONST_ELEPROMPT_BLOCKCOUNT = 30
Const CONST_OBJPROMPT_BLOCKCOUNT = 30
Const CONST_PROMTINDEX_BLOCKCOUNT = 5

'Customized Err Numbers
Const ERR_CUSTOM_NO_PROMPT_INDEX = -1
Const ERR_CUSTOM_NO_PROMPT_TYPE = -2
Const ERR_CUSTOM_NO_PROMPT_STYLE = -3
Const ERR_CUSTOM_NO_SPECIFIC_NODE = -4
Const ERR_CUSTOM_NO_SPECIFIC_ATTRIBUTE = -5
Const ERR_CUSTOM_UNKNOWN_PROMPT_TYPE = -6
Const ERR_VALIDATION_FAILED = -7
Const ERR_ONLYADDNONE_HIPROMPT = -8
Const ERR_GET_PROMPTQUESTION_FROMSERVER = -9
Const ERR_LOAD_PROMPT_QUESTION = -10
Const ERR_DISPLAY_FAILED = -11
Const ERR_CUSTOM_UNKNOWN_EXPPROMPT_TYPE = -12
Const ERR_CUSTOM_UNKNOWN_EXPPROMPT_OPERATOR = -13
Const ERR_ANSWERPROMPT = -14
Const ERR_PROJECT_NAME_NOT_EXIST = -15
Const ERR_GET_HYDRA_PROMPT = -16
Const ERR_NEED_PROFILE_ANSWER = -17
Const ERR_MORE_PROMPTS = -18
Const ERR_UNSUPPORTED_PROMPTS = -19

'Error number for displaying error messages
'Const ERR_BETWEEN_EXPECTS_TWOVALUES = 165 'asDescriptors(165)
'Const ERR_NOTBETWEEN_EXPECTS_TWOVALUES = 100 'asDescriptors(100)
'Const ERR_IN_EXPECTS_VALUE = 320 'asDescriptors(320)
'Const ERR_OPERATOR_EXPECTS_VALUE = 167 'asDescriptors(167)
'Const ERR_SELECTELEM_BEFOREDRILL = 169 'asDescriptors(169)
'Const ERR_NOT_QUALIFY_TWO = 170 'asDescriptors(170)
'Const ERR_NO_OPEN_PROMPTS = 171 'asDescriptors(171)
'Const ERR_TOOMANY_SELECTIONS_OBJECTPROMPT = 161 'asDescriptors(161)
'Const ERR_TOOFEW_SELECTIONS_OBJECTPROMPT = 162 'asDescriptors(162)
'Const ERR_TOOMANY_SELECTIONS_ELEMENTPROMPT = 163 'asDescriptors(163)
'Const ERR_TOOFEW_SELECTIONS_ELEMENTPROMPT = 164 'asDescriptors(164)
'Const ERR_TOOMANY_SELECTIONS_LEVELPROMPT = 161 'asDescriptors(161)
'Const ERR_TOOFEW_SELECTIONS_LEVELPROMPT = 162 'asDescriptors(162)
'Const ERR_TOOMANY_SELECTIONS_EXPRESSIONPROMPT = 238 'asDescriptors(238)
'Const ERR_REQUIRED_PROMPT = 166 'asDescriptors(166)
'Const ERR_NOT_SUPPORT_DATA_TYPE_NONTEXT = 237 'asDescriptors(237)
'Const ERR_NOT_DISPLAY_DEFAULT_ANSWER = 243 'asDescriptors(243)
'Const ERR_TOOMANY_SELECTIONS_EXPPROMPT = 238 'asDescriptors(238)
'Const ERR_TOOFEW_SELECTIONS_EXPPROMPT = 239 'asDescriptors(239)
'Const ERR_TOOMANY_SELECTIONS_HIPROMPT = 240 'asDescriptors(240)
'Const ERR_TOOFEW_SELECTIONS_HIPROMPT = 241 'asDescriptors(241)
'Const ERR_TOOLONG_TEXT_CONSTANTPROMPT = 265 'asDescriptors(265)
'Const ERR_TOOSHORT_TEXT_CONSTANTPROMPT = 266 'asDescriptors(266)
'Const ERR_NOT_QUALIFY_ZERO = 330 'asDescriptors(330)
'Const ERR_TEXTFILE_NOT_VALID = 330 'TEMP!!! 'asDescriptors(330)
'Const ERR_SERVER_NOT_FOUND = 160 'asDescriptors(160)

Const ERR_BETWEEN_EXPECTS_TWOVALUES = 472
Const ERR_NOTBETWEEN_EXPECTS_TWOVALUES = 177
Const ERR_IN_EXPECTS_VALUE = 972
Const ERR_OPERATOR_EXPECTS_VALUE = 476
Const ERR_SELECTELEM_BEFOREDRILL = 481
Const ERR_NOT_QUALIFY_TWO = 482
Const ERR_NO_OPEN_PROMPTS = 483
Const ERR_TOOMANY_SELECTIONS_OBJECTPROMPT = 464
Const ERR_TOOFEW_SELECTIONS_OBJECTPROMPT = 465
Const ERR_TOOMANY_SELECTIONS_ELEMENTPROMPT = 466
Const ERR_TOOFEW_SELECTIONS_ELEMENTPROMPT = 467
Const ERR_TOOMANY_SELECTIONS_LEVELPROMPT = 464
Const ERR_TOOFEW_SELECTIONS_LEVELPROMPT = 465
Const ERR_TOOMANY_SELECTIONS_EXPRESSIONPROMPT = 592
Const ERR_REQUIRED_PROMPT = 475
Const ERR_NOT_SUPPORT_DATA_TYPE_NONTEXT = 591
Const ERR_NOT_DISPLAY_DEFAULT_ANSWER = 608
Const ERR_TOOFEW_SELECTIONS_EXPPROMPT = 594
Const ERR_TOOMANY_SELECTIONS_HIPROMPT = 595
Const ERR_TOOFEW_SELECTIONS_HIPROMPT = 596
Const ERR_TOOLONG_TEXT_CONSTANTPROMPT = 741
Const ERR_TOOSHORT_TEXT_CONSTANTPROMPT = 742
Const ERR_NOT_QUALIFY_ZERO = 1007
Const ERR_ANSWER_MUST_BE_A_DATE = 2361
Const ERR_ANSWER_MUST_BE_NUMERIC = 2360
Const ERR_TEXTFILE_NOT_VALID = 1007 'TEMP!!!
Const ERR_ALL_ANSWERS_MUST_BE_NUMERIC = 2413
Const ERR_ALL_ANSWERS_MUST_BE_A_DATE = 2414



'Global Variables
Dim lMaxID

lMaxID = 10
'Hydra
Const FOR_APPENDING = 8

'Prompt General Info
Const MAX_PROMPTGENERAL_INFO = 71
Const PROMPT_S_MSGID = 0
Const PROMPT_B_ISDOC = 1
Const PROMPT_S_REPORTID = 2
Const PROMPT_S_DOCUMENTID = 3
Const PROMPT_S_VIEWMODE = 4
Const PROMPT_B_REDIRECTTOREBUILD = 5
Const PROMPT_B_MESSAGESAVED = 6
Const PROMPT_B_REPROMPT = 7
Const PROMPT_B_XML = 8
Const PROMPT_B_DHTML = 9
Const PROMPT_B_NEEDPROCESS = 10
Const PROMPT_S_CURORDER = 11
Const PROMPT_L_MAXPIN = 12
Const PROMPT_L_ACTIVEPROMPT = 13
Const PROMPT_B_SAVE = 14
Const PROMPT_B_SENDANSWER = 15
Const PROMPT_B_VALIDATE = 16
Const PROMPT_B_CANCEL = 17
Const PROMPT_B_EXECUTE = 18
Const PROMPT_B_REEXECUTE = 19
Const PROMPT_B_ANYERROR = 20
Const PROMPT_S_BETWEENSEPERATOR = 21
Const PROMPT_S_INSEPERATOR = 22
Const PROMPT_B_REQUIREDFIRST = 23
Const PROMPT_B_ALLPROMPTSINONEPAGE = 24
Const PROMPT_S_TARGETPAGE = 25
Const PROMPT_O_PROMPTSOBJECT = 26
Const PROMPT_B_SUMMARY = 27
Const PROMPT_B_BACK = 28
Const B_ADD_SUBSCRIPTION_PROMPT = 29
Const B_EDIT_SUBSCRIPTION_PROMPT = 30
Const B_TRIGGERS_ONLY_PROMPT = 31
Const S_TRIGGER_ID_PROMPT = 32
Const B_DISPLAY_TRIGGER_PROMPT = 33
Const S_OLD_TRIGGER_ID_PROMPT = 34
Const B_DISPLAY_SUBSCRIBE_BUTTON = 35
Const PROMPT_B_ANY_TEXTFILE = 36
Const PROMPT_O_QUESTIONSXML = 37
Const PROMPT_O_QUESTIONPIFS = 38
Const PROMPT_O_DISPLAYXML = 39
Const PROMPT_S_ANSWERSXML = 40
Const PROMPT_O_TEMPANSWERSXML = 41
Const PROMPT_S_FILTERID = 42
Const PROMPT_S_TEMPLATEID = 43
Const PROMPT_B_SPECIAL_FORM = 44
Const PROMPT_N_TRIGGERS_COUNT = 45
    'Narrowcast Integration:
Const PROMPT_B_USE_NC = 46
Const PROMPT_B_REEXECUTED = 47
Const PROMPT_B_DISABLE_SAVE = 48

'Hydra
Const PROMPT_S_SUBSCRIPTIONGUID = 51
Const PROMPT_S_QUESTIONOBJECT_ID = 52
Const PROMPT_S_INFORMATIONSOURCE_ID = 53
Const PROMPT_L_ISM_TYPE = 54
Const PROMPT_S_SECURITY_FILTERID = 55
Const PROMPT_S_SECURITY_PROMPTID = 56
Const PROMPT_S_SRC = 57
Const PROMPT_S_PROFILE_NAME = 58
Const PROMPT_S_PROFILE_ORIGINAL_NAME = 59
Const PROMPT_S_PROFILE_DESC = 60
Const PROMPT_O_HYDRAPROMPTS = 61
Const PROMPT_S_FOLDERID = 62
Const PROMPT_B_FINISH_ENABLED = 63
Const PROMPT_B_USER_DEFAULT = 64
Const PROMPT_S_STATUS_FLAG = 65
Const PROMPT_S_PREF_ID = 66
Const PROMPT_B_ALLOW_PROFILE = 67
Const PROMPT_B_CHANGE_STYLE = 68
Const PROMPT_B_NOSUBS = 69

'API Errors
Const APIERROR_PROMPT_MESSAGE_CANNOT_BE_RETRIEVED = -2147468986
Const APIERROR_PROMPT_FAIL_LOAD_OBJECT_FROM_XML = -2147217090
Const APIERROR_PROMPT_MESSAGE_NOT_FOUND = -2147206852
Const APIERROR_PROMPT_ANSWER_PROMPT = -2147217091
Const APIERROR_SERVER_NOT_FOUND = -2147206923

'Helper Errors
Const HELPERERROR_PROMPT_REQUIRED = -2147205101
Const HELPERERROR_PROMPT_TOOMANY = -2147205099
Const HELPERERROR_PROMPT_TOOFEW = -2147205100

'option values
Const USER_OPTION_DHTML_YES = "3"
Const USER_OPTION_DHTML_NO = "4"
Const ADMIN_OPTION_DHTML_ALWAYSNO = "2"

'flag for search field
Public Const SEARCHFIELD_ELEPROMPT = 1
Public Const SEARCHFIELD_HIPROMPT = 2
Public Const SEARCHFIELD_HIPROMPT_BEFOREDRILL = 3

'flag for Attibute Qualifier
Public Const AQ_ATTRQUAL = 1
Public Const AQ_HIPROMPT = 2

'flag for Qualification type
Public Const QUAL_ATTRIBUTE = 1
Public Const QUAL_METRIC = 2

'flag for getting blockcount
Public Const BLOCKCOUNT_ELEPROMPT = 1
Public Const BLOCKCOUNT_OBJPROMPT = 2

'prompt info array index
Public Const MAX_PROMPTINFO_S_INDEX = 20
Public Const PROMPTINFO_S_INDEX = 1
Public Const PROMPTINFO_S_XSLFILE = 2
Public Const PROMPTINFO_B_ISCART = 3
'Public Const PROMPTINFO_S_TYPE = 4     'XXXXX
'Public Const PROMPTINFO_S_SUBTYPE = 5  'XXXXX
Public Const PROMPTINFO_B_REQUIRED = 6  'XXXXX
'Public Const PROMPTINFO_L_MIN = 7      'XXXXX
'Public Const PROMPTINFO_L_MAX = 8      'XXXXX
Public Const PROMPTINFO_S_STEP = 9
Public Const PROMPTINFO_S_MSG = 10
Public Const PROMPTINFO_B_ISALLDIMENSION = 11
'Public Const PROMPTINFO_B_UNKNOWN_DEFANSWER = 12
'Public Const PROMPTINFO_S_USED = 13        'XXXXX
'Public Const PROMPTINFO_S_CLOSED = 14  'XXXXX
Public Const PROMPTINFO_O_QUESTION = 15
Public Const PROMPTINFO_O_ANSWER = 16
Public Const PROMPTINFO_O_TEMPANSWER = 17

'Metric Operator Type
Const OperatorType_Metric = "M"
Const OperatorType_Rank = "R"
Const OperatorType_Percent = "P"

'indent TAB
Const PROMPT_INDENTTAB = "&nbsp;&nbsp;&nbsp;&nbsp;"


'From Helper Definition

'EnumDSSXMLOperatorType
Const DssXmlOperatorGeneric = 1
Const DssXmlOperatorPercent = 3
Const DssXmlOperatorRank = 2

'EnumDSSXMLMRPFunction
Const DssXmlMRPFunctionBetween = 3
Const DssXmlMRPFunctionBottom = 2
Const DssXmlMRPFunctionDifferentFrom = 8
Const DssXmlMRPFunctionEquals = 7
Const DssXmlMRPFunctionExcludeBottom = 5
Const DssXmlMRPFunctionExcludeTop = 4
Const DssXmlMRPFunctionNotBetween = 6
Const DssXmlMRPFunctionTop = 1

'EnumDSSXMLAnswerFormat
Const DssXmlAnswerFormatDefault = 1
Const DssXmlAnswerFormatFlat = 2

'status flag for Hydra
Const EDIT_SUBSCRIPTION = "0"
Const CREATE_SUBSCRIPTION = "1"

Private Const S_OBJECT_ID_OBJECT = 0
Private Const S_NAME_OBJECT = 1
Private Const L_TYPE_OBJECT = 2
Private Const S_DESCRIPTION_OBJECT = 3
Private Const S_ROOT_ID_OBJECT = 4
Private Const S_ROOT_NAME_OBJECT = 5
Private Const S_PARENT_ID_OBJECT = 6
Private Const S_PARENT_NAME_OBJECT = 7
Private Const N_REPORTS_OBJECT = 8
Private Const N_DOCUMENTS_OBJECT = 9
Private Const N_FOLDERS_OBJECT = 10
Private Const L_BROWSING_FLAGS_OBJECT = 11
Private Const S_SPECIAL_OBJECT_ID_OBJECT = 12
Private Const N_NUMBER_OF_FOLDERS_TO_SHOW_OBJECT = 13
Private Const N_NUMBER_OF_REPORTS_TO_SHOW_OBJECT = 14
Private Const N_NUMBER_OF_DOCUMENTS_TO_SHOW_OBJECT = 15
Private Const L_MAXIMUM_OBJECTS_OBJECT = 16
Private Const L_START_OBJECT_OBJECT = 17
Private Const N_VIEW_MODE_OBJECT = 18
Private Const S_TARGET_PAGE_OBJECT = 19
Private Const S_URL_OBJECT = 20
Private Const S_DHTML_FUNCTION_OBJECT = 21
Private Const O_CONTENTS_XML_OBJECT = 22
Private Const B_SHOW_LOCATION_OBJECT = 23
Private Const B_ALLOW_BROWSING_LINKS_ONLY_OBJECT = 24
Private Const B_SHOW_MORE_OBJECTS_OBJECT = 25
Private Const B_SHOW_COMPLETE_PATH_OBJECT = 26
Private Const B_REQUEST_RECEIVED_OBJECT = 27
Private Const S_TIME_ZONE_OBJECT = 28
Private Const MAX_OBJECT_INFO = 28

%>
