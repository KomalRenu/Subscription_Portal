<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/CommonCoLib.asp" -->
<!-- #include file="ReadOptionsCuLib.asp" -->
<!-- #include file="ErrorCuLib.asp" -->
<%
'Device descriptions
Public Const N_DEVICE_DESC_NONE = 0
Public Const N_DEVICE_DESC_TEXT = 1
Public Const N_DEVICE_DESC_HTML = 2
Public Const N_DEVICE_DESC_BOTH = 3

'Object Types
Public Const TYPE_LOCALE = "1"
Public Const TYPE_FOLDER = "2"
Public Const TYPE_QUESTION = "5"
Public Const TYPE_INFORMATION_SOURCE = "6"
Public Const TYPE_DEVICE = "15"
Public Const TYPE_PUBLICATION = "16"
Public Const TYPE_SUBSET = "17"
Public Const TYPE_SCHEDULE = "18"
Public Const TYPE_SERVICE = "19"

Public Const TYPE_SITE = "1001"
Public Const TYPE_CHANNEL = "1002"
Public Const TYPE_DBALIAS = "1003"
Public Const TYPE_PROFILE = "1004"
Public Const TYPE_DEVICE_TYPE = "1005"
Public Const TYPE_ADDRESS = "1006"
Public Const TYPE_LOGGINGLEVEL = "1007"
Public Const TYPE_PREFERENCEOBJECT = "1008"
Public Const TYPE_QUESTION_CONFIG = "1011"
Public Const TYPE_STORAGE_MAPPING = "1012"
Public Const TYPE_SUBSSET_CONFIG = "1013"
Public Const TYPE_SERVICE_CONFIG = "1017"

'User Property Constants
Public Const USER_PROP_USER_ID = 0
Public Const USER_PROP_USER_NAME = 1
Public Const USER_PROP_PASSWORD = 2
Public Const USER_PROP_HINT = 3
Public Const USER_PROP_LOCALE_ID = 4
Public Const USER_PROP_DEFAULT_ADDRESS_ID = 5
Public Const USER_PROP_AGREEMENT_ID = 6
Public Const USER_PROP_ACCOUNT_ID = 7
Public Const USER_PROP_STATUS = 8
Public Const USER_PROP_EXPIRATION_DATE = 9
Public Const USER_PROP_CREATED_DATE = 10
Public Const USER_PROP_CREATED_BY = 11
Public Const USER_PROP_LAST_MODIFIED_DATE = 12
Public Const USER_PROP_LAST_MODIFIED_BY = 13
Public Const MAX_USER_PROP = 13

'Address Property Constants
Public Const ADDR_PROP_ADDRESS_ID = 0
Public Const ADDR_PROP_ADDRESS_NAME = 1
Public Const ADDR_PROP_PHYSICAL_ADDRESS = 2
Public Const ADDR_PROP_ADDRESS_DISPLAY = 3
Public Const ADDR_PROP_DEVICE_ID = 4
Public Const ADDR_PROP_DELIVERY_WINDOW = 5
Public Const ADDR_PROP_TIMEZONE_ID = 6
Public Const ADDR_PROP_STATUS = 7
Public Const ADDR_PROP_CREATED_BY = 8
Public Const ADDR_PROP_LAST_MODIFIED_BY = 9
Public Const ADDR_PROP_TRANSMISSION_PROPERTIES_ID = 10
Public Const ADDR_PROP_PIN = 11
Public Const ADDR_PROP_EXPIRATION_DATE = 12
Public Const ADDR_PROP_CREATED_DATE = 13
Public Const ADDR_PROP_LAST_MODIFIED_DATE = 14
Public Const MAX_ADDR_PROP = 14

'Site settings
Public Const SUBS_CACHE_FILE = 1
Public Const SUBS_CACHE_SESSION = 2
Public Const MAX_NUM_TABS = 3
Public Const N_VIEW_LARGE_ICONS = 1
Public Const N_VIEW_LIST = 2
Public Const N_DEFAULT_SUBS_VIEW_MODE = 2
Public Const N_DEFAULT_SERV_VIEW_MODE = 1
Public Const DEFAULT_START_PAGE = "home"
Public Const S_PAGE_HOME = "home"
Public Const S_PAGE_SUBSCRIPTIONS = "subscriptions"
Public Const S_PAGE_ADDRESSES = "addresses"
Public Const S_PAGE_REPORTS = "reports"
Public Const S_PAGE_SERVICES = "services"
Public Const SITE_DEFAULT_NO_EXPIRATION = "12/31/9999"
Public Const S_VALID_CHARS_NUM_ADDRESS = "0123456789-()"

Public Const S_DEVICE_VALIDATION_EMAIL = "e"
Public Const S_DEVICE_VALIDATION_NUMBER = "n"
Public Const S_DEVICE_VALIDATION_NONE = "x"

'Cookies
Public Const USER_COOKIE = "usr"
Public Const SESSION_TOKEN = "st"
Public Const USER_PASSWORD = "upd"
Public Const SAVE_PASSWORD = "spd"
Public Const PORTAL_ADDRESS = "pad"
Public Const USER_COOKIE_SUMMARY_PAGE= "sum"

Public Const SITE_COOKIE = "stc"
Public Const CURRENT_SITE = "cst"
Public Const PREF_JAVASCRIPT = "pjs" '2: determine automatically, 1: yes, 2: no
Public Const USE_JAVASCRIPT = "js" '1: yes, 0: no
Public Const START_PAGE = "sp"
Public Const SITE_LOCALE = "loc"
Public Const SERVICE_VIEW_MODE = "sv"
Public Const SUBSCRIPTION_VIEW_MODE = "suv"
Public Const DELIVERY_ORDER_BY = "dob"
Public Const DELIVERY_SORT_ORDER = "dso"
Public Const REPORT_ORDER_BY = "rob"
Public Const REPORT_SORT_ORDER = "rso"

'Language Constants
Public Const SYSTEM_LOCALE_ID = "FBBF7C1E37EC11D4887C00C04F48F8FD"
Public Const GERMAN		= "1031"
Public Const ENGLISH_US	= "1033"
Public Const SPANISH_SP	= "3082"
Public Const FRENCH		= "1036"
Public Const ITALIAN	= "1040"
Public Const JAPANESE	= "1041"
Public Const KOREAN		= "1042"
Public Const PORTUGUESE_BR = "1046"
Public Const SWEDISH	= "1053"
Public Const CHINESE_SP	= "2052"
Public Const ENGLISH_UK	= "2057"
Public Const NUMBER_OF_DESCRIPTORS = 2000
Public Const NUMBER_OF_WEB_DESCRIPTORS = 3000

'Font constants
Public Const B_DOUBLE_BYTE_FONT = 0
Public Const N_SMALL_FONT = 1
Public Const N_MEDIUM_FONT = 2
Public Const N_LARGE_FONT = 3
Public Const S_FAMILY_FONT = 4
Public Const B_OVERWRITE_CSS_FONT = 5

Private Const MAX_FONT_INFO = 5

'*** Storage API ***'
Public Const SOURCE_COOKIES = 1
Public Const SOURCE_STORAGE_COMPONENT = 2

'*** PROGID constants ***'
Const PROGID_WEBOM = "MSIXMLLib.DSSXMLServerSession.1"
Const PROGID_STRING_UTIL = "M9StrUtl.StringUtilities"
Const PROGID_ADDRESS = "Bridge2API.Addresses"
Const PROGID_SITE_INFO   = "Bridge2API.SiteInfo"
Const PROGID_SYSTEM_INFO = "Bridge2API.SystemInfo"
Const PROGID_ADMIN       = "Bridge2API.Admin"
Const PROGID_PORTAL_VB_ADMIN = "MGAdmin.Admin"
Const PROGID_PERSONALIZATION_INFO = "Bridge2API.PersonalizationInfo"
Const PROGID_USER = "Bridge2API.User"
Const PROGID_SUBSCRIPTION = "Bridge2API.Subscription"
Const PROGID_DOC_REPOSITORY = "Bridge2API.DocRepository"
Const PROGID_GUID_GEN = "GuidGen.CGuidGen"
Const PROGID_BASE64 = "Base64Lib.Base64"
Const PROGID_NETSHARE = "MSTRNetShare.ShareEngineUtil2"
Const PROGID_NCS_GUI_LIB = "MSTRGUILibrary.Library"
Const PROGID_NCS_COM_TIME_ZONE= "MCCOMTIMEZONE.MSTRCOMTimeZone"
Const PROGID_TIMEZONE_LIB= "MSTRTimeZone.Library.1"


Public Const COOKIES_EXPIRATION_DATE = #1/1/2038#
Public Const EXPIRE_COOKIES_NOW = #1/1/1993#

Public Const MAIN_TOOLBAR = 1
Public Const SEARCH_TOOLBAR = 2
Public Const HELP_TOOLBAR = 4

'*** Control Constants ***'
Private Const INTEGER_VALUE = 0
Private Const FLOAT_VALUE = 1
Private Const POSITIVE_VALUE = 1
Private Const NEGATIVE_VALUE = 2
Private Const ZERO_VALUE = 4

'*** QO Flag Constants ***'
Const QO_TYPE_NORMAL = 0
Const QO_TYPE_CUSTOM = 1
Const QO_TYPE_SLICING = 2
Const QO_TYPE_CUSTOM_MAPWITHSUBINFO = 0
Const QO_TYPE_CUSTOM_MAPNOSUBINFO = 1
Const QO_TYPE_CUSTOM_NOMAPPING = 2

'*** ISM Flag Constants ***'
Const ISM_TYPE_CASTOR = 0
Const ISM_TYPE_USERDETAIL = 1
Const ISM_TYPE_CUSTOM = 2

'*** Site property values for SITE_PROP_SUMMARY_PAGE ***'
Const SITE_PROPVALUE_SUMMARY_PAGE_ALWAYS = 1
Const SITE_PROPVALUE_SUMMARY_PAGE_WHENMORETHANONEQO = 2
Const SITE_PROPVALUE_SUMMARY_PAGE_NEVER = 3

'Default Welcome Page:
Const WELCOME_PAGE = "welcome.asp"

'Application Types
Const APPLICATION_TYPE_PORTAL = 25

'Error codes returned by checkEngineConfig
Const CONFIG_OK     = 0
Const CONFIG_MISSING_ENGINE = 1
Const CONFIG_MISSING_MD     = 2
Const CONFIG_MISSING_SITE   = 4
Const CONFIG_MISSING_AUREP  = 8
Const CONFIG_MISSING_SBREP  = 16

'Site properties array values:
Const SITE_PROP_ID            = 0
Const SITE_PROP_NAME          = 1
Const SITE_PROP_DESC          = 2
Const SITE_PROP_NEW_USERS     = 3
Const SITE_PROP_NEW_LOCALE    = 4
Const SITE_PROP_NEW_EXPIRE    = 5
Const SITE_PROP_EXPIRE_VALUE  = 6
Const SITE_PROP_GUI_LANG      = 7
Const SITE_PROP_USE_DHTML     = 8
Const SITE_PROP_TMP_DIR       = 9
Const SITE_PROP_PROMPT_CACHE  = 10
Const SITE_PROP_EMAIL         = 11
Const SITE_PROP_PHONE         = 12
Const SITE_PROP_DEFAULT_ANSWER          = 13
Const SITE_PROP_AUREP_ID                = 14
Const SITE_PROP_AUREP                   = 15
Const SITE_PROP_AUREP_PREFIX            = 16
Const SITE_PROP_SBREP_ID                = 17
Const SITE_PROP_SBREP                   = 18
Const SITE_PROP_SBREP_PREFIX            = 19
Const SITE_PROP_PORTAL_DEV_ID           = 20
Const SITE_PROP_PORTAL_DEV_NAME         = 21
Const SITE_PROP_PORTAL_FOLDER_ID        = 22
Const SITE_PROP_DEFAULT_DEV_ID          = 23
Const SITE_PROP_DEFAULT_DEV_NAME        = 24
Const SITE_PROP_DEFAULT_FOLDER_ID       = 25
Const SITE_PROP_DEFAULT_DEV_VALIDATION  = 26
Const SITE_PROP_SUMMARY_PAGE            = 27
Const SITE_LOGIN_MODE                   = 28
Const SITE_AUTHENTICATION_SERVER_NAME   = 29
Const SITE_ELEMENT_PROMPT_BLOCK_COUNT   = 30
Const SITE_OBJECT_PROMPT_BLOCK_COUNT    = 31
Const SITE_PROP_STREAM_ATTACHMENTS      = 32
Const SITE_IS_DEFAULT				    = 33
Const SITE_PROMPT_MATCH_CASE			= 34
Const SITE_PROP_TIMEZONE	            = 35
Const SITE_AUTHENTICATION_SERVER_PORT   = 36
Const MAX_SITE_PROP = 36

Const FLAG_PROP_GROUP_NAME     = &H0001
Const FLAG_PROP_GROUP_CONN     = &H0002
Const FLAG_PROP_GROUP_OTHER    = &H0004
Const FLAG_PROP_GROUP_DEVICES  = &H0008
Const FLAG_PROP_GROUP_SERVICES = &H0010

'Hidden Objects Folder ID
CONST HIDDEN_OBJECTS_FOLDER_ID = "F3FD42809E1C11D5B25A00B0D024259E"

Const SITE_USES_CASTOR = "1"

Function SetLocaleInformation(asDescriptors, aFontInfo)
'*******************************************************************************
'Purpose: To get the descriptor array and set font settings
'Inputs:
'Outputs: asDescriptors, aFontInfo
'*******************************************************************************
	On Error Resume Next
	Dim sLanguage

	sLanguage = GetLng()

	If InStr(1, (KOREAN & "," & JAPANESE & "," & CHINESE_SP), sLanguage, 0) = 0 Then
		aFontInfo(B_DOUBLE_BYTE_FONT) = False
		aFontInfo(N_SMALL_FONT) = 1
		aFontInfo(N_MEDIUM_FONT) = 2
		aFontInfo(N_LARGE_FONT) = 3
		aFontInfo(S_FAMILY_FONT) = "Verdana,Arial,MS Sans Serif"
	Else
		aFontInfo(B_DOUBLE_BYTE_FONT) = True
		aFontInfo(N_SMALL_FONT) = 2
		aFontInfo(N_MEDIUM_FONT) = 3
		aFontInfo(N_LARGE_FONT) = 4
		aFontInfo(S_FAMILY_FONT) = "Verdana,Arial,MS Sans Serif"
	End If

	Call LoadDescriptors(sLanguage)
	asDescriptors = Application.Contents(sLanguage)

	SetLocaleInformation = Err.number
	Err.Clear
End Function


Function SetLocaleInformationForPromptPage(asDescriptors, aFontInfo)
'*******************************************************************************
'Purpose: To get the descriptor array and set font settings
'Inputs:
'Outputs: asDescriptors, aFontInfo
'*******************************************************************************
	On Error Resume Next
	Dim sLanguage

	sLanguage = GetLng()

	If InStr(1, (KOREAN & "," & JAPANESE & "," & CHINESE_SP), sLanguage, 0) = 0 Then
		aFontInfo(B_DOUBLE_BYTE_FONT) = False
		aFontInfo(N_SMALL_FONT) = 1
		aFontInfo(N_MEDIUM_FONT) = 2
		aFontInfo(N_LARGE_FONT) = 3
		aFontInfo(S_FAMILY_FONT) = "Verdana,Arial,MS Sans Serif"
	Else
		aFontInfo(B_DOUBLE_BYTE_FONT) = True
		aFontInfo(N_SMALL_FONT) = 2
		aFontInfo(N_MEDIUM_FONT) = 3
		aFontInfo(N_LARGE_FONT) = 4
		aFontInfo(S_FAMILY_FONT) = "Verdana,Arial,MS Sans Serif"
	End If

	Call LoadWebDescriptors(sLanguage)
	asDescriptors = Application.Contents("Web_" & sLanguage)

	SetLocaleInformation = Err.number
	Err.Clear
End Function

Function GetLng()
'*******************************************************************************
'Purpose: Gets the language setting from user's cookie.  If the cookie is blank,
'         gets the browser language.
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim sUserLanguage

	Dim sFilePath

	lErrNumber = NO_ERR

	'Read the right cookie based on where the request comes from
	sFilePath = Server.MapPath("./")
	If Right(sFilePath, 5) <> "admin" Then
		sUserLanguage = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, Array(SITE_COOKIE, "Lng", ""))
	End If

	If Len(sUserLanguage) <> 0 Then
		GetLng = sUserLanguage
	Else
	    GetLng = CStr(TransformLanguage(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE").Item))
	End If

	Err.Clear
End Function

Function TransformLanguage(sLanguage)
'*******************************************************************************
'Purpose: To transform the language into a decimal value
'Inputs:  sLanguage
'Outputs: The decimal value for the language
'*******************************************************************************
    On Error Resume Next
    Dim sTemp

	If IsLanguageSupported(sLanguage) And (sLanguage <> "") Then
		TransformLanguage = CStr(sLanguage)
	Else
		sTemp = Left(sLanguage, 2)
		Select Case CStr(sTemp)
		    Case "en"
	            TransformLanguage = ENGLISH_US
		    Case "es"
	            TransformLanguage = SPANISH_SP
		    Case "pt"
		        If StrComp(Left(sLanguage, 5), "pt-br", 1) = 0 Then
					TransformLanguage = PORTUGUESE_BR
		        End If
		    Case "fr"
		        TransformLanguage = FRENCH
		    Case "de"
		        TransformLanguage = GERMAN
		    Case "it"
		        TransformLanguage = ITALIAN
			Case "ja"
				TransformLanguage = JAPANESE
		    Case "ko"
		        TransformLanguage = KOREAN
		    Case "sv"
		        TransformLanguage = SWEDISH
		    Case "ch"
		    	TransformLanguage = CHINESE_SP
			Case Else
				TransformLanguage = ENGLISH_US
		End Select
	End If

    Err.Clear
End Function

Function LoadDescriptors(sLanguage)
'********************************************************
'*Purpose: Loads the descriptors for a given language into an
'          Application variable array
'*Inputs: sLanguage
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim szPath
	Dim oFs
	Dim oTextFile
	Dim iCount
	Dim sTemp
	Dim lErrNumber

	lErrNumber = NO_ERR

	Redim asDesc(NUMBER_OF_DESCRIPTORS)

	sTemp = CStr(Application.Contents(CStr(sLanguage))(0))
	Err.Clear

	If (InStr(CStr(Application("Languages_Loaded")), CStr(sLanguage)) = 0) Or Len(sTemp) = 0 Then
		'If not then load it and add the language to the loaded languages list

		'Get the Path
		szPath = Server.MapPath("Internationalization/MGNCS_" & sLanguage & ".txt")

		Set oFs = Server.CreateObject("Scripting.FileSystemObject")
		If(oFs.FileExists(szPath)) Then
			Set oTextFile = oFs.OpenTextFile(szPath)
		Else
			szPath = Server.MapPath("../Internationalization/MGNCS_" & sLanguage & ".txt")
			If (oFs.FileExists(szPath)) Then
				Set oTextFile = oFs.OpenTextFile(szPath)
			End If
		End If

		If(Not IsEmpty(oTextFile)) Then

			iCount = 0
			Do While oTextFile.AtEndOfStream <> True And Err.number = 0
			    asDesc(iCount) = oTextFile.ReadLine
			    iCount = iCount + 1
			Loop

			Application.Contents(CStr(sLanguage)) = asDesc
		End If
		oTextFile.Close
		If Err.number <> NO_ERR Then
            lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "LoadDescriptors", "", "Error loading descriptors for " & sLanguage, LogLevelError)
		Else
		    Application.Value("Languages_Loaded") = Application("Languages_Loaded") & sLanguage & ";"
		End If

	End If

    Set oFs = Nothing
    Set oTextFile = Nothing

	LoadDescriptors = lErrNumber
End Function

Function LoadWebDescriptors(sLanguage)
'********************************************************
'*Purpose: Loads the descriptors for a given language into an
'          Application variable array
'*Inputs: sLanguage
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim szPath
	Dim oFs
	Dim oTextFile
	Dim iCount
	Dim sTemp
	Dim lErrNumber

	lErrNumber = NO_ERR

	Redim asDesc(NUMBER_OF_WEB_DESCRIPTORS)

	sTemp = CStr(Application.Contents("Web_" & CStr(sLanguage))(0))
	Err.Clear

	If (InStr(CStr(Application("Web_Languages_Loaded")), CStr(sLanguage)) = 0) Or Len(sTemp) = 0 Then
		'If not then load it and add the language to the loaded languages list

		'Get the Path
		szPath = Server.MapPath("Internationalization/M9DSSWeb_" & sLanguage & ".txt")

		Set oFs = Server.CreateObject("Scripting.FileSystemObject")
		If(oFs.FileExists(szPath)) Then
			Set oTextFile = oFs.OpenTextFile(szPath)
		Else
			szPath = Server.MapPath("../Internationalization/M9DSSWeb_" & sLanguage & ".txt")
			If (oFs.FileExists(szPath)) Then
				Set oTextFile = oFs.OpenTextFile(szPath)
			End If
		End If

		If(Not IsEmpty(oTextFile)) Then

			iCount = 0
			Do While oTextFile.AtEndOfStream <> True And Err.number = 0
			    asDesc(iCount) = oTextFile.ReadLine
			    iCount = iCount + 1
			Loop

			Application.Contents("Web_" &CStr(sLanguage)) = asDesc
		End If
		oTextFile.Close
		If Err.number <> NO_ERR Then
            lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "LoadDescriptors", "", "Error loading descriptors for " & sLanguage, LogLevelError)
		Else
		    Application.Value("Web_Languages_Loaded") = Application("Web_Languages_Loaded") & sLanguage & ";"
		End If

	End If

    Set oFs = Nothing
    Set oTextFile = Nothing

	LoadDescriptors = lErrNumber
End Function

Function putMETATagWithCharSet()
'*******************************************************************************
'Purpose: To add the HTML with the charset for the current language
'Outputs: A String with the <META> tag
'*******************************************************************************
    On Error Resume Next
    Dim sMETATag
    Dim sUserLanguage
    Dim sWebServerLanguage

    sUserLanguage = GetLng()
	Select Case sUserLanguage
	    Case KOREAN
	        sMETATag = "<META HTTP-EQUIV=""Content-Type"" content=""text/html; charset=ks_c_5601-1987"" />"
	    Case SWEDISH
	        sMETATag = "<META HTTP-EQUIV=""Content-Type"" content=""text/html; charset=windows-1252"" />"
	    Case JAPANESE
	        sMETATag = "<META HTTP-EQUIV=""Content-Type"" Content=""text/html; charset=shift_jis"" />"
		Case GERMAN, ENGLISH_US, SPANISH_SP, FRENCH, ITALIAN, PORTUGUESE_BR
			sMETATag = "<META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" />"
		Case CHINESE_SP
			sMETATag = "<META HTTP-EQUIV=""Content-Type"" content=""text/html; charset=gb-2312"" />"
		Case Else
			sMETATag = ""
	End Select

    putMETATagWithCharSet = sMETATag
    Err.Clear
End Function

Function WriteToSource(aConnectionInfo, sValue, iSource, aSourceInfo)
'*******************************************************************************
'Purpose: Write a value to a source (currently only a cookie)
'Inputs:  aConnectionInfo, iSource, sValue, aSourceInfo
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim sPrimaryKey
	Dim sSecondaryKey
	Dim sExpirationDate

	sPrimaryKey = CStr(aSourceInfo(0))
	sSecondaryKey = CStr(aSourceInfo(1))
	Select Case iSource
		Case SOURCE_COOKIES
			sExpirationDate = CStr(aSourceInfo(2))

			If Len(sPrimaryKey) > 0 Then
				If Not IsEmpty(sValue) Then
					If Len(sSecondaryKey) > 0 Then
						Response.Cookies(sPrimaryKey)(sSecondaryKey) = sValue
					Else
						Response.Cookies(sPrimaryKey) = sValue
					End If
				End If
				If Len(sExpirationDate) > 0 Then
					Response.Cookies(sPrimaryKey).expires = sExpirationDate
				End If
			Else
				Call LogErrorXML(aConnectionInfo, 0, "Cookie not written", CStr(Err.source), "CommonLib.asp", "WriteToSource", "", "Error writing cookie: Empty key passed in", LogLevelError)
			End If
		Case SOURCE_STORAGE_COMPONENT
			If Not IsEmpty(sValue) Then
				Dim sGUID
				Dim oStorage
				Dim sFolderPath
				Dim oFolder
				sFolderPath = Server.MapPath("Preferences") & "\"
				Set oFolder = Server.CreateObject("Scripting.FileSystemObject")
				If Not (oFolder.FolderExists(sFolderPath)) Then
					oFolder.CreateFolder (sFolderPath)
				End If

				sGUID = RetrieveGUID()
				Set oStorage = Application.Value("StorageObject")
				If Len(sSecondaryKey) > 0 Then
					Call oStorage.Add(sValue, sGUID, sPrimaryKey, sSecondaryKey, sFolderPath & sGUID & ".xml")
				Else
					Call oStorage.Add(sValue, sGUID, sPrimaryKey, EMPTY, sFolderPath & sGUID & ".xml")
				End If

			    Set oStorage = Nothing
			    Set oFolder = Nothing
			End If
			If Err.number <> 0 Then
				Call LogErrorXML(aConnectionInfo, 0, "Could not write to VB Collection", CStr(Err.source), "CommonLib.asp", "WriteToSource", "", "Error writing to collection: Invalid string- " & sValue & "passed in", LogLevelError)
			End If
	End Select
End Function

Function ReadFromSource(aConnectionInfo, iSource, aSourceInfo)
'*******************************************************************************
'Purpose: Read a value from a source (currently only a cookie)
'Inputs:  aConnectionInfo, iSource, aSourceInfo
'Outputs: ReadFromSource
'*******************************************************************************
	On Error Resume Next
	Dim sPrimaryKey
	Dim sSecondaryKey
	Dim sGUID
	Dim oStorage

	sPrimaryKey = CStr(aSourceInfo(0))
	sSecondaryKey = CStr(aSourceInfo(1))

	Select Case iSource
		Case SOURCE_COOKIES
			If Len(sPrimaryKey) > 0 Then
				If Len(sSecondaryKey) > 0 Then
					ReadFromSource = Request.Cookies(sPrimaryKey)(sSecondaryKey)
				Else
					ReadFromSource = Request.Cookies(sPrimaryKey)
				End If
			Else
				Call LogErrorXML(aConnectionInfo, 0, "Cookie not read", CStr(Err.source), "CommonLib.asp", "ReadFromSource", "", "Error reading cookie: Empty key passed in", LogLevelError)
			End If
		Case SOURCE_STORAGE_COMPONENT
			sGUID = RetrieveGUID()
			Set oStorage = Application.Value("StorageObject")
			If Len(sPrimaryKey) > 0 Then
				If Len(sSecondaryKey) > 0 Then
					ReadFromSource = oStorage.Read(sGUID, sPrimaryKey, sSecondaryKey)
				Else
					ReadFromSource = oStorage.Read(sGUID, sPrimaryKey, EMPTY)
				End If
			Else
				Call LogErrorXML(aConnectionInfo, 0, "Collection not read", CStr(Err.source), "CommonLib.asp", "ReadFromSource", "", "Error reading VB Collection: Invalid string passed in", LogLevelError)
			End If
	End Select

	Set oStorage = Nothing

End Function

Function checkReturnValue(sReturnXML, sErrDesc)
'********************************************************
'*Purpose: checks XML returned by API for errors
'*Inputs: sReturnXML
'*Outputs: sErrDesc
'********************************************************
    Dim oDOM
    Dim lErr

    If Err.number <> NO_ERR Then
        Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		oDOM.async = False
		If oDOM.loadXML(Err.description) = False Then
			lErr = Err.number
			sErrDesc = Err.description
		Else
			If oDOM.selectSingleNode("mi/er") Is Nothing Then
				lErr = Err.number
				sErrDesc = Err.description
        	Else
				lErr = oDOM.selectSingleNode("mi/er").getAttribute("id")
				sErrDesc = oDOM.selectSingleNode("mi/er").getAttribute("des")
			End If
		End If
    Else
        Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
        oDOM.async = False
        If oDOM.loadXML(sReturnXML) = False Then
            lErr = ERR_XML_LOAD_FAILED
            sErrDesc = "Failed to load Results XML"
        Else
            If oDOM.selectSingleNode("mi/er") Is Nothing Then
                lErr = NO_ERR
            Else
                If oDOM.selectSingleNode("mi/er").getAttribute("id") = "0" Then
                    lErr = NO_ERR
                Else
                    lErr = oDOM.selectSingleNode("mi/er").getAttribute("id")
                    sErrDesc = oDOM.selectSingleNode("mi/er").getAttribute("des")
                End If
            End If
        End If

    End If

    Set oDOM = Nothing

    checkReturnValue = lErr
    Err.Clear
End Function

Function GetGUID()
'********************************************************
'*Purpose: Generates a new GUID and returns it
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim oGUID
	Dim lErrNumber
	lErrNumber = NO_ERR

	Set oGUID = Server.CreateObject(PROGID_GUID_GEN)
	If Err.number <> NO_ERR Then
	    lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "GetGUID", "", "Error creating " & PROGID_GUID_GEN, LogLevelError)
	Else
	    GetGUID = oGUID.GetGuid
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "GetGUID", "", "Error calling GetGuid method", LogLevelError)
	    End If
	End If

	Set oGUID = Nothing
End Function

Function Logout()
'********************************************************
'*Purpose: Sets user's SESSION_TOKEN cookie to empty
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim dCookieExpiration
	lErrNumber = NO_ERR

	If GetSavePasswordSetting() = "1" Then
	    dCookieExpiration = COOKIES_EXPIRATION_DATE
	Else
	    dCookieExpiration = ""
	End If

	aSourceInfo(0) = USER_COOKIE
	aSourceInfo(1) = SESSION_TOKEN
	aSourceInfo(2) = dCookieExpiration
	Call WriteToSource(aConnectionInfo, "", SOURCE_COOKIES, aSourceInfo)

	'Clear all the IServer cookies
	aSourceInfo(0) = ISERVER_USERNAME
	aSourceInfo(1) = ""
	aSourceInfo(2) = ""
	Call WriteToSource(aConnectionInfo, "", SOURCE_COOKIES, aSourceInfo)

	aSourceInfo(0) = ISERVER_PASSWORD
	aSourceInfo(1) = ""
	aSourceInfo(2) = ""
	Call WriteToSource(aConnectionInfo, "", SOURCE_COOKIES, aSourceInfo)

	aSourceInfo(0) = ISERVERUSER
	aSourceInfo(1) = ""
	aSourceInfo(2) = ""
	Call WriteToSource(aConnectionInfo, "", SOURCE_COOKIES, aSourceInfo)

	aSourceInfo(0) = ISERVERNTUSER
	aSourceInfo(1) = ""
	aSourceInfo(2) = ""
	Call WriteToSource(aConnectionInfo, "", SOURCE_COOKIES, aSourceInfo)


	Call Session.Abandon()

	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "Logout", "", "Error while calling WriteToSource", LogLevelError)
	End If

	Logout = lErrNumber
	Err.Clear
End Function

Function DisplayDateAndTime(vDate, vTime)
'*******************************************************************************
'Purpose: To return a string with a localized Date and Time
'Inputs:  vDate, vTime
'Outputs: A localized string with Date and Time
'*******************************************************************************
    On Error Resume Next
    Dim iYear
    Dim iMonth
    Dim iDay
    Dim iHours
    Dim iMinutes
    Dim iSeconds
    Dim sAMPM
    Dim sDateFormat
    Dim sTimeFormat
	Dim oTZ

    If Len(CStr(vDate)) > 0 Then

    	If Len(CStr(vDate)) > 10 Then
    		vDate = CDate(Left(CStr(vDate),10))
    	End If

    	If Len(CStr(vTime)) > 0 Then
			Set oTZ = Server.CreateObject(PROGID_NCS_COM_TIME_ZONE)
			oTZ.Name = Application.Value("Timezone")
			vDate = oTZ.LocalTime(CStr(vDate) & " " & Right(CStr(vTime), 11))
			vDate = CDate(Left(CStr(vDate),10))
    	End If

		iYear = CInt(Year(vDate))
		iMonth = CInt(Month(vDate))
		iDay = CInt(Day(vDate))
		sDateFormat = asDescriptors(327) 'Descriptor: MM/DD/YYYY
		If Len(sDateFormat) = 0 Then sDateFormat = "MM/DD/YYYY"
		sDateFormat = Replace(sDateFormat, "YYYY", iYear)
		If iMonth < 10 Then
			sDateFormat = Replace(sDateFormat, "MM", "0" & iMonth)
		Else
			sDateFormat = Replace(sDateFormat, "MM", iMonth)
		End If
		If iDay < 10 Then
			sDateFormat = Replace(sDateFormat, "DD", "0" & iDay)
		Else
			sDateFormat = Replace(sDateFormat, "DD", iDay)
		End If
	End If

    If Len(CStr(vTime)) > 0 Then

    	Set oTZ = Server.CreateObject(PROGID_NCS_COM_TIME_ZONE)
	    oTZ.Name = Application.Value("Timezone")
		vTime = oTZ.LocalTime(vTime)

		iHours = CInt(Hour(vTime))
		iMinutes = CInt(Minute(vTime))
		iSeconds = CInt(Second(vTime))
		sTimeFormat = asDescriptors(328) 'Descriptor: HH:MM:SS PM
		If Len(sTimeFormat) = 0 Then sTimeFormat = "HH:MM:SS PM"
		If InStr(1, sTimeFormat, "PM", 0) > 0 Then
			If iHours = 0 Then
				iHours = 12
				sAMPM = "AM"
			ElseIf iHours < 12 Then
				sAMPM = "AM"
			ElseIf iHours = 12 Then
				sAMPM = "PM"
			Else
				iHours = iHours - 12
				sAMPM = "PM"
			End If
			sTimeFormat = Replace(Replace(sTimeFormat, "HH", iHours), "PM", sAMPM)
		Else
			If iHours < 10 Then
				sTimeFormat = Replace(sTimeFormat, "HH", "0" & iHours)
			Else
				sTimeFormat = Replace(sTimeFormat, "HH", iHours)
			End If
		End If
		If iMinutes < 10 Then
			sTimeFormat = Replace(sTimeFormat, "MM", "0" & iMinutes)
		Else
			sTimeFormat = Replace(sTimeFormat, "MM", iMinutes)
		End If
		If iSeconds < 10 Then
			sTimeFormat = Replace(sTimeFormat, "SS", "0" & iSeconds)
		Else
			sTimeFormat = Replace(sTimeFormat, "SS", iSeconds)
		End If
	End If

    If Len(sTimeFormat) > 0 Then
    	DisplayDateAndTime = sDateFormat & " " & sTimeFormat & " " & Application.Value("Timezone")
    Else
    	DisplayDateAndTime = sDateFormat
    End If
    Err.Clear
End Function

Function cu_GetChannels(sChannelsXML)
'********************************************************
'*Purpose: Gets the Channels XML data from Application variable
'*Inputs:
'*Outputs: sChannelsXML
'********************************************************
Dim lErr
Dim sSiteID

    On Error Resume Next
    lErr = NO_ERR

    'Since the channels are the same for all objects, we store them
    'on an application variable:
    sChannelsXML = Application.Value("Channels_XML")

    cu_GetChannels = lErr
    Err.Clear
End Function

Function RenderSites()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const SITE_ACTIVE = 1
	Dim sChannelsXML
	Dim oSitesDOM
	Dim oSites
	Dim oCurrentSite
	Dim iNumSites
	Dim i
	Dim nStart
	Dim lErrNumber

	lErrNumber = NO_ERR
	iNumSites = 0
	nStart = 0

    lErrNumber = cu_GetChannels(sChannelsXML)
    If lErrNumber <> NO_ERR Then
        'add error handling
    End If

    If lErrNumber = NO_ERR Then
	    Set oSitesDOM = Server.CreateObject("Microsoft.XMLDOM")
	    oSitesDOM.async = False
	    If oSitesDOM.loadXML(sChannelsXML) = False Then
            lErrNumber = ERR_XML_LOAD_FAILED
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "RenderSites", "", "Error loading sChannelsXML", LogLevelError)
	    	'Add error message
	    Else
	        Set oSites = oSitesDOM.selectNodes("//oi[prs/pr[@id='active' and @v='1']]")
	        If Err.number <> NO_ERR Then
	            lErrNumber = Err.number
	            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "RenderSites", "", "Error retrieving site nodes", LogLevelError)
	    	    'Add error message
	        End If
	    End If
	End If

	If lErrNumber = NO_ERR Then

		iNumSites = oSites.length

	    If iNumSites > 0 Then
	    	'For Each oCurrentSite in oSites
	    	For i=0 To (oSites.length - 1)
	    		If i < MAX_NUM_TABS Then
	    			nStart = 0
	    		Else
	    			nStart = (i\MAX_NUM_TABS)*MAX_NUM_TABS
	    		End If
	    		If (i+1) Mod 2 = 1 Then
	    			Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
	    			Response.Write "<TR><TD WIDTH=""50%"">"
	    		Else
	    			Response.Write "<TD WIDTH=""50%"">"
	    		End If
	    		Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
	    		Response.Write "<TR>"
	    		Response.Write "<TD><A HREF=""" & GetStartPage() & ".asp?start=" & nStart & "&site=" & oSites.item(i).getAttribute("id") & """><img src=""images/project.gif"" HEIGHT=""60"" WIDTH=""60"" BORDER=""0"" ALT=""""></A></TD>"
	    		Response.Write "<TD VALIGN=""TOP""><A HREF=""" & GetStartPage() & ".asp?start=" & nStart & "&site=" & oSites.item(i).getAttribute("id") & """ STYLE=""text-decoration:none""><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """ color=""#cc0000""><b>" & oSites.item(i).getAttribute("n") & "</b></font></A><BR />"
	    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oSites.item(i).getAttribute("des") & "</font></TD>"
	    		Response.Write "</TR>"
	    		Response.Write "</TABLE>"
	    		If (i+1) Mod 2 = 1 Then
	    			Response.Write "</TD>"
	    			If i = (iNumSites-1) Then
	    				Response.Write "<TD></TD></TR></TABLE>"
	    			End If
	    		Else
	    			Response.Write "</TD></TR></TABLE><BR />"
	    		End If
	    	Next
	    End If
	End If

	Set oSitesDOM = Nothing
	Set oSites = Nothing
	Set oCurrentSite = Nothing

	RenderSites = lErrNumber
	Err.Clear
End Function

Function RenderTabs(sSiteID, sSelectedTabColor, sBarTitle, nStart)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const SITE_ACTIVE = 1
	Dim sChannelsXML
	Dim oTabsDOM
	Dim oTabs
	Dim oCurrentTab
	Dim sBackgroundColor
	Dim sFontColor
	Dim i
	Dim sName
	Dim sId
	Dim bMore
	Dim nSiteIndex
	Dim nEnd
	Dim nCount

	Dim lErrNumber

    lErrNumber = NO_ERR
	bMore = False
	nSiteIndex = 0

	lErrNumber = cu_GetChannels(sChannelsXML)
	If lErrNumber <> NO_ERR Then
	    'add error handling
	End If

    If lErrNumber = NO_ERR Then
	    Set oTabsDOM = Server.CreateObject("Microsoft.XMLDOM")
	    oTabsDOM.async = False
	    If oTabsDOM.loadXML(sChannelsXML) = False Then
	    	lErrNumber = ERR_XML_LOAD_FAILED
	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "RenderTabs", "", "Error loading sChannelsXML", LogLevelError)
	    	'Add error message
	    Else
	        Set oTabs = oTabsDOM.selectNodes("//oi[prs/pr[@id='active' and @v='1']]")
	        If Err.number <> NO_ERR Then
	            lErrNumber = Err.number
	            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "RenderTabs", "", "Error retrieving site nodes", LogLevelError)
	    	    'Add error message
	        End If
	    End If
	End If

	If lErrNumber = NO_ERR Then
        nCount = oTabs.length

        If nCount > 0 Then
            If nStart = -1 Then
	    		For i = 0 To (oTabs.length - 1)
	    			If oTabs.item(i).getAttribute("id") = sSiteID Then
	    				nSiteIndex = i
	    				Exit For
	    			End If
	    		Next
	    		If nSiteIndex < MAX_NUM_TABS Then
	    			nStart = 0
	    		Else
	    			nStart = (i\MAX_NUM_TABS)*MAX_NUM_TABS
	    		End If
	    	End If

            If nCount > MAX_NUM_TABS Then
                bMore = True
            End If

            If CInt(nStart) > 0 Then
                If CInt((MAX_NUM_TABS-1) + CInt(nStart)) > (nCount - 1) Then
                  nEnd = oTabs.length - 1
                Else
                  nEnd = CInt((MAX_NUM_TABS-1) + CInt(nStart))
                End If
            Else
                If CInt((MAX_NUM_TABS-1) + CInt(nStart)) < (nCount - 1) Then
                    nEnd = CInt((MAX_NUM_TABS - 1) + CInt(nStart))
                Else
                    nEnd = nCount - 1
                End If
            End If

	    	sBackgroundColor  = "cccccc"
            For i = nStart to nEnd
				sName =  oTabs.item(i).getAttribute("n")
				sId =  oTabs.item(i).getAttribute("id")

				sBackgroundColor = "AAAAAA"
				If sId = sSiteID Then
					sFontColor = "eedd82"
					sSelectedTabColor = sBackgroundColor
					sBarTitle = sName
				Else
				    sFontColor = "ffffff"
				End If

				Response.Write "<TD><img src=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
				Response.Write "<TD ALIGN=RIGHT VALIGN=BOTTOM >"
				Response.Write "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 HEIGHT=""100%"">"
				'Response.Write "<TR><TD HEIGHT=8 VALIGN=BOTTOM COLSPAN=3><IMG SRC=""images/tab_top_" & sBackgroundColor & ".gif"" WIDTH=""100%"" HEIGHT=""9"" BORDER=""0"" /></TD></TR>"
				'Response.Write "<TR><TD BGCOLOR=""#AAAAAA"" WIDTH=""1"" ROWSPAN=2><IMG SRC=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD><TD BGCOLOR=""#" & sBackgroundColor & """><IMG SRC=""images/1ptrans.gif"" WIDTH=""76"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD><TD BGCOLOR=""#000000"" WIDTH=""1"" ROWSPAN=""2""><IMG SRC=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD></TR>"
				Response.Write "<TR "
				'Response.Write "><TD HEIGHT=24 BGCOLOR=""#" & sBackgroundColor & """ ALIGN=""CENTER"" NOWRAP=""1"" WIDTH=""100%"""
				Response.Write "><TD HEIGHT=24 ALIGN=""CENTER"" NOWRAP=""1"" WIDTH=""100%"""
				Response.Write ">&nbsp;&nbsp;"
				If LoggedInStatus() = "" Then

				Else
				    Response.Write "<A HREF=""" & GetStartPage() & ".asp?start=" & nStart & "&site=" & sId & """ STYLE=""text-decoration:none;"">"
				End If
				Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#" & sFontColor & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """><B"
				Response.Write "><NOBR>"
				Response.Write Server.HTMLEncode(sName)
				Response.Write "</NOBR></B></FONT>"
				If LoggedInStatus() = "" Then

				Else
				    Response.Write "</A>"
				End If
			    Response.write "&nbsp;&nbsp;<BR />"
				Response.write "</TD"
				Response.write "></TR>"
				Response.write "</TABLE>"
				Response.write "</TD>"
            Next

	    	If bMore = True Then
	    		Response.Write "<TD><img src=""images/1ptrans.gif"" WIDTH=""2"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
	    		If nStart = 0 Then
	    			Response.Write "<TD><img src=""images/arrow_left_inc_fetch_disabled.gif"" WIDTH=""5"" HEIGHT=""10"" BORDER=""0"" ALT=""""></TD>"
	    		Else
	    			Response.Write "<TD><A HREF=""" & GetStartPage() & ".asp?start=" & CStr(CInt(nStart)- MAX_NUM_TABS) & "&site=" & oTabs.item(CInt(nStart) - MAX_NUM_TABS).getAttribute("id") & """><img src=""images/arrow_left_inc_fetch.gif"" WIDTH=""5"" HEIGHT=""10"" BORDER=""0"" ALT=""""></A></TD>"
	    		End If
	    		Response.Write "<TD><img src=""images/1ptrans.gif"" WIDTH=""4"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
	    		If nEnd >= (oTabs.length - 1) Then
	    			Response.Write "<TD><img src=""images/arrow_right_inc_fetch_disabled.gif"" WIDTH=""5"" HEIGHT=""10"" BORDER=""0"" ALT=""""></TD>"
	    		Else
	    		    If LoggedInStatus() = "" Then
	    		        Response.Write "<TD><img src=""images/arrow_right_inc_fetch_disabled.gif"" WIDTH=""5"" HEIGHT=""10"" BORDER=""0"" ALT=""""></TD>"
	    		    Else
	    			    Response.Write "<TD><A HREF=""" & GetStartPage() & ".asp?start=" & CStr(nEnd + 1) & "&site=" & oTabs.item(nEnd + 1).getAttribute("id") & """><img src=""images/arrow_right_inc_fetch.gif"" WIDTH=""5"" HEIGHT=""10"" BORDER=""0"" ALT=""""></A></TD>"
	    			End If
	    		End If
	    	End If
	    End If
	End If


	Set oTabsDOM = Nothing
	Set oTabs = Nothing
	Set oCurrentTab = Nothing

	RenderTabs = lErrNumber
	Err.Clear
End Function

Function getLoginMode()
	Dim loginMode
	Dim sUserName
	Dim sPassword
	Dim IServerUser

	IServerUser = false

	Call GetIServerCookies(sUserName,sPassword,IServerUser)

	loginMode = CStr(Application.Value("Login_Mode"))
	if( (loginMode="IS_NORMAL" or loginMode="NC_IS_NORMAL" or loginMode="NC_IS_NT_NORMAL" or loginMode="NT_NORMAL" or loginMode="IS_NT_NORMAL" or loginMode="NC_NT_NORMAL") And IServerUser = true) Then
		getLoginMode = SITE_USES_CASTOR
	else
		getLoginMode = "0"
	end if
End Function

Function getIServerName()
	'getIServerName = "malkovich"
	'Read the site properties and check for the server name
	Dim aSiteProperties()
	Dim lErr

	If Application.Value("Login_Authentication_Server_Name")= "" Then
		lErr = getSiteProperties(aSiteProperties)
		If lErr = 0 Then
			getIServerName = aSiteProperties(SITE_AUTHENTICATION_SERVER_NAME)
		Else
			getIServerName = ""
		End IF
	Else
		getIServerName = Application.Value("Login_Authentication_Server_Name")
	End If
End Function

Function getIServerPort()
	If Application.Value("Login_Authentication_Server_Port")= "" Then
		lErr = getSiteProperties(aSiteProperties)
		If lErr = 0 Then
			getIServerPort = aSiteProperties(SITE_AUTHENTICATION_SERVER_PORT)
		Else
			getIServerPort = 0
		End If
	Else
		getIServerPort = Application.Value("Login_Authentication_Server_Port")
	End If
End Function

Function getIServerUserName()
	getIServerUserName = Request.Cookies("ISERVER_USERNAME")
End Function

Function getIServerUserPassword()
	getIServerUserPassword = Decrypt(Request.Cookies("ISERVER_PASSWORD"))
End Function

Function getIServerNTUser()
	getIServerNTUser = Request.Cookies("ISERVERNTUSER")
End Function

Function GetNumberOfInformationSourceThatNeedAuthentication(sGetInformationSourcesForSiteXML, lNumberOfIS)
	On Error Resume Next
	Const SHOW_IS_LOGIN = "ISM_displayed"
	Dim oXMLDOM
	Dim oProjects
	Dim oProject

	lNumberOfIS = 0

    Set oXMLDOM = Server.CreateObject("Microsoft.XMLDOM")
    oXMLDOM.async = False
	Call oXMLDOM.loadXML(sGetInformationSourcesForSiteXML)

	lErrNumber = Err.Number
	If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, lErrNumber, Err.Description, CStr(Err.source), "CommonLib.asp", "LoadXMLDOMFromString", "loadXML()", "Error loading XML", LogLevelError)
	Else
   		Set oProjects = oXMLDOM.selectNodes("//oi[prs/pr[@id = '" & SHOW_IS_LOGIN & "']/@v = '1']")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "RenderInformationSourceLogins", "", "Error retrieving oi nodes", LogLevelError)
        Else
        	If GetLoginMode() <> SITE_USES_CASTOR Then
        		lNumberOfIS = oProjects.length
 			Else
 	           If oProjects.length > 0 Then
			   		For Each oProject in oProjects
				 		If GetIServerName() <> oProject.selectSingleNode("prs/pr[@id='" & "IS_ServerName" & "']").getAttribute("v") Then
				 			lNumberOfIS = lNumberOfIS + 1
				 		End If
				 	Next
				End If
			End If
	    End If
	End If
End Function


Function RenderInformationSourceLogins(sGetInformationSourcesForSiteXML, sUserAuthenticationObjectsXML, bIgnoreLoginMode)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const SHOW_IS_LOGIN = "ISM_displayed"
    Const IS_LOGIN_REQUIRED = "ISM_required"
    Const IS_MODULE = "ISM_admin_progid"
    Const ISM_CONN_INFO = "ISM_connInfo"
    Const ISM_PHYSICAL_ISID = "ISM_physical"
    Dim sISID
    Dim oOutputDOM
    Dim oProjects
    Dim oProject
    Dim oDecoder
    Dim sEncodedData
    Dim oDecodeDOM
    Dim lErrNumber
	Dim oUserAuthDOM
	Dim sReqFlag
	Dim oUserAuth
	Dim sUserName
	Dim sPwd
	Dim sUserID

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetInformationSourcesForSiteXML, oOutputDOM)
	If lErrNumber = NO_ERR Then
		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sUserAuthenticationObjectsXML, oUserAuthDOM)
	End If
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	End If

    If lErrNumber = NO_ERR Then
        Set oProjects = oOutputDOM.selectNodes("//oi[prs/pr[@id = '" & SHOW_IS_LOGIN & "']/@v = '1']")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "RenderInformationSourceLogins", "", "Error retrieving oi nodes", LogLevelError)
        Else
            If oProjects.length > 0 Then
				Set oDecoder = Server.CreateObject(PROGID_BASE64)

	    	    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AO_COUNT"" VALUE=""" & oProjects.length & """ />"
	    	    Response.Write "<TR><TD></TD><TD>"
	    	    For Each oProject in oProjects
	    			If bIgnoreLoginMode or GetLoginMode() <> SITE_USES_CASTOR or GetIServerName() <> oProject.selectSingleNode("prs/pr[@id='" & "IS_ServerName" & "']").getAttribute("v") Then
	    				sISID = oProject.selectSingleNode("prs/pr[@id = '" & ISM_PHYSICAL_ISID & "']").getAttribute("v")
	    				Set oUserAuth = oUserAuthDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_INFORMATION_SOURCE & "' $and$ @id='" & sISID & "']")
						If oUserAuth is Nothing Then
							sUserName = ""
							sPwd = ""
							sUserID = ""
						Else
							lErrNumber = ParseAuthenticationObject(oUserAuth.getAttribute("v"), sUserName, sPwd, sUserID)
						End If

	    				If oProject.selectSingleNode("prs/pr[@id = '" & IS_LOGIN_REQUIRED & "']").getAttribute("v") = "1" Then
	    				    sReqFlag = "<FONT COLOR=""#cc0000"">*</FONT>"
	    				    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AO_Req_" & sISID & """ VALUE=""1"" />"
	    				Else
	    				    sReqFlag = ""
	    				    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AO_Req_" & sISID & """ VALUE=""0"" />"
	    				End If
	    				'Response.Write "value is: " & oProject.selectSingleNode("prs/pr[@id='" & "IS_ServerName" & "']").getAttribute("v")
	    				If (oProject.selectSingleNode("prs/pr[@id='" & IS_MODULE & "']").getAttribute("v") = "CastorISM.cCastorISM") Then
	    					'If it is a Castor IS, then check if the IS should be displayed
	    					sEncodedData = oProject.selectSingleNode("prs/pr[@id='" & ISM_CONN_INFO & "']").text

							lErrNumber = LoadXMLDOMFromString(aConnectionInfo, oDecoder.Decode(sEncodedData), oDecodeDOM)

							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AO_Serv_" & sISID & """ VALUE=""" & oDecodeDOM.selectSingleNode("/info_source_props/server/primary").getAttribute("name") & """ />"
	    					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AO_Port_" & sISID & """ VALUE=""" & oDecodeDOM.selectSingleNode("/info_source_props/server/primary").getAttribute("port") & """ />"
	    					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AO_Proj_" & sISID & """ VALUE=""" & oDecodeDOM.selectSingleNode("/info_source_props/project").getAttribute("id") & """ />"
	    				Else
	    				    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AO_Serv_" & sISID & """ VALUE="""" />"
	    				    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AO_Port_" & sISID & """ VALUE="""" />"
	    				    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AO_Proj_" & sISID & """ VALUE="""" />"
	    				End If
	    				Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#000000"" SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & oProject.getAttribute("n") & "</b></FONT><BR />"
	    				Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#000000"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"
	    	    		Response.Write asDescriptors(369) & sReqFlag & "<BR />" 'Descriptor: User name:
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""AO_User_" & sISID & """ SIZE=""25"" MAXLENGTH=""250"" STYLE=""font-family: courier"" VALUE=""" & sUserName & """ /><BR />"
	    	    		Response.Write asDescriptors(370) & sReqFlag & "<BR />" 'Descriptor: Password:
	    	    		Response.Write "<INPUT TYPE=""PASSWORD"" NAME=""AO_Password_" & sISID & """ SIZE=""25"" MAXLENGTH=""250"" STYLE=""font-family: courier"" /><BR /><BR />"
	    	    		Response.Write "</FONT>"
	    	    	End If
	    	    Next
	    	    Response.Write "</TD></TR>"
            End If
        End If
    End If

    Set oOutputDOM = Nothing
    Set oProjects = Nothing
    Set oProject = Nothing
    Set oDecoder = Nothing
    Set oDecodeDOM = Nothing

    RenderInformationSourceLogins = lErrNumber
    Err.Clear
End Function

Function LoggedInStatus()
    On Error Resume Next

	aSourceInfo(0) = USER_COOKIE
	aSourceInfo(1) = SESSION_TOKEN
	LoggedInStatus = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function GetPortalAddress()
    On Error Resume Next

    aSourceInfo(0) = USER_COOKIE
    aSourceInfo(1) = PORTAL_ADDRESS
    GetPortalAddress = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function GetCurrentChannel()
    On Error Resume Next

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = CURRENT_SITE
    GetCurrentChannel = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function GetJavaScriptSetting()
    On Error Resume Next

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = USE_JAVASCRIPT
    GetJavaScriptSetting = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function GetSummaryPageSetting()
    On Error Resume Next

    aSourceInfo(0) = USER_COOKIE
    aSourceInfo(1) = USER_COOKIE_SUMMARY_PAGE
    GetSummaryPageSetting = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

    If (Len(GetSummaryPageSetting) = 0) Then
        GetSummaryPageSetting = Application("Default_display_summary")
    End If

    If (Len(GetSummaryPageSetting) = 0) Then
		GetSummaryPageSetting = CStr(SITE_PROPVALUE_SUMMARY_PAGE_WHENMORETHANONEQO)
    End If
End Function

Function GetJavaScriptPreference()
    On Error Resume Next

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = PREF_JAVASCRIPT
    GetJavaScriptPreference = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function SupportsJavaScript()
    On Error Resume Next
    Dim sUserBrowser
    Dim iPosition

	sUserBrowser = CStr(Request.ServerVariables("HTTP_USER_AGENT"))
	iPosition = InStr(1, sUserBrowser, "MSIE", vbTextCompare)
	If iPosition > 0 Then
		'IE
		If CLng(Mid(sUserBrowser, iPosition + 5, 1)) >= 4 Then
			SupportsJavaScript = "1"
		Else
			SupportsJavaScript = "0"
		End If
	ElseIf InStr(1, sUserBrowser, "Mozilla", vbTextCompare)	> 0 And InStr(1, sUserBrowser, "compatible", vbTextCompare) = 0 _
		And InStr(1, sUserBrowser, "sun", vbTextCompare) = 0 Then
		'Netscape
		iPosition = InStr(1, sUserBrowser, "Mozilla", vbTextCompare)
		If CLng(Mid(sUserBrowser, iPosition + 8, 1)) >= 4 Then
			SupportsJavaScript = "1"
		Else
			SupportsJavaScript = "0"
		End If
	Else 'Neither
		SupportsJavaScript = "0"
	End If

End Function

Function SetSiteLocale(sSiteLocale)
    On Error Resume Next

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = SITE_LOCALE
    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
    Call WriteToSource(aConnectionInfo, CStr(sSiteLocale), SOURCE_COOKIES, aSourceInfo)

End Function

Function GetSiteLocale()
    On Error Resume Next

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = SITE_LOCALE
    GetSiteLocale = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function GetLanguageSetting()
    On Error Resume Next

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = "Lng"
    GetLanguageSetting = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function GetUsername()
    On Error Resume Next

    aSourceInfo(0) = USER_COOKIE
    aSourceInfo(1) = "uname"
    GetUsername = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function SetSessionID(sSavePwd, sSessionID)
    On Error Resume Next
    Dim dUserCookieExpiration

   	If sSavePwd = "1" Then
	    dUserCookieExpiration = COOKIES_EXPIRATION_DATE
	Else
	    dUserCookieExpiration = ""
    End If

    aSourceInfo(0) = USER_COOKIE
    aSourceInfo(1) = SESSION_TOKEN
    aSourceInfo(2) = dUserCookieExpiration
    Call WriteToSource(aConnectionInfo, CStr(sSessionID), SOURCE_COOKIES, aSourceInfo)

End Function

Function GetSessionID()
    On Error Resume Next

    aSourceInfo(0) = USER_COOKIE
    aSourceInfo(1) = SESSION_TOKEN
    GetSessionID = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function GetSubscriptionsViewMode()
    On Error Resume Next
    Dim nViewMode

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = SUBSCRIPTION_VIEW_MODE
    nViewMode = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

    If Len(nViewMode) > 0 Then
        GetSubscriptionsViewMode = nViewMode
    Else
        GetSubscriptionsViewMode = N_DEFAULT_SUBS_VIEW_MODE
    End If

End Function

Function GetServiceViewMode()
    On Error Resume Next
    Dim nViewMode

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = SERVICE_VIEW_MODE
    nViewMode = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

    If Len(nViewMode) > 0 Then
        GetServiceViewMode = nViewMode
    Else
        GetServiceViewMode = N_DEFAULT_SERV_VIEW_MODE
    End If

End Function

Function SetSubscriptionViewMode(sViewMode)
    On Error Resume Next

	aSourceInfo(0) = SITE_COOKIE
	aSourceInfo(1) = SUBSCRIPTION_VIEW_MODE
    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	Call WriteToSource(aConnectionInfo, CStr(sViewMode), SOURCE_COOKIES, aSourceInfo)

End Function

Function SetServiceViewMode(sViewMode)
    On Error Resume Next

	aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = SERVICE_VIEW_MODE
    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	Call WriteToSource(aConnectionInfo, CStr(sViewMode), SOURCE_COOKIES, aSourceInfo)

End Function

Function SetReportsSorting(sRep_OrderBy, sRep_SortOrder)
    On Error Resume Next

	aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	aSourceInfo(1) = REPORT_ORDER_BY
	Call WriteToSource(aConnectionInfo, CStr(sRep_OrderBy), SOURCE_COOKIES, aSourceInfo)

	aSourceInfo(1) = REPORT_SORT_ORDER
	Call WriteToSource(aConnectionInfo, CStr(sRep_SortOrder), SOURCE_COOKIES, aSourceInfo)

End Function

Function SetDeliverySorting(sDeliv_OrderBy, sDeliv_SortOrder)
    On Error Resume Next

	aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	aSourceInfo(1) = DELIVERY_ORDER_BY
	Call WriteToSource(aConnectionInfo, CStr(sDeliv_OrderBy), SOURCE_COOKIES, aSourceInfo)

	aSourceInfo(1) = DELIVERY_SORT_ORDER
	Call WriteToSource(aConnectionInfo, CStr(sDeliv_SortOrder), SOURCE_COOKIES, aSourceInfo)
End Function

Function GetReportsOrderBy()
    On Error Resume Next
    Dim sRepOrderBy

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = REPORT_ORDER_BY
    sRepOrderBy = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

    If Len(sRepOrderBy) > 0 Then
        GetReportsOrderBy = sRepOrderBy
    Else
        GetReportsOrderBy = "TIME"
    End If
End Function

Function GetDeliveryOrderBy()
    On Error Resume Next
    Dim sDelOrderBy

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = DELIVERY_ORDER_BY
    sDelOrderBy = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

    If Len(sDelOrderBy) > 0 Then
        GetDeliveryOrderBy = sDelOrderBy
    Else
        GetDeliveryOrderBy = "TIME"
    End If
End Function

Function GetReportsSortOrder()
    On Error Resume Next
    Dim sRepSortOrder

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = REPORT_SORT_ORDER
    sRepSortOrder = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

    If Len(sRepSortOrder) > 0 Then
        GetReportsSortOrder = sRepSortOrder
    Else
        GetReportsSortOrder = "DESC"
    End If
End Function

Function GetDeliverySortOrder()
    On Error Resume Next
    Dim sDelSortOrder

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = DELIVERY_SORT_ORDER
    sDelSortOrder = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

    If Len(sDelSortOrder) > 0 Then
        GetDeliverySortOrder = sDelSortOrder
    Else
        GetDeliverySortOrder = "DESC"
    End If
End Function

Function SetStartPage(sStartPage)
    On Error Resume Next

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = START_PAGE
    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
    Call WriteToSource(aConnectionInfo, CStr(sStartPage), SOURCE_COOKIES, aSourceInfo)

End Function

Function GetStartPage()
    On Error Resume Next
    Dim sPage
    Dim sPortalDevice

    aSourceInfo(0) = SITE_COOKIE
    aSourceInfo(1) = START_PAGE
    sPage = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

    If Len(sPage) = 0 Then
        sPage = DEFAULT_START_PAGE
    Else
        'My reports page cannot be the default if not Portal device:
        If StrComp(sPage, S_PAGE_REPORTS) = 0 Then
            If Len(Application.Value("Portal_device")) = 0 Then
                sPage = DEFAULT_START_PAGE
            End If
        End If
    End If

    GetStartPage = sPage
    Err.Clear

End Function

Function GetSavePasswordSetting()
    On Error Resume Next

    aSourceInfo(0) = USER_COOKIE
    aSourceInfo(1) = SAVE_PASSWORD
    GetSavePasswordSetting = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

End Function

Function SetSummaryPageSetting(sSetting)
    On Error Resume Next

    aSourceInfo(0) = USER_COOKIE
    aSourceInfo(1) = USER_COOKIE_SUMMARY_PAGE
    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	Call WriteToSource(aConnectionInfo, CStr(sSetting), SOURCE_COOKIES, aSourceInfo)

End Function

Function SetHideChannelWizardIntro(sHide)
    On Error Resume Next

	aSourceInfo(0) = USER_COOKIE
	aSourceInfo(1) = SHOW_CHANNEL_WIZARD_INTRO
    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	Call WriteToSource(aConnectionInfo, CStr(sHide), SOURCE_COOKIES, aSourceInfo)

End Function

Function GetHideChannelWizardIntro()
    On Error Resume Next
    Dim sHide

    aSourceInfo(0) = USER_COOKIE
    aSourceInfo(1) = SHOW_CHANNEL_WIZARD_INTRO
    sHide = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)

    If Len(sHide) > 0 Then
        GetHideChannelWizardIntro = sHide
    Else
        GetHideChannelWizardIntro = "0"
    End If

End Function

Function RemoveParameterFromURL(sURL, sParameter)
'******************************************************************************
'Purpose: To remove a parameter from a URL
'Inputs:  sURL, sParameter
'Outputs: A string with the url without the parameter
'******************************************************************************
	On Error Resume Next
	Dim iInitialPos
	Dim iFinalPos
	Dim sTempURL

	sTempURL = CStr(sURL)
	If StrComp(Left(sTempURL, Len("&")), "&", vbBinaryCompare) = 0 Then
		sTempURL = Right(sTempURL, Len(sTempURL) - Len("&"))
	End If
	iInitialPos = InStr(1, sTempURL, "&" & sParameter & "=", vbTextCompare)
	If iInitialPos = 0 Then
		iInitialPos = InStr(1, sTempURL, "?" & sParameter & "=", vbTextCompare)
	End If
	If iInitialPos = 0 Then
		iInitialPos = InStr(1, sTempURL, sParameter & "=", vbTextCompare)
		If iInitialPos <> 1 Then
			iInitialPos = 0
		End If
	End If
	If iInitialPos > 0 Then
		If iInitialPos <> 1 Then
			iInitialPos = iInitialPos + Len("&")
		End If
		iFinalPos = InStr(iInitialPos, sTempURL, "&", vbTextCompare)
		If iFinalPos > 0 Then
			iFinalPos = iFinalPos + Len("&")
		End If
		If iInitialPos = 1 Then
			If iFinalPos > 0 Then
			    sTempURL = Mid(sTempURL, iFinalPos)
			Else
				sTempURL = ""
			End If
		Else
			sTempURL = Mid(sTempURL, 1, iInitialPos - 1)
			If iFinalPos > 0 Then
				sTempURL = sTempURL & Mid(sURL, iFinalPos)
			End If
		End If
	End If
	If StrComp(Right(sTempURL, Len("&")), "&", vbBinaryCompare) = 0 Then
		sTempURL = Left(sTempURL, Len(sTempURL) - Len("&"))
	End If
	If ((InStr(1, sTempURL, "&") = 0) And (InStr(1, sTempURL, "=") = 0) And (InStr(1, sTempURL, "?") = 0)) And Len(sTempURL) > 0 Then
		sTempURL = sTempURL & "?"
	End If

	RemoveParameterFromURL = sTempURL
	Err.Clear
End Function

Function ReplaceURLValue(oRequest, sFieldToChange, sValueToChange)
'******************************************************************************
'Purpose: To replace the value of a parameter with the given value
'Inputs:  oRequest, sFieldToChange, sValueToChange
'Outputs: A string representing the URL with the new value for the parameter
'******************************************************************************
    On Error Resume Next
    Dim sURL

    sURL = RemoveParameterFromURL(oRequest, sFieldToChange)
    If Len(sURL) > 0 Then
        If StrComp(Right(sURL, Len("?")), "?", vbBinaryCompare) = 0 Then
            sURL = sURL & sFieldToChange & "=" & Server.URLEncode(sValueToChange)
        Else
            sURL = sURL & "&" & sFieldToChange & "=" & Server.URLEncode(sValueToChange)
        End If
    Else
        sURL = sFieldToChange & "=" & Server.URLEncode(sValueToChange)
    End If

    ReplaceURLValue = sURL
    Err.Clear
End Function

Function AddInputsToXML(aConnectionInfo, oXML, asDictionary)
'*******************************************************************************
'Purpose: To receive a dictionary object and inserts it on a loaded XML object
'         If an error occurs the execution continues and the error is logged
'Inputs:  aConnectionInfo, oXML, asDictionary
'Outputs: sErrDescription, Err.number
'*******************************************************************************
    On Error Resume Next
    Dim i
    Dim root
    Dim oInputsElement
    Dim oInputsNode
    Dim lErrNumber

    root = oXML.documentElement.nodename
    If Err.Number = 0 Then  'No error accessing XML object.
        Set oInputsNode = oXML.documentElement.SelectSingleNode("inputs")
        If oInputsNode Is Nothing Then
            'no inputs element found on the xml object, so, create the new element
            Set oInputsNode = oXML.createElement("inputs")
            Set oInputsNode = oXML.documentElement.appendChild(oInputsNode)
        End If
        'loop through the dictionary array
        For i = 0 To UBound(asDictionary, 2)
            'append to the inputs element the key and the item fromt the dictionary object
            Set oInputsElement = oXML.createElement(asDictionary(0, i))
            oInputsNode.appendChild oInputsElement
            oInputsNode.lastChild.Text = asDictionary(1, i)
            If Err.Number <> 0 Then Exit For
        Next
        Set oInputsElement = Nothing
        Set oInputsNode = Nothing
    End If
    lErrNumber = Err.Number
    If lErrNumber <> NO_ERR Then
                'sErrDescription = asDescriptors(274) & " " & Err.Description 'Descriptor: MicroStrategy Server error:
        Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "CommonLib.asp", "AddInputsToXML", "", "Error filling XML with Inputs for xsl transformation", LogLevelError)
    End If

    AddInputsToXML = lErrNumber
End Function

Function DateTimeToString(vDateTime)
'*******************************************************************************
'Purpose: To return the date and time as a unique sortable string
'Outputs: A string with the date and time
'*******************************************************************************
	On Error Resume Next
    Dim n
    Dim sResult
    sResult = ""

    If IsDate(vDateTime) Then
        n = CDate(vDateTime)

        sResult = CStr(Year(n))
        If Len(CStr(Month(n))) = 1 Then
                sResult = sResult + "0" + CStr(Month(n))
        Else
                sResult = sResult + CStr(Month(n))
        End If
        If Len(CStr(Day(n))) = 1 Then
                sResult = sResult + "0" + CStr(Day(n))
        Else
                sResult = sResult + CStr(Day(n))
        End If
        If Len(CStr(Hour(n))) = 1 Then
                sResult = sResult + "0" + CStr(Hour(n))
        Else
                sResult = sResult + CStr(Hour(n))
        End If
        If Len(CStr(Minute(n))) = 1 Then
                sResult = sResult + "0" + CStr(Minute(n))
        Else
                sResult = sResult + CStr(Minute(n))
        End If
        If Len(CStr(Second(n))) = 1 Then
                sResult = sResult + "0" + CStr(Second(n))
        Else
                sResult = sResult + CStr(Second(n))
        End If
    End If

    DateTimeToString = sResult
End Function

Function ValidateEmailAddress(sPhysicalAddress)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim i
	Dim bFound
	Dim bInvalid
	Dim atLocation

	bFound = False
	bInvalid = False

	For i = 0 to Ubound(asReservedChars,1)
		If InStr(sPhysicalAddress, asReservedChars(i)) > 0 Then
			bFound = True
			Exit For
		End If
	Next

	If bFound = True Then
		ValidateEmailAddress = -1
		Exit Function
	End If

	If InStr(sPhysicalAddress, "@") = 0 Then
		bInvalid = True
	Else
		atLocation = InStr(sPhysicalAddress, "@")
		If Instr(atLocation, sPhysicalAddress, ".") = 0 Then
			bInvalid = True
		ElseIf Instr(atLocation, sPhysicalAddress, ".") = atLocation + 1 Then
			bInvalid = True
		ElseIf Right(sPhysicalAddress, 1) = "." Then
			bInvalid = True
		ElseIf Left(sPhysicalAddress, 1) = "@" Then
			bInvalid = True
		End If
	End If

	If bInvalid = True Then ValidateEmailAddress = -1
End Function

Function ValidateNumberAddress(sPhysicalAddress)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
    Dim bIsNumber
    Dim sValidChars
    Dim sCurrentChar
    Dim i

    bIsNumber = True
    sValidChars = S_VALID_CHARS_NUM_ADDRESS

    For i=1 to Len(sPhysicalAddress)
        sCurrentChar = Mid(sPhysicalAddress, i, 1)
        If Instr(1, sValidChars, sCurrentChar) = 0 Then
            bIsNumber = false
            Exit For
        End If
    Next

	If bIsNumber = False Then ValidateNumberAddress = -1
End Function

Function RenderLocaleChoices(sGetLocalesForSiteXML, sCurrentLocale, sCurrentLang, sLanguageID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim oOutputDOM
    Dim oLocales
    Dim oLocale
    Dim lErrNumber
    Dim sSelectedLocale
    Dim sBrowserLang
    Dim oBrowserLangLocale
    Dim oEnglishLangLocale
    Dim oCurrentLangLocale
    Dim bFirst

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetLocalesForSiteXML, oOutputDOM)

	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, PAGE_NAME, PROCEDURE_NAME, "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
        Set oLocales = oOutputDOM.selectNodes("/mi/in/oi[@tp='" & TYPE_LOCALE & "' $and$ @plid!='']")

        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "RenderLocaleChoices", "", "Error retrieving oi nodes", LogLevelError)
		End If
	End If

	sSelectedLocale = ""

    If Len(sCurrentLocale) > 0 Then
		sSelectedLocale = sCurrentLocale
    ElseIf Len(sCurrentLang) > 0 Then
		Set oCurrentLangLocale = oOutputDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_LOCALE & "' $and$ @plid='" & CLng(sCurrentLang) & "']")
		sSelectedLocale = oCurrentLangLocale.getAttribute("id")
    ElseIf Len(sLanguageID) > 0 Then
		Set oCurrentLangLocale = oOutputDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_LOCALE & "' $and$ @id='" & sLanguageID & "']")
		sSelectedLocale = oCurrentLangLocale.getAttribute("id")
	Else
		sBrowserLang = TransformLanguage(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE").Item)
		If IsLanguageSupported(sBrowserLang) Then
			Set oBrowserLangLocale = oOutputDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_LOCALE & "' $and$ @plid='" & CLng(sBrowserLang) & "']")
			If Not oBrowserLangLocale Is Nothing Then
				sSelectedLocale = oBrowserLangLocale.getAttribute("id")
			Else
				Set oEnglishLangLocale = oOutputDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_LOCALE & "' $and$ @plid='" & CLng(ENGLISH_US) & "']")
				If Not oEnglishLangLocale Is Nothing Then
					sSelectedLocale = oEnglishLangLocale.getAttribute("id")
				End If
			End If
		End If
	End If

	bFirst = True
	Response.Write "<SELECT CLASS=""pulldownClass"" name=""Locale"">"
	If oLocales.length > 0 Then
	    For Each oLocale in oLocales

			If IsLanguageSupported(oLocale.getAttribute("plid")) Then
				bFirst = False
				Response.Write "<OPTION "
				If (Len(sSelectedLocale)=0 And bFirst) Or _
					Len(sSelectedLocale)>0 And Strcomp(oLocale.getAttribute("id"), sSelectedLocale, vbTextCompare)=0 Then
				    Response.Write "SELECTED "
				End If
				Response.Write "VALUE=""" & oLocale.getAttribute("id") & ";" & oLocale.getAttribute("plid") & """>" & oLocale.getAttribute("n") & "</OPTION>"
			End If
	    Next
	End If
	If bFirst Then
	    Response.Write "<OPTION VALUE=""" & SYSTEM_LOCALE_ID & ";" & CStr(ENGLISH_US) & """>" & asDescriptors(315) & "</OPTION>"	 'Descriptor: Default
	End If
	Response.Write "</SELECT>"

    Set oOutputDOM = Nothing
    Set oLocales = Nothing
    Set oLocale = Nothing
	Set oBrowserLangLocale = Nothing
    Set oEnglishLangLocale = Nothing
    Set oCurrentLangLocale = Nothing

    RenderLocaleChoices = lErrNumber
    Err.Clear
End Function


Function ReadCache(sFileName, sFolderName, sContentXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oFolder
    Dim iSubsCacheSetting

    sContentXML = ""
    lErrNumber = NO_ERR
    iSubsCacheSetting = CInt(Application("iSubsCache"))

    Select Case iSubsCacheSetting
        Case SUBS_CACHE_FILE
            Set oFolder = Server.CreateObject("Scripting.FileSystemObject")
	        If Not (oFolder.FolderExists(Left(APP_CACHE_FOLDER, Len(APP_CACHE_FOLDER) - 1))) Then
	            oFolder.CreateFolder(Left(APP_CACHE_FOLDER, Len(APP_CACHE_FOLDER) - 1))
	        Else
	            If Len(sFolderName) > 0 Then
	                If Not (oFolder.FolderExists(APP_CACHE_FOLDER & sFolderName)) Then
	                    oFolder.CreateFolder(APP_CACHE_FOLDER & sFolderName)
	                Else
                        Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
	                    oCacheDOM.async = False
	                    oCacheDOM.load(APP_CACHE_FOLDER & sFolderName & "\" & sFileName & ".xml")
	                    sContentXML = oCacheDOM.xml
	                End If
	            Else
                    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
	                oCacheDOM.async = False
	                oCacheDOM.load(APP_CACHE_FOLDER & sFileName & ".xml")
	                sContentXML = oCacheDOM.xml
	            End If
	        End If
        Case Else
            sContentXML = Session(sFileName)
    End Select

    Set oCacheDOM = Nothing
    Set oFolder = Nothing

    ReadCache = lErrNumber
    Err.Clear
End Function

Function DeleteCache(sFileName, sFolderName)
'********************************************************
'*Purpose: Deletes XML cache from file system or session variable
'*         sFolderName & sFileName not blank: deletes sFileName from sFolderName
'*         sFileName not blank: deletes sFileName from APP_CACHE_FOLDER
'*         sFolderName not blank: deletes all files in sFolderName
'*Inputs: sFileName, sFolderName
'*Outputs: None
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "DeleteCache"
    Dim lErrNumber
    Dim oFolder
    Dim iSubsCacheSetting

    iSubsCacheSetting = CInt(Application("iSubsCache"))

    Select Case iSubsCacheSetting
        Case SUBS_CACHE_FILE
            Set oFolder = Server.CreateObject("Scripting.FileSystemObject")
            If Err.number <> NO_ERR Then
                lErrNumber = Err.number
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error creating Scripting.FileSystemObject", LogLevelError)
            Else
                If Len(sFolderName) > 0 Then
                    If Len(sFileName) > 0 Then
                        oFolder.DeleteFile APP_CACHE_FOLDER & sFolderName & "\" & sFileName & ".xml"
                    Else
                        oFolder.DeleteFile APP_CACHE_FOLDER & sFolderName & "\*.*"
                    End If
                Else
                    If Len(sFileName) > 0 Then
                        oFolder.DeleteFile APP_CACHE_FOLDER & sFileName & ".xml"
                    End If
                End If

                If Err.number <> NO_ERR Then
                    lErrNumber = Err.number
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling DeleteFile method", LogLevelError)
                End If
            End If
        Case Else
            If Len(sFileName) > 0 Then
                Session(sFileName) = ""
                Set Session(sFileName) = Nothing
            End If
    End Select

    Set oFolder = Nothing

    DeleteCache = lErrNumber
    Err.Clear
End Function

Function WriteCache(sFileName, sFolderName, sContentXML)
'********************************************************
'*Purpose: Writes cache XML to the file system or to a session variable
'*Inputs:
'*      sFileName: a unique name for the file (usually a GUID)
'*      sFolderName: used to create a sub-folder when using the
'*                   file system method (optional)
'*      sContentXML: the XML data to be stored
'*Outputs: None
'*TO DO: Add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim iSubsCacheSetting

    lErrNumber = NO_ERR
    iSubsCacheSetting = CInt(Application("iSubsCache"))

    Select Case iSubsCacheSetting
        Case SUBS_CACHE_FILE
            Call LoadXMLDOMFromString(aConnectionInfo, sContentXML, oCacheDOM)
            If Len(sFolderName) > 0 Then
                oCacheDOM.save(APP_CACHE_FOLDER & sFolderName & "\" & sFileName & ".xml")
            Else
                oCacheDOM.save(APP_CACHE_FOLDER & sFileName & ".xml")
            End If
        Case Else
            Session(sFileName) = sContentXML
    End Select

    Set oCacheDOM = Nothing

    WriteCache = lErrNumber
    Err.Clear
End Function

Function UserAuthenticate(sServerName, sProjectID, lPort, sUserName, sPwd, sUserID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim sSessionID
	Dim sProjectName
	Dim sOtherInfo
	Dim lErrNumber
	Dim oSession
	Dim sErrDescription
	Dim sUserFullName

	sUserID = ""
    aConnectionInfo(S_SERVER_NAME_CONNECTION) = sServerName
    aConnectionInfo(N_PORT_CONNECTION) = lPort
	Call GetDSSSession(aConnectionInfo, oSession, sErrDescription)
	lErrNumber = MapProjectIDToName(oSession, sProjectID, sProjectName)
	If lErrNumber = NO_ERR Then
		sSessionID = oSession.CreateSession(sUserName, sPwd, ,sProjectName, GetLng())
		If Err.number = 0 Then
			Call oSession.GetUserInfo(sSessionID, sUserID, sUserFullName, sOtherInfo)
			Call oSession.CloseSession(sSessionID)
		Else
			lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErrDescription, CStr(Err.source), "LoginCuLib.asp", "UserAuthenticate", "oSession.CreateSession", "Error oSession.CreateSession", LogLevelError)
		End If
	Else
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErrDescription, CStr(Err.source), "LoginCuLib.asp", "UserAuthenticate", "", "Error MapProjectIDToName", LogLevelTrace)
	End If

	UserAuthenticate = lErrNumber
	Err.Clear
End Function

Function GetDSSSession(aConnectionInfo, oSession, sErrDescription)
'*******************************************************************************
'Purpose: To call Server.CreateObject on MSIXML, set the server name & port to the given parameters
'Inputs:  aConnectionInfo
'Outputs: oSession, sErrDescription, Err.number
'*******************************************************************************
    On Error Resume Next
    Dim lErrNumber

    Set oSession = Server.CreateObject(PROGID_WEBOM)
    If Not IsObject(oSession) Or Err.Number <> 0 Then
        lErrNumber = Err.Number
        sErrDescription = asDescriptors(168) 'Descriptor: The following file has not been registered correctly in the Web Server: M8dssxml.dll. Please ask the Administrator for more details.
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), CStr(Err.source), "CommonLib.asp", "GetDSSSession", "", "Error creating MSIXMLLib object", LogLevelError)
    Else
        oSession.ServerName = aConnectionInfo(S_SERVER_NAME_CONNECTION)
        oSession.Port = CLng(aConnectionInfo(N_PORT_CONNECTION))
        oSession.ApplicationType = APPLICATION_TYPE_PORTAL
        If Err.Number <> 0 Then
			lErrNumber = Err.Number
			sErrDescription = asDescriptors(274) & " " & CStr(Err.Description) 'Descriptor: MicroStrategy Server error:
			Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), CStr(Err.source), "CommonLib.asp", "GetDSSSession", "oSession.ServerName", "Error assigning server name and port", LogLevelError)
        End If
    End If

    GetDSSSession = lErrNumber
    Err.Clear
End Function

Function MapProjectIDToName(oSession, sProjectID, sProjectName)
'******************************************************************************
'Purpose: Map a projectID into ProjectName
'Inputs:  sServerName, sProjectID
'Outputs: sProjectName
'******************************************************************************
    On Error Resume Next
	Dim sProjectsXML
    Dim oProjectsXML
    Dim oProject

    sProjectsXML = ""
    sProjectsXML = oSession.GetProjects
    lErrNumber = Err.Number
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), Err.Description, CStr(Err.source), "CommonLib.asp", "GetProjectsFromServer", "oServerSession.GetProjects", "Error getting projects from server", LogLevelError)
    Else
		Set oProjectsXML = Server.CreateObject("Microsoft.XMLDOM")
		Call oProjectsXML.loadXML(sProjectsXML)
		set oProject = oProjectsXML.selectSingleNode("/mi/srps/sp[@ps = '0' $and$ @pgd= '" & sProjectID & "']")
		if oProject is nothing then
			lErrNumber = ERR_PROJECT_NAME_NOT_EXIST
		else
			sProjectName = oProject.getAttribute("pn")
		end if
    End If

    Set oProjectsXML = nothing
    set oProject = nothing

	MapProjectIDToName = lErrNumber
	Err.Clear
End Function

Function cu_GetProfile(sPreferenceObjectID, sQuestionObjectID, sGetProfileXML)
'********************************************************
'*Purpose:
'*Inputs: sPreferenceObjectID, sQuestionObjectID
'*Outputs: sGetProfileXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetProfile"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

	lErrNumber = co_GetProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sGetProfileXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_GetProfile", LogLevelTrace)
	End If

	cu_GetProfile = lErrNumber
	Err.Clear
End Function


Function SetApplicationlVariables()
'********************************************************
'*Purpose: Checks if the Application variables have been set,
'           if not, set them with the default values as
'           stored in MD
'*Inputs:
'*Outputs:
'********************************************************
Const PROCEDURE_NAME = "SetApplicationlVariables"
Dim lErr
Dim sChannelsXML
Dim sLocalesXML
Dim sApplicationVars
Dim sPortalName

Dim aSiteProperties()
Redim aSiteProperties(MAX_SITE_PROP)

    On Error Resume Next

    lErr = NO_ERR
	sPortalName = GetVirtualDirectoryName()

	Call co_GetSharedPropertyManager(sPortalName,sApplicationVars)

	If sApplicationVars = "FALSE" Then
		Application.Value("VARS_READY") = ""

		lErrNumber = co_SetSharedPropertyManager(sPortalName,"TRUE")
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_SetSharedPropertyManager for Portal:" & sPortalName, LogLevelTrace)
		End If
	End If

    'First check if they are already set, if they're ready nothing to do:
    If InStr(1, Application.Value("VARS_READY"), "CONFIG") = 0 Then
		lErr = GetConfigSettings()
    End If

    If lErr = NO_ERR Then
        If InStr(1, Application.Value("VARS_READY"), "PORTAL") = 0 Then

            'If the site is not properly configured we cannot
            'call
            If checkSiteConfiguration() = CONFIG_OK Then
                If lErr = NO_ERR Then
                    lErr = getSiteProperties(aSiteProperties)
                    If lErr <> 0 Then Call LogErrorXML(aConnectionInfo, CStr(lErr), "", CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling getSiteProperties", LogLevelTrace)
                End If

                'We also load at this point channels and Locales for this site:
                If lErr = NO_ERR Then
                    lErr = co_GetChannelsForSite(aSiteProperties(SITE_PROP_ID), sChannelsXML)
                    If lErr <> 0 Then Call LogErrorXML(aConnectionInfo, CStr(lErr), "", CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_GetChannelsForSite", LogLevelTrace)
                End If

                If lErr = NO_ERR Then
                    lErr = co_GetLocalesForSite(aSiteProperties(SITE_PROP_ID), sLocalesXML)
                    If lErr <> 0 Then Call LogErrorXML(aConnectionInfo, CStr(lErr), "", CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_GetLocalesForSite", LogLevelTrace)
                End If

                If lErr = NO_ERR Then
                    Application.Value("Cache_folder") = aSiteProperties(SITE_PROP_TMP_DIR)
                    Application.Value("iSubsCache")   = aSiteProperties(SITE_PROP_PROMPT_CACHE)
                    Application.Value("Admin_email")  = aSiteProperties(SITE_PROP_EMAIL)
                    Application.Value("Admin_phone")  = aSiteProperties(SITE_PROP_PHONE)
                    Application.Value("Allow_New_users")= aSiteProperties(SITE_PROP_NEW_USERS)

                    Application.Value("Login_Mode")= aSiteProperties(SITE_LOGIN_MODE)
                    Application.Value("Login_Authentication_Server_Name")= aSiteProperties(SITE_AUTHENTICATION_SERVER_NAME)
                    Application.Value("Login_Authentication_Server_Port")= aSiteProperties(SITE_AUTHENTICATION_SERVER_PORT)

                    Application.Value("ELE_PROMPT_BLOCK_COUNT_OPTION")= aSiteProperties(SITE_ELEMENT_PROMPT_BLOCK_COUNT)
                    Application.Value("OBJ_PROMPT_BLOCK_COUNT_OPTION")= aSiteProperties(SITE_OBJECT_PROMPT_BLOCK_COUNT)
					Application.Value("PROMPT_MATCH_CASE_OPTION")= aSiteProperties(SITE_PROMPT_MATCH_CASE)

                    Application.Value("Default_locale") = aSiteProperties(SITE_PROP_NEW_LOCALE)
                    Application.Value("Default_expire") = aSiteProperties(SITE_PROP_NEW_EXPIRE)
                    Application.Value("Default_expire_value") = aSiteProperties(SITE_PROP_EXPIRE_VALUE)
                    If Application.Value("Default_expire") = "0" Then Application.Value("Default_expire_value") = SITE_DEFAULT_NO_EXPIRATION

                    Application.Value("Default_gui_language") = aSiteProperties(SITE_PROP_GUI_LANG)
                    Application.Value("Default_use_dhtml") = aSiteProperties(SITE_PROP_USE_DHTML)
					Application.Value("Default_display_summary") = aSiteProperties(SITE_PROP_SUMMARY_PAGE)

                    Application.Value("Default_slicing_answer") = aSiteProperties(SITE_PROP_DEFAULT_ANSWER)
                    Application.Value("Portal_Device")  = aSiteProperties(SITE_PROP_PORTAL_DEV_ID)
                    Application.Value("Default_Device") = aSiteProperties(SITE_PROP_DEFAULT_DEV_ID)
                    Application.Value("Default_Device_Name") = aSiteProperties(SITE_PROP_DEFAULT_DEV_NAME)
                    If aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION) = "email" Then
                        Application.Value("Device_Validation") = S_DEVICE_VALIDATION_EMAIL
                    ElseIf aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION) = "number" Then
                        Application.Value("Device_Validation") = S_DEVICE_VALIDATION_NUMBER
                    Else
                        Application.Value("Device_Validation") = S_DEVICE_VALIDATION_NONE
                    End If

                    Application.Value("Channels_XML") = sChannelsXML
                    Application.Value("Locales_XML") = sLocalesXML

                    Application.Value("Stream_Attachments")= aSiteProperties(SITE_PROP_STREAM_ATTACHMENTS)

                    Application.Value("VARS_READY") = Application.Value("VARS_READY") & "PORTAL;"

                    Application.Value("Timezone")= aSiteProperties(SITE_PROP_TIMEZONE)
                End If
            End If
        End If
    End If

    SetApplicationlVariables = lErr
    Err.Clear

End Function


Function ResetApplicationVariables()
'********************************************************
'*Purpose: Reset all application variables (except those from configuration)
'           to empty values
'*Inputs:
'*Outputs:
'********************************************************
Const PROCEDURE_NAME = "ResetApplicationVariables"
Dim lErr
Dim sFilePath
Dim oFso
Dim oTxtStream
Dim aPortals
Dim ncount
Dim i
    On Error Resume Next
    lErr = NO_ERR

    'These are Site specific
    Application.Value("Cache_folder") = ""
    Application.Value("Admin_email")  = ""
    Application.Value("Admin_phone")  = ""
    Application.Value("Allow_New_users")= ""
    Application.Value("Default_locale") = ""
    Application.Value("Default_expire") = ""
    Application.Value("Default_expire_value") = ""
    Application.Value("Default_gui_language") = ""
    Application.Value("Default_use_dhtml") = ""
    Application.Value("Portal_Device")  = ""
    Application.Value("Default_Device") = ""
    Application.Value("Default_Device_Name") = ""
    Application.Value("Device_Validation") = ""
    Application.Value("Channels_XML") = ""
    Application.Value("Locales_XML") = ""
    Application.Value("iSubsCache") = SUBS_CACHE_FILE
    Application.Value("Default_slicing_answer") = ""
    Application.Value("Stream_Attachments") = ""
	Application.Value("PROMPT_MATCH_CASE_OPTION") = ""

    'We also need to kill cache files:
	sFilePath = Server.MapPath("./")
	If Right(sFilePath, 5) = "admin" Then
	    sFilePath = Server.MapPath("../") & "\"
	Else
	    sFilePath = sFilePath & "\"
	End If


	Set oFso = Server.CreateObject("Scripting.FileSystemObject")
	If IsObject(oFso) Then
	    oFso.DeleteFile(sFilePath & "deviceTypes_*.xml")
	End If


    'Reset Portal and Config Application variables
    Application.Value("VARS_READY") = ""


    lErr = cu_GetAllPortals(aPortals, nCount)
    If lErr <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_GellAllPortals", LogLevelTrace)
    End If

    For i=0 to (nCount - 1)
		lErrNumber = co_SetSharedPropertyManager(aPortals(i,0),"FALSE")
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_SetSharedPropertyManager for Portal:" & aPortals(i,0), LogLevelTrace)
		End If
	Next

    Set oFso = Nothing

    ResetApplicationVariables = lErr
    Err.Clear

End Function

Function GetConfigSettings()
'********************************************************
'*Purpose:Retrieve global information about this portal, such as
'*   Subscription Engine, MDConnection and Site Id; this helps
'*   to check if this Portal is correctly configured:
'*Inputs:  none
'*Outputs: none
'********************************************************
Dim oAdmin
Dim oDOM
Dim sResults
Dim sConfigName
Dim lErr

    On Error Resume Next
    lErr = NO_ERR

    If lErr = NO_ERR Then
        Set oAdmin = Server.CreateObject(PROGID_ADMIN)
        Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
        oDOM.async = False
        lErr = Err.number
    End If

    'Get the values of the Subscription Engine Location
    If lErr = NO_ERR Then
        sResults = oAdmin.getSubscriptionEngineLocation()
        'sResults = "<mi><oi><prs><pr id='SUBSCRIPTION_ENGINE' v='a_paz'></pr></prs></oi></mi>"

        If oDOM.loadXML(sResults) Then
            If Not oDOM.selectSingleNode("//pr[@id='SUBSCRIPTION_ENGINE']") Is Nothing Then
                Application.Value("SE") = oDOM.selectSingleNode("//pr[@id='SUBSCRIPTION_ENGINE']").getAttribute("v")
            End If
        End If

        lErr = Err.number
    End If

    'Get the values of the MD Connection:
    If lErr = NO_ERR Then
        sResults = oAdmin.getMetadataConnectionProperties()

        If oDOM.loadXML(sResults) Then
            If Not oDOM.selectSingleNode("//pr[@id='GROUP_DSN']") Is Nothing Then
                Application.Value("MD_CONN") = oDOM.selectSingleNode("//pr[@id='GROUP_DSN']").getAttribute("v")
            End If
        End If

        lErr = Err.number
    End If

    'Check if the site_id and the DBConns are ready:
    If lErr = NO_ERR Then
        sConfigName = GetVirtualDirectoryName()
        sResults = oAdmin.getSiteConfigurationProperties(sConfigName)

        If oDOM.loadXML(sResults) Then
            If Not oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".SITE_ID']") Is Nothing Then
                Application.Value("SITE_ID") = oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".SITE_ID']").getAttribute("v")
            End If

            If Not oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".SITE_NAME']") Is Nothing Then
                Application.Value("SITE_NAME") = oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".SITE_NAME']").getAttribute("v")
            End If

            If Not oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".AUREP_CONN']") Is Nothing Then
                Application.Value("AUREP_CONN") = oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".AUREP_CONN']").getAttribute("v")
            End If

            If Not oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".SBREP_CONN']") Is Nothing Then
                Application.Value("SBREP_CONN") = oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".SBREP_CONN']").getAttribute("v")
            End If
        End If

        lErr = Err.number
    End If

    If lErr = NO_ERR Then
		Application.Value("VARS_READY") = Application.Value("VARS_READY") & "CONFIG;"
	End If

    Set oAdmin = Nothing
    Set oDOM = Nothing

    GetConfigSettings = lErr
    Err.Clear

End Function


Function GetVirtualDirectoryName()
'***************************************************************************************
'*Purpose: Parses the Servervariable "URL" and returns the Virtual directory name from it
'*Inputs:  none
'*Outputs: Name of the Virtual Directory
'***************************************************************************************

CONST PROCEDURE_NAME = "GetVirtualDirectoryName"

Dim lErr
Dim sURL
Dim lPosition
Dim sVirtualDirectoryName

    On Error Resume Next
    lErr = NO_ERR
    sURL = Request.ServerVariables("URL")

    lPosition = InStr(2,sURL, "/", vbTextCompare)
    sVirtualDirectoryName = LCase(Mid(sURL, 2, lPosition-2))

    GetVirtualDirectoryName = sVirtualDirectoryName
    Err.Clear

End function


Function GerPortalName(sConfigName)
'********************************************************
'*Purpose: Returns the Portal name of the given configuration
'*Inputs:  sConfigName
'*Outputs: Returns the name, if this configuration has no display name returns the configname
'********************************************************
Dim oAdmin
Dim oDOM
Dim sResults
Dim lErr
Dim sName

    On Error Resume Next
    lErr = NO_ERR

	sName = sConfigName

    If lErr = NO_ERR Then
        Set oAdmin = Server.CreateObject(PROGID_ADMIN)
        Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
        oDOM.async = False
        lErr = Err.number
    End If

    If lErr = NO_ERR Then
        sResults = oAdmin.getSiteConfigurationProperties(sConfigName)

        If oDOM.loadXML(sResults) Then
            If Not oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".DISPLAY_NAME']") Is Nothing Then
                sName = oDOM.selectSingleNode("//pr[@id='site." & sConfigName & ".DISPLAY_NAME']").getAttribute("v")
            End If
        End If

        lErr = Err.number
    End If

    Set oAdmin = Nothing
    Set oDOM = Nothing

    GerPortalName = sName
    Err.Clear

End Function


Function cu_GetPreferenceObjects(asPreferenceObjectID, sGetPreferenceObjectsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetPreferenceObjects"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

	If lErrNumber = NO_ERR Then
	    lErrNumber = co_GetPreferenceObjects(sSessionID, asPreferenceObjectID, sGetPreferenceObjectsXML)
	    If lErrNumber <> NO_ERR Then
	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_GetPreferenceObjects", LogLevelTrace)
	    End If
	End If

	cu_GetPreferenceObjects = lErrNumber
	Err.Clear
End Function

Function cu_GetUserAuthenticationObjects(sGetUserAuthenticationObjectsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
    Const PROCEDURE_NAME = "cu_GetUserAuthenticationObjects"
	Dim lErrNumber
    Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    lErrNumber = co_GetUserAuthenticationObjects(sSessionID, sGetUserAuthenticationObjectsXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_GetUserAuthenticationObjects", LogLevelTrace)
    End If

	cu_GetUserAuthenticationObjects = lErrNumber
	Err.Clear
End Function

Function GetSDKVersion()
    Dim oSystemInfo
    Dim sVersion

	Set oSystemInfo = Server.CreateObject(PROGID_SYSTEM_INFO)
	sVersion = oSystemInfo.getVersion()
	Set oSystemInfo = Nothing

	GetSDKVersion = sVersion
End Function

Function IsFinishEnabled(sCacheXML, bFinishEnabled)
'******************************************************************************
'Purpose: Determine if every QO has an <answer> node
'Inputs:  sCacheXML
'Outputs: bFinishEnabled
'******************************************************************************
    On Error Resume Next
	Dim oCacheDOM
	Dim oCurrQO

	Call GetXMLDOM(aConnectionInfo, oCacheDOM, sErrDescription)
	call oCacheDOM.loadXML(sCacheXML)

	bFinishEnabled = True
	For each oCurrQO in oCacheDOM.selectNodes("/mi/qos/mi/in/oi[@tp='5']")
		If (oCurrQO.selectSingleNode("answer") is nothing) Then
			bFinishEnabled = False
			Exit For
		End If
	Next

	Set oCacheDOM = nothing
	Set oCurrQO = nothing

	IsFinishEnabled = Err.number
	Err.Clear
End Function

Function cu_GetLocalesForSite(sGetLocalesForSiteXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
Const PROCEDURE_NAME = "cu_GetLocalesForSite"
Dim lErrNumber

	On Error Resume Next
	lErrNumber = NO_ERR

	sGetLocalesForSiteXML = Application.Value("Locales_XML")

	cu_GetLocalesForSite = lErrNumber
	Err.Clear
End Function

Function cu_GetLocalesForSiteByCurrentLocale(sGetLocalesForSiteXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
Const PROCEDURE_NAME = "cu_GetLocalesForSiteByCurrentLocale"
Dim lErrNumber

	On Error Resume Next
	lErrNumber = NO_ERR

	lErr = co_GetLocalesForSite(Application.Value("SITE_ID"), sGetLocalesForSiteXML)

	cu_GetLocalesForSiteByCurrentLocale = lErrNumber
	Err.Clear
End Function

Function cu_GetAvailableSubscriptions(asSubscriptionGUIDS, sGetAvailableSubscriptionsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
    Const PROCEDURE_NAME = "cu_GetAvailableSubscriptions"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    lErrNumber = co_GetAvailableSubscriptions(sSessionID, asSubscriptionGUIDS, sGetAvailableSubscriptionsXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_GetAvailableSubscriptions", LogLevelTrace)
    End If

	cu_GetAvailableSubscriptions = lErrNumber
	Err.Clear
End Function

Function GetSubscriptionsArray_Reports(sSubsXML, asSubscriptionGUIDS)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oSubsDOM
    Dim oSubs
    Dim i

    lErrNumber = NO_ERR
    Redim asSubscriptionGUIDS(-1)

    Call GetXMLDOM(aConnectionInfo, oSubsDOM, sErrDescription)
    oSubsDOM.async = False
    Call oSubsDOM.loadXML(sSubsXML)

    Set oSubs = oSubsDOM.selectNodes("/mi/subs/sub[@adid = '" & GetPortalAddress() & "']")

    If oSubs.length > 0 Then
        Redim asSubscriptionGUIDS(oSubs.length - 1)
        For i=0 to (oSubs.length - 1)
            asSubscriptionGUIDS(i) = oSubs.item(i).getAttribute("guid")
        Next
    End If

    Set oSubsDOM = Nothing
    Set oSubs = Nothing

    GetSubscriptionsArray_Reports = lErrNumber
    Err.Clear
End Function

Function PutHiddenInputsForConfirmDelete(oRequest)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim Item
	Dim i

	If oRequest.Count > 0 Then
		For Each Item in oRequest
			If Item <> "formPage" And Item <> "cancelButton" Then
				If oRequest(Item).Count > 0 Then
					For i=1 to oRequest(Item).Count
						Response.Write "<input type=""HIDDEN"" name=""" & CStr(Item) & """ value=""" & Server.HTMLEncode(CStr(oRequest(Item)(i))) & """ />"
					Next
				Else
					Response.Write "<input type=""HIDDEN"" name=""" & CStr(Item) & """ value=""" & Server.HTMLEncode(CStr(oRequest(Item))) & """ />"
				End If
			End If
		Next
	End If

	PutHiddenInputsForConfirmDelete = Err.number
	Err.Clear
End Function

Function ParseRequestForDeleteConfirm(oRequest, sFormPage, sCancelButton, sDeleteType, sServiceID, sFolderID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
    Dim lErrNumber
    lErrNumber = NO_ERR

	sFormPage = ""
	sCancelButton = ""
	sDeleteType = ""
	sServiceID = ""
	sFolderID = ""

	sFormPage = Trim(CStr(oRequest("formPage")))
	sCancelButton = Trim(CStr(oRequest("cancelButton")))
	sDeleteType = Trim(CStr(oRequest("deleteType")))
	sServiceID = Trim(CStr(oRequest("serviceID")))
	sFolderID = Trim(CStr(oRequest("folderID")))

	If Err.number <> NO_ERR Then
	    lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", "ParseRequestForDeleteConfirm", "", "Error setting variables equal to Request variables", LogLevelError)
	End If

	ParseRequestForDeleteConfirm = lErrNumber
	Err.Clear
End Function


Function GetXMLDOM(aConnectionInfo, oXMLDOM, sErrDescription)
'*******************************************************************************
'Purpose: To get an instance of the XMLDOM object
'Inputs:  aConnectionInfo
'Outputs: oXMLDOM, sErrDescription, Err.number
'*******************************************************************************
    On Error Resume Next
    Dim lErrNumber
    Set oXMLDOM = Server.CreateObject("Microsoft.XMLDOM")
    lErrNumber = Err.Number
    If lErrNumber = NO_ERR Then
        oXMLDOM.async = False
    Else
		sErrDescription = asDescriptors(271) 'Descriptor: Error loading XML data.
		Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, CStr(Err.source), "CommonLib.asp", "getXMLDOM", "Server.CreateObject(Microsoft.XMLDOM)", "Error instancing the XML Document Object Model", LogLevelError)
    End If

    GetXMLDOM = lErrNumber
    Err.Clear
End Function

Function LoadXMLDOMFromString(aConnectionInfo, sXMLString, oXMLDOM)
'*******************************************************************************
'Purpose: To get an instance of the XMLDOM object
'Inputs:  aConnectionInfo, sXMLString, oXMLDOM
'Outputs: oXMLDOM
'*******************************************************************************
    On Error Resume Next
    Dim lErrNumber

    Set oXMLDOM = Server.CreateObject("Microsoft.XMLDOM")
    lErrNumber = Err.Number
    If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, lErrNumber, Err.Description, CStr(Err.source), "CommonLib.asp", "LoadXMLDOMFromString", "Server.CreateObject(Microsoft.XMLDOM)", "Error instancing the XML Document Object Model", LogLevelError)
    Else
        oXMLDOM.async = False
		Call oXMLDOM.loadXML(sXMLString)
		lErrNumber = Err.Number
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, lErrNumber, Err.Description, CStr(Err.source), "CommonLib.asp", "LoadXMLDOMFromString", "loadXML()", "Error loading XML", LogLevelError)
		End If
    End If

    LoadXMLDOMFromString = lErrNumber
    Err.Clear
End Function

Function LoadXMLDOMFromFile(aConnectionInfo, sFileName, oXMLDOM)
'*******************************************************************************
'Purpose: To get an instance of the XMLDOM object
'Inputs:  aConnectionInfo, sFileName, oXMLDOM
'Outputs: oXMLDOM
'*******************************************************************************
    On Error Resume Next
    Dim lErrNumber

    Set oXMLDOM = Server.CreateObject("Microsoft.XMLDOM")
    lErrNumber = Err.Number
    If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, lErrNumber, Err.Description, CStr(Err.source), "CommonLib.asp", "LoadXMLDOMFromFile", "Server.CreateObject(Microsoft.XMLDOM)", "Error instancing the XML Document Object Model", LogLevelError)
    Else
        oXMLDOM.async = False
		Call oXMLDOM.load(sFileName)
		lErrNumber = Err.Number
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, lErrNumber, Err.Description, CStr(Err.source), "CommonLib.asp", "LoadXMLDOMFromFile", "load()", "Error loading XML", LogLevelError)
		End If
    End If

    LoadXMLDOMFromFile = lErrNumber
    Err.Clear
End Function

Function cu_GetVersions(aVersionInfo)
    Dim lErrNumber

    aVersionInfo(0) = ASP_VERSION 'ASP Version
    aVersionInfo(1) = GetSDKVersion()      'SDK Version

    cu_GetVersions = lErrNumber
    Err.Clear

End Function


Function GetWorkingSet()
'********************************************************
'*Purpose: Get the working set of the inetinfo process
'*Inputs:  None
'*Outputs: The current working set
'********************************************************
Const PROCEDURE_NAME = "GetWorkingSet"
Dim oMemInfo
Dim lWorkingSet

    On Error Resume Next

    Set oMemInfo = Server.CreateObject("MGMemInfo.MemInfo")
    If Err.number <> 0 Then
        lWorkingSet = 0
        Err.Clear
    Else
        lWorkingSet = oMemInfo.GetWorkingSet("")
    End If


    Set oMemInfo = Nothing

    GetWorkingSet = lWorkingSet

End Function

Function checkSiteConfiguration()
'********************************************************
'*Purpose: Checks if the Portal has a valid Engine Configuration; if it doesn't
'          The pages will not be available.
'          It doesn't return the actual values, just returns an error if they
'          are not properly set.
'*Inputs:  none
'*Outputs: A long value:
'           CONFIG_OK If the Portal is configured.
'           CONFIG_MISSING_X Of each value that is missing.
'********************************************************
    On Error Resume Next
    Dim lErr

    lErr = CONFIG_OK

    'For the Engine Configuration to be valid, the following Application Variables
    'should exist:
    'If Len(CStr(Application.Value("SE"))) = 0         Then lErr = lErr Or CONFIG_MISSING_ENGINE
    If Len(CStr(Application.Value("SITE_ID"))) = 0    Then lErr = lErr Or CONFIG_MISSING_SITE
    If Len(CStr(Application.Value("MD_CONN"))) = 0    Then lErr = lErr Or CONFIG_MISSING_MD
    If Len(CStr(Application.Value("AUREP_CONN"))) = 0 Then lErr = lErr Or CONFIG_MISSING_AUREP
    If Len(CStr(Application.Value("SBREP_CONN"))) = 0 Then lErr = lErr Or CONFIG_MISSING_SBREP

    checkSiteConfiguration = lErr
    Err.Clear

End Function

Function getSiteProperties(aSiteProperties)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "getSiteProperties"
    Dim lErrNumber
    Dim oPropsDOM
    Dim sSiteId
    Dim sSitePropsXML

    lErrNumber = NO_ERR

    'If no SiteID given as parameter, use the current site Id:
    If Len(CStr(aSiteProperties(SITE_PROP_ID))) = 0 Then
        sSiteId = Application.Value("SITE_ID")
    Else
        sSiteId = aSiteProperties(SITE_PROP_ID)
    End If

    lErrNumber = co_GetSiteProperties(sSiteId, sSitePropsXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_GetSiteProperties", LogLevelTrace)
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sSitePropsXML, oPropsDOM)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, Err.source, "CommonLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
        End If
    End If

    If lErrNumber = NO_ERR Then

        Redim aSiteProperties(MAX_SITE_PROP)

        aSiteProperties(SITE_PROP_ID) = sSiteId
        aSiteProperties(SITE_PROP_NAME) = GetPropertyValue(oPropsDOM, "NAME")
        aSiteProperties(SITE_PROP_DESC) = GetPropertyValue(oPropsDOM, "DESC")
        aSiteProperties(SITE_PROP_NEW_USERS) = GetPropertyValue(oPropsDOM, "SITE_ALLOW_NEW_USERS")
        aSiteProperties(SITE_PROP_NEW_LOCALE) = GetPropertyValue(oPropsDOM, "SITE_DEFAULT_LOCALE_ID")
        aSiteProperties(SITE_PROP_NEW_EXPIRE) = GetPropertyValue(oPropsDOM, "SITE_DEFAULT_EXPIRE")
        aSiteProperties(SITE_PROP_EXPIRE_VALUE) = GetPropertyValue(oPropsDOM, "SITE_DEFAULT_EXPIRE_VALUE")
        aSiteProperties(SITE_PROP_GUI_LANG) = GetPropertyValue(oPropsDOM, "SITE_GUI_LANG")
        aSiteProperties(SITE_PROP_USE_DHTML) = GetPropertyValue(oPropsDOM, "SITE_DHTML")
        aSiteProperties(SITE_PROP_TMP_DIR) = GetPropertyValue(oPropsDOM, "SITE_TMP_DIR")
        aSiteProperties(SITE_PROP_PROMPT_CACHE) = GetPropertyValue(oPropsDOM, "SITE_PROMPT_CACHE")
        aSiteProperties(SITE_PROP_SUMMARY_PAGE) = GetPropertyValue(oPropsDOM, "SITE_SUMMARY_PAGE")
        If Len(aSiteProperties(SITE_PROP_SUMMARY_PAGE)) = 0 Then
			aSiteProperties(SITE_PROP_SUMMARY_PAGE) = CStr(SITE_PROPVALUE_SUMMARY_PAGE_WHENMORETHANONEQO)
		End If
        aSiteProperties(SITE_PROP_EMAIL) = GetPropertyValue(oPropsDOM, "SITE_ADMIN_EMAIL")
        aSiteProperties(SITE_PROP_PHONE) = GetPropertyValue(oPropsDOM, "SITE_ADMIN_PHONE")
        aSiteProperties(SITE_PROP_DEFAULT_ANSWER) = GetPropertyValue(oPropsDOM, "SITE_DEFAULT_ANSWER")

        aSiteProperties(SITE_PROP_PORTAL_DEV_NAME) = GetPropertyValue(oPropsDOM, "SITE_PORTAL_DEV_NAME")
        aSiteProperties(SITE_PROP_PORTAL_DEV_ID) = GetPropertyValue(oPropsDOM, "SITE_PORTAL_DEV_ID")
        aSiteProperties(SITE_PROP_PORTAL_FOLDER_ID) = GetPropertyValue(oPropsDOM, "SITE_PORTAL_FOLDER_ID")
        aSiteProperties(SITE_PROP_DEFAULT_DEV_NAME) = GetPropertyValue(oPropsDOM, "SITE_DEFAULT_DEV_NAME")
        aSiteProperties(SITE_PROP_DEFAULT_DEV_ID) = GetPropertyValue(oPropsDOM, "SITE_DEFAULT_DEV_ID")
        aSiteProperties(SITE_PROP_DEFAULT_FOLDER_ID) = GetPropertyValue(oPropsDOM, "SITE_DEFAULT_FOLDER_ID")
        aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION) = GetPropertyValue(oPropsDOM, "SITE_DEFAULT_DEV_VALIDATION")

        aSiteProperties(SITE_LOGIN_MODE) = GetPropertyValue(oPropsDOM, "SITE_LOGIN_MODE")
        aSiteProperties(SITE_AUTHENTICATION_SERVER_NAME) = GetPropertyValue(oPropsDOM, "SITE_AUTHENTICATION_SERVER_NAME")
        aSiteProperties(SITE_AUTHENTICATION_SERVER_PORT) = GetPropertyValue(oPropsDOM, "SITE_AUTHENTICATION_SERVER_PORT")

        aSiteProperties(SITE_ELEMENT_PROMPT_BLOCK_COUNT) = GetPropertyValue(oPropsDOM, "SITE_ELEMENT_PROMPT_BLOCK_COUNT")
        aSiteProperties(SITE_OBJECT_PROMPT_BLOCK_COUNT) = GetPropertyValue(oPropsDOM, "SITE_OBJECT_PROMPT_BLOCK_COUNT")
        aSiteProperties(SITE_PROMPT_MATCH_CASE) = GetPropertyValue(oPropsDOM, "SITE_PROMPT_MATCH_CASE")

        aSiteProperties(SITE_PROP_STREAM_ATTACHMENTS) = GetPropertyValue(oPropsDOM, "SITE_STREAM_ATTACHMENTS")

		aSiteProperties(SITE_IS_DEFAULT) = GetPropertyValue(oPropsDOM, "SITE_IS_DEFAULT")
		aSiteProperties(SITE_PROP_TIMEZONE) = GetPropertyValue(oPropsDOM, "SITE_TIMEZONE")
    End If

    Set oPropsDOM = Nothing

    getSiteProperties = lErrNumber
    Err.Clear
End Function

Function setSiteProperties(aSiteProperties, lFlags)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "setSiteProperties"
    Dim lErrNumber
    Dim sSiteId
    Dim sConfigXML

    lErrNumber = NO_ERR

    If Len(CStr(aSiteProperties(SITE_PROP_ID))) = 0 Then
        sSiteId = Application.Value("SITE_ID")
    Else
        sSiteId = aSiteProperties(SITE_PROP_ID)
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = GenerateSitePropertiesXML(aSiteProperties, lFlags, sConfigXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), Err.source, "CommonLib.asp", PROCEDURE_NAME, "", "Error calling GenerateSitePropertiesXML", LogLevelTrace)
        End If
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_SetSiteProperties(sSiteId, sConfigXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_SetSiteProperties", LogLevelTrace)
        End If
    End If

    setSiteProperties = lErrNumber
    Err.Clear
End Function

Function cu_GetAllPortals(aPortals, nCount)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "cu_GetAllPortals"
    Dim lErrNumber
    Dim oDOM
    Dim sPortalXML
    Dim oPortal
    Dim i

    lErrNumber = NO_ERR

    lErrNumber = co_GetAllPortals(sPortalXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_GetAllPortals", LogLevelTrace)
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sPortalXML, oDOM)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
        End If
    End If

    If lErrNumber = NO_ERR Then
        Set oPortal = oDOM.selectNodes("//oi[@tp='1014']")

        If (Not (oPortal Is Nothing)) Then
            Redim aPortals(oPortal.length, 6)
            For i = 0 to (oPortal.length)
                aPortals(i,0) = oPortal(i).getAttribute("n")
                aPortals(i,1) = oPortal(i).getAttribute("sn")
                aPortals(i,2) = oPortal(i).getAttribute("default")
                aPortals(i,3) = oPortal(i).getAttribute("id")
                aPortals(i,4) = oPortal(i).getAttribute("ac")
                aPortals(i,5) = oPortal(i).getAttribute("sc")
                aPortals(i,6) = GerPortalName(aPortals(i, 0))
            Next
            nCount = UBound(aPortals)
	    Else
            nCount = 0
        End If
    End If

    Set oDOM = Nothing
    Set oPortal = Nothing

    cu_GetAllPortals = lErrNumber
    Err.Clear
End Function

Function CreateSiteLoginModeProperty()
'********************************************************
'*Purpose:  Creates a new site Property for Login Mode,Authentication server name
'*Inputs:   None
'*Outputs:  None
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "CreateSiteLoginModeProperty"
    Dim lErr
    Dim sPropertiesXML
    Dim sSiteID

	lErr = NO_ERR

    sPropertiesXML = "<mi><in><oi tp='tp_site'><prs>"
    sPropertiesXML = sPropertiesXML & "<pr id='SITE_LOGIN_MODE'  v='NC_NORMAL' />"
    sPropertiesXML = sPropertiesXML & "<pr id='SITE_AUTHENTICATION_SERVER_NAME'  v='' />"
    sPropertiesXML = sPropertiesXML & "</prs></oi></in></mi>"

    If Len(CStr(aSiteProperties(SITE_PROP_ID))) = 0 Then
        sSiteId = Application.Value("SITE_ID")
    Else
        sSiteId = aSiteProperties(SITE_PROP_ID)
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CreateSiteProperties(sSiteId, sPropertiesXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateSiteProperties", LogLevelTrace)
        End If
    End If

    CreateSiteLoginModeProperty = lErr
    Err.Clear

End Function

Function CreateSitePromptCountProperty()
'********************************************************
'*Purpose:  Creates a new site Property for element and object prompt counts
'*Inputs:   None
'*Outputs:  None
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "CreateSitePromptCountProperty"
    Dim lErr
    Dim sPropertiesXML
    Dim sSiteID

	lErr = NO_ERR

    sPropertiesXML = "<mi><in><oi tp='tp_site'><prs>"
    sPropertiesXML = sPropertiesXML & "<pr id='SITE_ELEMENT_PROMPT_BLOCK_COUNT'  v='30' />"
    sPropertiesXML = sPropertiesXML & "<pr id='SITE_OBJECT_PROMPT_BLOCK_COUNT'  v='30' />"
    sPropertiesXML = sPropertiesXML & "</prs></oi></in></mi>"

    If Len(CStr(aSiteProperties(SITE_PROP_ID))) = 0 Then
        sSiteId = Application.Value("SITE_ID")
    Else
        sSiteId = aSiteProperties(SITE_PROP_ID)
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CreateSiteProperties(sSiteId, sPropertiesXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateSiteProperties", LogLevelTrace)
        End If
    End If

    CreateSitePromptCountProperty = lErr
    Err.Clear

End Function

Function CreateStreamAttachmentsProperty()
'********************************************************
'*Purpose:  Creates a new site Property for Stream attachments mode
'*Inputs:   None
'*Outputs:  None
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "CreateStreamAttachmentsProperty"
    Dim lErr
    Dim sPropertiesXML
    Dim sSiteID

	lErr = NO_ERR

    sPropertiesXML = "<mi><in><oi tp='tp_site'><prs>"
    sPropertiesXML = sPropertiesXML & "<pr id='SITE_STREAM_ATTACHMENTS'  v='1' />"
    sPropertiesXML = sPropertiesXML & "</prs></oi></in></mi>"

    If Len(CStr(aSiteProperties(SITE_PROP_ID))) = 0 Then
        sSiteId = Application.Value("SITE_ID")
    Else
        sSiteId = aSiteProperties(SITE_PROP_ID)
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CreateSiteProperties(sSiteId, sPropertiesXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateSiteProperties", LogLevelTrace)
        End If
    End If

    CreateStreamAttachmentsProperty = lErr
    Err.Clear

End Function

Function CreatePromptMatchCaseProperty()
'********************************************************
'*Purpose:  Creates a new site Property for Stream attachments mode
'*Inputs:   None
'*Outputs:  None
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "CreatePromptMatchCaseProperty"
    Dim lErr
    Dim sPropertiesXML
    Dim sSiteID

	lErr = NO_ERR

    sPropertiesXML = "<mi><in><oi tp='tp_site'><prs>"
    sPropertiesXML = sPropertiesXML & "<pr id='SITE_PROMPT_MATCH_CASE'  v='1' />"
    sPropertiesXML = sPropertiesXML & "</prs></oi></in></mi>"

    If Len(CStr(aSiteProperties(SITE_PROP_ID))) = 0 Then
        sSiteId = Application.Value("SITE_ID")
    Else
        sSiteId = aSiteProperties(SITE_PROP_ID)
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CreateSiteProperties(sSiteId, sPropertiesXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateSiteProperties", LogLevelTrace)
        End If
    End If

    CreatePromptMatchCaseProperty = lErr
    Err.Clear

End Function

Function CreateIserverPortProperty()
'********************************************************
'*Purpose:  Creates a new site Property for I-server port
'*Inputs:   None
'*Outputs:  None
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "CreateIserverPortProperty"
    Dim lErr
    Dim sPropertiesXML
    Dim sSiteID

	lErr = NO_ERR

    sPropertiesXML = "<mi><in><oi tp='tp_site'><prs>"
	sPropertiesXML = sPropertiesXML & "<pr id='SITE_AUTHENTICATION_SERVER_PORT'  v='0' />"
    sPropertiesXML = sPropertiesXML & "</prs></oi></in></mi>"


    If Len(CStr(aSiteProperties(SITE_PROP_ID))) = 0 Then
        sSiteId = Application.Value("SITE_ID")
    Else
        sSiteId = aSiteProperties(SITE_PROP_ID)
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CreateSiteProperties(sSiteId, sPropertiesXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateSiteProperties", LogLevelTrace)
        End If
    End If

    CreateIserverPortProperty = lErr
    Err.Clear

End Function

Function CreateTimeZoneProperty()
'********************************************************
'*Purpose:  Creates a new site Property for time zone info
'*Inputs:   None
'*Outputs:  None
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "CreateTimeZoneProperty"
    Dim lErr
    Dim sPropertiesXML
    Dim sSiteID

	lErr = NO_ERR

    sPropertiesXML = "<mi><in><oi tp='tp_site'><prs>"
	sPropertiesXML = sPropertiesXML & "<pr id='SITE_TIMEZONE'  v='GMT Standard Time' />"
    sPropertiesXML = sPropertiesXML & "</prs></oi></in></mi>"


    If Len(CStr(aSiteProperties(SITE_PROP_ID))) = 0 Then
        sSiteId = Application.Value("SITE_ID")
    Else
        sSiteId = aSiteProperties(SITE_PROP_ID)
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CreateSiteProperties(sSiteId, sPropertiesXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateSiteProperties", LogLevelTrace)
        End If
    End If

    CreateTimeZoneProperty = lErr
    Err.Clear

End Function


Function GenerateSitePropertiesXML(aSiteProperties, lFlags, sPropertiesXML)
'********************************************************
'*Purpose:  Set the Properties of a site
'*Inputs:   aSiteProperties: A properites array,
'           lFlags indicate which elements of the array
'           have valid information, using the FLAGS_PROP constants
'*Outputs:  sPropertiesXML: The XML generated for the properties specified in lFlags
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "GenerateSitePropertiesXML"
    Dim lErr

    lErr = NO_ERR

    'Create the XML necessary for the call:
    sPropertiesXML = "<mi><in><oi tp='tp_site'><prs>"

    If lFlags And FLAG_PROP_GROUP_NAME Then
        sPropertiesXML = sPropertiesXML & "<pr id='NAME'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_PROP_NAME)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='DESC'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_PROP_DESC)),"'", "&apos;") & "' />"
    End If

    If lFlags And FLAG_PROP_GROUP_OTHER Then
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_ALLOW_NEW_USERS'  v='" & aSiteProperties(SITE_PROP_NEW_USERS) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_DEFAULT_LOCALE_ID'  v='" & aSiteProperties(SITE_PROP_NEW_LOCALE) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_DEFAULT_EXPIRE'  v='" & aSiteProperties(SITE_PROP_NEW_EXPIRE) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_DEFAULT_EXPIRE_VALUE'  v='" & aSiteProperties(SITE_PROP_EXPIRE_VALUE) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_GUI_LANG'  v='" & aSiteProperties(SITE_PROP_GUI_LANG) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_DHTML'  v='" & aSiteProperties(SITE_PROP_USE_DHTML) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_TMP_DIR'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_PROP_TMP_DIR)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_PROMPT_CACHE'  v='" & aSiteProperties(SITE_PROP_PROMPT_CACHE) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_SUMMARY_PAGE'  v='" & aSiteProperties(SITE_PROP_SUMMARY_PAGE) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_ADMIN_EMAIL'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_PROP_EMAIL)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_LOGIN_MODE'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_LOGIN_MODE)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_AUTHENTICATION_SERVER_NAME" & "'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_AUTHENTICATION_SERVER_NAME)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_AUTHENTICATION_SERVER_PORT" & "'  v='" & aSiteProperties(SITE_AUTHENTICATION_SERVER_PORT) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_ELEMENT_PROMPT_BLOCK_COUNT'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_ELEMENT_PROMPT_BLOCK_COUNT)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_OBJECT_PROMPT_BLOCK_COUNT'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_OBJECT_PROMPT_BLOCK_COUNT)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_PROMPT_MATCH_CASE'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_PROMPT_MATCH_CASE)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_ADMIN_PHONE'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_PROP_PHONE)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_STREAM_ATTACHMENTS'  v='" & aSiteProperties(SITE_PROP_STREAM_ATTACHMENTS) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_TIMEZONE'  v='" & aSiteProperties(SITE_PROP_TIMEZONE) & "' />"
    End If

    If lFlags And FLAG_PROP_GROUP_SERVICES Then
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_DEFAULT_ANSWER' v='" & aSiteProperties(SITE_PROP_DEFAULT_ANSWER) & "' />"
    End If

    If lFlags And FLAG_PROP_GROUP_DEVICES Then
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_PORTAL_DEV_ID'  v='" & aSiteProperties(SITE_PROP_PORTAL_DEV_ID) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_PORTAL_DEV_NAME'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_PROP_PORTAL_DEV_NAME)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_PORTAL_FOLDER_ID'  v='" & aSiteProperties(SITE_PROP_PORTAL_FOLDER_ID) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_DEFAULT_DEV_ID'  v='" & aSiteProperties(SITE_PROP_DEFAULT_DEV_ID) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_DEFAULT_DEV_NAME'  v='" & Replace(Server.HTMLEncode(aSiteProperties(SITE_PROP_DEFAULT_DEV_NAME)),"'", "&apos;") & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_DEFAULT_DEV_VALIDATION'  v='" & aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION) & "' />"
        sPropertiesXML = sPropertiesXML & "<pr id='SITE_DEFAULT_FOLDER_ID'  v='" & aSiteProperties(SITE_PROP_DEFAULT_FOLDER_ID) & "' />"
    End If

    sPropertiesXML = sPropertiesXML & "</prs>"


    If lFlags And FLAG_PROP_GROUP_CONN Then
        sPropertiesXML = sPropertiesXML & "<mi><in>"

        sPropertiesXML = sPropertiesXML & "<oi id='" & GetGUID() & "' tp='1003'><prs>"
        sPropertiesXML = sPropertiesXML & " <pr id='GROUP' v='AUREP' />"
        sPropertiesXML = sPropertiesXML & " <pr id='GROUP_PREFIX' v=' ' />"
        sPropertiesXML = sPropertiesXML & " <pr id='GROUP_DSN' v=' ' />"
        sPropertiesXML = sPropertiesXML & "</prs></oi>"

        sPropertiesXML = sPropertiesXML & "<oi id='" & GetGUID() & "' tp='1003'><prs>"
        sPropertiesXML = sPropertiesXML & " <pr id='GROUP' v='SBREP' />"
        sPropertiesXML = sPropertiesXML & " <pr id='GROUP_PREFIX' v=' ' />"
        sPropertiesXML = sPropertiesXML & " <pr id='GROUP_DSN' v=' ' />"
        sPropertiesXML = sPropertiesXML & "</prs></oi>"

        sPropertiesXML = sPropertiesXML & "</in></mi>"
    End If

    sPropertiesXML = sPropertiesXML & "</oi></in></mi>"

    GenerateSitePropertiesXML = lErr
    Err.Clear
End Function

Function GetPropertyValue(oDOM, sPropertyId)
'********************************************************
'*Purpose:  Returns the value of a property (if exists)
'           searching inside the DOM Object
'*Inputs:   oDOM: A valid DOM object (probably from a getSiteProperties call)
'           sPropertyId: The id of the property we're looking for.
'*Outputs:  The property Value (if exists) or ""
'********************************************************
    On Error Resume Next
    Dim oNode
    Dim sValue

    Set oNode = Nothing
    Set oNode = oDom.selectSingleNode(".//pr[@id=""" & sPropertyId & """]")

    'By default Return an Empty Value, if the node exist, return its value:
    sValue = ""
    If Not oNode Is Nothing Then
        sValue = oNode.getAttribute("v")
    End If

    Set oNode = Nothing

    GetPropertyValue = sValue
    Err.Clear
End Function

Function SetRootFolder(sChannelId)
'********************************************************
'*Purpose: Sets the APP_ROOT_FOLDER variable to point to the svcFolderId
'*Inputs:  sChannelId: The current channel
'          sChannelsXML: The XML with all the channels:
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "SetRootFolder"
    Dim lErr
    Dim sErr
    Dim sChannelsXML
    Dim oDOM
    Dim oChannel

    lErr = NO_ERR

    sChannelsXML = CStr(Application.Value("Channels_XML"))
    If Len(sChannelsXML) > 0 Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sChannelsXML, oDOM)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErr), "", "", "CommonLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
        Else
            Set oChannel = oDOM.selectSingleNode("//oi[@id='" & sChannelId & "']")

            If oChannel Is Nothing Then
                lErr = ERR_XML_LOAD_FAILED
                Call LogErrorXML(aConnectionInfo, CStr(lErr), "", "", "CommonLib.asp", PROCEDURE_NAME, "", "Could not find a channel with Id: " & sChannelId, LogLevelError)
            Else
                Application.Value("Root_Folder") = oChannel.selectSingleNode("prs/pr[@id='serviceFolderID']").getAttribute("v")
            End If
        End If
    End If

    Set oDOM = Nothing
    Set oChannel = Nothing

    SetRootFolder = lErr
    Err.Clear
End Function

Function cu_CreateTransmissionProperties(sTransPropsID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_CreateTransmissionProperties"
	Dim lErrNumber
	Dim sSiteID
	Dim asTransPropID(0)
	Dim asTransProperty(0)
	Dim sCreateTransmissionPropertiesXML

	lErrNumber = NO_ERR
	sSiteID = SITE_ID
	If Len(sTransPropsID) = 0 Then
	    sTransPropsID = GetGUID()
	End If
	asTransPropID(0) = sTransPropsID
    asTransProperty(0) = ""

	lErrNumber = co_CreateTransmissionProperties(sSiteID, asTransPropID, asTransProperty, sCreateTransmissionPropertiesXML)
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error while calling co_CreateTransmissionProperties", LogLevelTrace)
	End If

	cu_CreateTransmissionProperties = lErrNumber
	Err.Clear
End Function

Function cu_DeleteTransmissionProperties(asTransPropID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_DeleteTransmissionProperties"
	Dim lErrNumber
	Dim sSiteID
	Dim sDeleteTransmissionPropertiesXML

	lErrNumber = NO_ERR
	sSiteID = SITE_ID

	lErrNumber = co_DeleteTransmissionProperties(sSiteID, asTransPropID, sDeleteTransmissionPropertiesXML)
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error while calling co_DeleteTransmissionProperties", LogLevelTrace)
	End If

	cu_DeleteTransmissionProperties = lErrNumber
	Err.Clear
End Function

Function CloseCastorSession(aConnectionInfo)
'******************************************************************************
'Purpose: Close the current castor session
'Inputs:
'Outputs:
'******************************************************************************
    On Error Resume Next
    Dim oSession

    Call GetDSSSession(aConnectionInfo, oSession, sErrDescription)
    Call oSession.CloseSession(aConnectionInfo(S_TOKEN_CONNECTION))

    CloseCastorSession = Err.Number
    Err.Clear
End Function

Function Decrypt(sEncrypted)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "Decrypt"
	Dim lErrNumber
	Dim oDecryptObj

	'Set oDecryptObj = Server.CreateObject("MSTRCOMCrypto.COMCrypto.1")
	Set oDecryptObj = Server.CreateObject("MSTRHydraCOMCrypto.COMCrypto2.1")
	Decrypt = oDecryptObj.Decrypt(sEncrypted)
	lErrNumber = Err.number
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error while calling Decrypt", LogLevelTrace)
	End If

	Set oDecryptObj = Nothing
	Err.Clear
End Function

Function Encrypt(sPlain)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "Encrypt"
	Dim lErrNumber
	Dim oEncryptObj

	'Set oEncryptObj = Server.CreateObject("MSTRCOMCrypto.COMCrypto.1")
	Set oEncryptObj = Server.CreateObject("MSTRHydraCOMCrypto.COMCrypto2.1")
	Encrypt = oEncryptObj.Encrypt(sPlain)
	lErrNumber = Err.number
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error while calling Encrypt", LogLevelTrace)
	End If

	Set oEncryptObj = Nothing
	Err.Clear
End Function


Function ParseAuthenticationObject(sUserAuth, sUserName, sPwd, sUserID)
'********************************************************
'*Purpose: parse string AuthUserName="administrator" AuthUserPwd="xxx" AuthUserID="123"
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "ParseAuthenticationObject"
	Dim iBeginPos
	Dim iEndPos

	iBeginPos = InStr(1, sUserAuth, "AuthUserName=""", vbTextCompare) + Len("AuthUserName=""")
	iEndPos = InStr(iBeginPos, sUserAuth, """", vbTextCompare)
	sUserName = Mid(sUserAuth, iBeginPos, iEndPos - iBeginPos)

	iBeginPos = InStr(1, sUserAuth, "AuthUserPwd=""", vbTextCompare) + Len("AuthUserPwd=""")
	iEndPos = InStr(iBeginPos, sUserAuth, """", vbTextCompare)
	sPwd = Decrypt(Mid(sUserAuth, iBeginPos, iEndPos - iBeginPos))

	iBeginPos = InStr(1, sUserAuth, "AuthUserID=""", vbTextCompare) + Len("AuthUserID=""")
	iEndPos = InStr(iBeginPos, sUserAuth, """", vbTextCompare)
	sUserID = Mid(sUserAuth, iBeginPos, iEndPos - iBeginPos)

	If Err.number <> 0 Then
		lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonLib.asp", PROCEDURE_NAME, "", "Error parsing", LogLevelError)
	End If

	ParseAuthenticationObject = Err.number
	Err.Clear
End Function

Function BuildAuthenticationObject(sUserName, sPwd, sUserID, sUserAuth)
'********************************************************
'*Purpose: build string AuthUserName="administrator" AuthUserPwd="xxx" AuthUserID="123"
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "ParseAuthenticationObject"

	sUserAuth = "AuthUserName=""" & sUserName & """ AuthUserPwd=""" & Encrypt(sPwd) & """ AuthUserID=""" & sUserID & """"

	BuildAuthenticationObject = Err.number
	Err.Clear
End Function

Function IsLanguageSupported(sLanguage)
'********************************************************
'*Purpose: check if the language is supported by Hydra
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "IsLanguageSupported"

	If InStr(1, "1031,1033,3082,1036,1040,1041,1042,1046,1053,2052,2057", sLanguage, 0) > 0 Then
		IsLanguageSupported = True
	Else
		IsLanguageSupported = False
	End If

	Err.Clear

End Function

Function SetIServerCookies(sUserName,sPassword,IServerUser,sNTUser)
	Response.Cookies("ISERVER_USERNAME") = sUserName
	'Response.Cookies("ISERVER_USERNAME").expires = #1/1/1980#

	Response.Cookies("ISERVER_PASSWORD") = Encrypt(sPassword)
	'Response.Cookies("ISERVER_PASSWORD").expires = #1/1/1980#

	Response.Cookies("ISERVERUSER") = IServerUser
	'Response.Cookies("ISERVERUSER").expires = #1/1/1980#

	Response.Cookies("ISERVERNTUSER") = sNTUser
	'Response.Cookies("ISERVERNTUSER").expires = #1/1/1980#

End Function

Function GetIServerCookies(sUserName,sPassword,IServerUser)
	sUserName = Request.Cookies("ISERVER_USERNAME")
	sPassword = Decrypt(Request.Cookies("ISERVER_PASSWORD"))
	If not Request.Cookies("ISERVERUSER") is nothing and Request.Cookies("ISERVERUSER") <> "" Then
		IServerUser = CBool(Request.Cookies("ISERVERUSER"))
	Else
		IServerUser = false
	End If
End Function

'Functions from Web to be used in the Prompt Code

Function DecodeStr(strToDecode)
'******************************************************************************
' Purpose:	To decode a string. (For special characters like the Euro)
' Inputs:	strToDecode
' Outputs:  A decoded string
'******************************************************************************
	On Error Resume Next
	Dim sTemp
	sTemp = replace(strToDecode,"&#8364;","€")
	DecodeStr = sTemp
	Err.Clear
End Function

Function DecodeXML(sSource)
'*******************************************************************************
'Purpose: To restore " and & in xml strings that have been passed through forms
'Inputs:  sSource
'Outputs: The decoded string
'*******************************************************************************
	On Error Resume Next
    DecodeXML = Replace(Replace(Replace(Replace(sSource, "&#38;", "&"), "&#34;", """"), "%3C", "<"), "%3E", ">")
    Err.Clear
End Function

Function GetServerLanguage(aConnectionInfo, iSource)
'*******************************************************************************
'Purpose: Retrieves the current language
'Inputs:  aConnectionInfo, iSource
'Outputs: GetLanguage
'*******************************************************************************
	On Error Resume Next
'	Dim aLang(2)
'	Dim sLanguage
'	aLang(0) = "ServerLng"
'	aLang(1) = ""
'	sLanguage = CStr(ReadFromSource(aConnectionInfo, iSource, aLang))
'	If (Len(sLanguage) = 0) Or (StrComp(sLanguage, "0000", vbBinaryCompare) = 0) Or (InStr(1, Server.MapPath("."), "Admin", vbTextCompare)) Then
'		sLanguage = GetLanguage(aConnectionInfo, iSource)
'	End If
'	GetServerLanguage = sLanguage

	GetServerLanguage = GetLng()
	Err.Clear
End Function

Function GetDateFormatForLocale(aConnectionInfo, sErrDescription)
'*********************************************************************************************
'Purpose:	Get the Date Format for the locale - used in the calendar picker
'Input:
'Output:
'*********************************************************************************************
	On Error Resume Next
	Dim oLocaleInfo
	Dim sLangID
	Dim lErrNumber

	sLangID = GetServerLanguage(aConnectionInfo, Application.Value("iSourcePerm"))
	Set oLocaleInfo = Server.CreateObject("MBJUTBRI.LocaleInfo")
	lErrNumber = Err.number
    If lErrNumber <> NO_ERR Then
		sErrDescription = Err.description
		Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "CommonLib.asp", "GetDecimalSeparator", "", "Error after calling Server.CreateObject(""MBJUTBRI.LocaleInfo"")", LogLevelError)
		Exit Function
    End IF
	oLocaleInfo.LocaleID = Clng(sLangID)
	GetDateFormatForLocale = Ucase(Cstr(oLocaleInfo.ShortDateFormatString))
	Err.Clear
End Function

Function CleanXML(sSource)
'*******************************************************************************
'Purpose: To replace " and & in xml strings so they can safely be passed through forms
'Inputs:  sSource
'Outputs: The new string
'*******************************************************************************
	On Error Resume Next
    CleanXML = Replace(Replace(sSource, "&", "&#38;"), """", "&#34;")
    Err.Clear
End Function


Function SplitRequest(sRequest)
'*******************************************************************************
'Purpose:	This function is for the prompts code to handle the multiple elements in the same request Variable
'Inputs:	None
'Outputs:	Returns the array of inputs.
'*******************************************************************************
	On Error Resume Next
	Dim aElements
	Dim lCount
	Dim vItem
	Dim i

	lCount = sRequest.Count

	Redim aElements(lCount - 1)
	i = 0

	For each vItem in sRequest
		aElements(i) = vItem
		i = i+1
	Next

	SplitRequest = aElements

	Err.Clear
End Function

Function GetGeneralParasinURL(oRequest)
'******************************************************************************
'Purpose: To put basic parameters in oRequest in sURL such as:
'         MsgID, page, ReportID, view, Server, Uid, Project, Port
'Inputs:  oRequest
'Outputs: sURL, Err.Number
'******************************************************************************
    On Error Resume Next
    Dim sPondPos
    Dim sAmpPos
    Dim sURL

    sURL = "MsgID" & "=" & CStr(oRequest("MsgID")) & "&ReportID" & "=" & CStr(oRequest("ReportID"))
    sURL = sURL & "&FilterID" & "=" & CStr(oRequest("FilterID")) & "&TemplateID" & "=" & CStr(oRequest("TemplateID"))
    sURL = sURL & "&DocumentID" & "=" & CStr(oRequest("DocumentID")) & "&view" & "=" & CStr(oRequest("view"))
    sURL = sURL & "&Page" & "=" & CStr(oRequest("Page")) & "&doc" & "=" & CStr(oRequest("doc"))
    sURL = sURL & "&Server" & "=" & Server.URLEncode(CStr(oRequest("Server"))) & "&Project" & "=" & Server.URLEncode(CStr(oRequest("Project")))
    sURL = sURL & "&Port" & "=" & CStr(oRequest("Port")) & "&Uid" & "=" & Server.URLEncode(CStr(oRequest("Uid"))) & "&UMode" & "=" & CStr(oRequest("UMode"))

    sPondPos = InStr(1, sURL, "#", vbBinaryCompare)
    if sPondPos>0 then
		sAmpPos = InStr(sPondPos, sURL, "&", vbBinaryCompare)
		if sAmpPos>0 then
			sURL = Mid(sURL, 1, sPondPos-1) & Mid(sURL, sAmpPos)
		else
			sURL = Left(sURL, sPondPos-1)
		end if
	End if

    GetGeneralParasinURL = sURL
    Err.Clear
End Function

Function CleanErrorMessage(sErrDescription)
'*******************************************************************************
'Purpose: To replace <,>,' etc in the error description so that when we redirect we do not have such characters in the URL
'Inputs:  sErrDescription
'Outputs: The new string
'*******************************************************************************
	On Error Resume Next
		CleanErrorMessage = Replace(Replace(Replace(sErrDescription,"<","&#60;"),"'"," "),">","&#62;")
	Err.Clear
End Function

Function PopulateClientDescriptors(asDescriptors, bScriptTags)
	On Error Resume Next
	If bScriptTags Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">" & vbNewline
	End If
	Response.Write "var asDescriptors = new Array();" & vbNewline
	Response.Write "asDescriptors[0] = '" & asDescriptors(2377) & "';"  'Descriptor : Please enter a number in the following field
	Response.write "asDescriptors[1] = '" & asDescriptors(1956) & "';" 'January
	Response.write "asDescriptors[2] = '" & asDescriptors(1957) & "';" 'February
	Response.write "asDescriptors[3] = '" & asDescriptors(1958) & "';" 'March
	Response.write "asDescriptors[4] = '" & asDescriptors(1959) & "';" 'April
	Response.write "asDescriptors[5] = '" & asDescriptors(1960) & "';" 'May
	Response.write "asDescriptors[6] = '" & asDescriptors(1961) & "';" 'June
	Response.write "asDescriptors[7] = '" & asDescriptors(1962) & "';" 'July
	Response.write "asDescriptors[8] = '" & asDescriptors(1963) & "';" 'August
	Response.write "asDescriptors[9] = '" & asDescriptors(1964) & "';" 'September
	Response.write "asDescriptors[10] = '" & asDescriptors(1965) & "';" 'October
	Response.write "asDescriptors[11] = '" & asDescriptors(1966) & "';" 'November
	Response.write "asDescriptors[12] = '" & asDescriptors(1967) & "';" 'December
	Response.write "asDescriptors[13] = '" & asDescriptors(1968) & "';" 'Sunday
	Response.write "asDescriptors[14] = '" & asDescriptors(1969) & "';" 'Monday
	Response.write "asDescriptors[15] = '" & asDescriptors(1970) & "';" 'Tuesday
	Response.write "asDescriptors[16] = '" & asDescriptors(1971) & "';" 'Wednesday
	Response.write "asDescriptors[17] = '" & asDescriptors(1972) & "';" 'Thursday
	Response.write "asDescriptors[18] = '" & asDescriptors(1973) & "';" 'Friday
	Response.write "asDescriptors[19] = '" & asDescriptors(1974) & "';" 'Saturday
	Response.write "asDescriptors[20] = '" & asDescriptors(2034) & "';" 'Subtotal Names
	Response.write "asDescriptors[21] = '" & asDescriptors(2035) & "';" 'Subtotal Values
	Response.Write "asDescriptors[22] = '" & asDescriptors(2376) & "';" 'Descriptor: Please wait...
	Response.Write "asDescriptors[23] = '" & asDescriptors(2382) & "';" 'Descriptor: Please enter an integer greater than ### and less than ## in the following field
	Response.Write "asDescriptors[24] = '" & asDescriptors(2647) & "';" ' Descriptor: Width: ##px

	If bScriptTags Then
		Response.Write "</SCRIPT>" & vbNewLine
	End If
	Err.Clear
End Function

Function DisplayError1(sTitle, sErrDescription)
'*******************************************************************************
'Purpose: To display the error message
'Inputs:  sTitle, sErrDescription
'Outputs: HTML code with the error
'*******************************************************************************
        On Error Resume Next
        Dim sError

        sError = "<BR /><TABLE CELLSPACING=""0"" CELLPADDING=""0"" BORDER=""0"" WIDTH=""90%""><TR><TD ALIGN=""LEFT"">"
        sError = sError & "<HR SIZE=""1"" />"
        sError = sError & "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """ COLOR=""#CC0000""><B>"
        sError = sError & sTitle
        sError = sError & "</B></FONT>"
        If Len(sTitle) > 0 Then sError = sError & "<BR /><BR />"
        sError = sError & "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"
        sError = sError & sErrDescription & "</FONT>"
        sError = sError & "<HR SIZE=""1"" /></TD></TR></TABLE>"

        DisplayError = sError
        Err.Clear
End Function

Function  IsValidDate(sDateStr)
'************************************************************************************************
' Purpose: To determine whether to show or the format toolbar or not
' Inputs: aConnectionInfo, aReportInfo
' Outputs: boolean
'*************************************************************
On Error Resume Next

	Dim sLocaleDateStr

	Dim aValidDateSeparator(2)
	Dim lMaxSeparators
	Dim i,j

	IsValidDate = False

	If IsDate(sDateStr) Then
		IsValidDate = True
		Exit Function
	Else

		aValidDateSeparator(0) = "/"
		aValidDateSeparator(1) = "."
		aValidDateSeparator(2) = "-"

		lMaxSeparators = Ubound(aValidDateSeparator)

		For i = 0 To lMaxSeparators
			If Instr(1,sDateStr,aValidDateSeparator(i),vbTextCompare) > 0 Then
				For j = 0 To lMaxSeparators
					If i <> j Then
						sLocaleDateStr = Replace(sDateStr,aValidDateSeparator(i),aValidDateSeparator(j))
						If IsDate(sLocaleDateStr) Then
							IsValidDate = True
							Exit Function
						End If
					End If
				Next
			End If
		Next
	End If
    Err.Clear
End Function


Function GetVersion()
	Dim oGUILib
	Set oGUILib = Server.CreateObject(PROGID_NCS_GUI_LIB)

	GetVersion = oGUILib.GetVersion()

End Function

Function GetDefaultTimeZone()
	Dim oTZ
	Set oTZ = Server.CreateObject(PROGID_NCS_COM_TIME_ZONE)
	GetDefaultTimeZone = oTZ.Name
End Function

Function PopulateTimeZones (aTimeZones)

	Dim oTZLib
	Dim sDefaultTimeZone 'not used
	Dim aTimeZoneDisplayNames()
	Dim aTimeZoneStdNames()
	Dim lTimezoneCount
	Dim i

	Set oTZLib = Server.CreateObject(PROGID_TIMEZONE_LIB)

	Call oTZLib.GetTimeZoneDisplayStringsV(aTimeZoneDisplayNames, aTimeZoneStdNames, sDefaultTimeZone)

	lTimezoneCount = UBound(aTimeZoneDisplayNames) - 1

	Redim aTimeZones(lTimezoneCount, 1)

	For i = 0 To lTimezoneCount
		aTimeZones(i, 0) = aTimeZoneStdNames(i+1)
		aTimeZones(i, 1) = aTimeZoneDisplayNames(i+1)
	Next

End Function

Function ConvertLocalTimeToUTC(vTime)
	Dim oTZ
	Set oTZ = Server.CreateObject(PROGID_NCS_COM_TIME_ZONE)
	ConvertLocalTimeToUTC = oTZ.GMTTime(vTime)
End Function

%>