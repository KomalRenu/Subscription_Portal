<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%

Public Const NO_ERR = 0
Public Const ERR_UNEXPECTED = 823

Public Const ERR_INVALID_SESSION = -2147206924
Public Const ERR_NO_ACCESS_RIGHT = -2147214570
Public Const ERR_USER_NOT_FOUND = -2147209051
Public Const ERR_API_SERVER_NOT_FOUND = -2147206923
Public Const ERR_API_REPORT_EXPIRED = -2147468986
Public Const ERR_FULL_INBOX =  -2147206853 '&H8004393B
Public Const ERR_MAX_JOBS_PER_USER_EXCEEDED = -2147209053 '&H800430A3
Public Const ERR_MAX_JOBS_PER_PROJECT_EXCEEDED = -2147209164 '&H80043034
Public Const ERR_MSG_NOT_FOUND_IN_INBOX = -2147468986 '&H80003946
Public Const ERR_INBOX_FULL = -2147205473
Public Const ERR_UNABLE_CONNECT_TRANSACTOR = -2147220992

Public Const ERR_USER_FOLDER_NOT_FOUND = -2147468987
Public Const ERR_API_PROJECT_IDLE = -2147209192
Public Const ERR_API_NO_OBJECT_ACCESS = -2147214581
Public Const ERR_API_NO_PROJECT_ACCESS = -2147214578
Public Const ERR_API_NO_WRITE_ACCESS = -2147214579
Public Const ERR_API_SERVER_DOWN = -2147207419

Private Const ERR_NO_HTML_IN_SERVER = -2147469003
Private Const ERR_UNABLE_TO_OPEN_HTML_IN_SERVER = -2147468637 '&H80003AA3
Private Const ERR_NO_XSL_IN_SERVER = -2147468636 '&H80003AA4
Private Const ERR_XSL_AND_XML_IN_SERVER = -2147468475 '&H80003B45
Private Const ERR_NO_DISK_SPACE_IN_SERVER = -2147468890 '&H800039A6

Private Const ERR_TOO_MANY_ROWS_RETURNED = -2147212544 '&H80042300
Public Const ERR_API_CACHE_NOT_FOUND = -2147467618 '&H80003E9E
Public Const ERR_API_INVALID_STATE_ID = -2147468474 '&H80003B46
Public Const ERR_API_REQUEST_TIMED_OUT = -2147206497 '&H80043A9F

'Castor API errors
Public Const API_ERR_PROJECT_OFFLINE = -2147209183
Public Const API_ERR_LOGIN_PASSWORD_INVALID = -2147216959
Public Const API_ERR_SERVER_NOT_FOUND = -2147206923
Public Const API_ERR_USER_PRIVILEGES = -2147216961
Public Const AUTHEN_E_LOGIN_FAILED_NEW_PASSWORD_REQD = -2147216960
Public Const AUTHEN_E_ACCOUNT_DISABLED = -2147216965
Public Const AUTHEN_E_LOGIN_FAIL_EXPIRED_PWD = -2147216963
Public Const API_ERR_NT_NOT_LINKED = -2147216953

Private Const ERR_NO_PROMPT_ANSWER_DOCUMENT = -2147271032


'Intillegent Server Logging Error codes
Public Const IS_ERR_LOGIN_ERROR = -2147216959
Public Const IS_SERVER_NOT_FOUND = -2147207418

'NarrowCast Server Portal Error codes
Public Const ERR_DOC_BODY_NOT_FOUND = "C0045191"
Public Const ERR_LOGIN_ERROR = "C0045210"
Public Const ERR_USER_INACTIVE = "C0045218"
Public Const ERR_PRIMARY_KEY_VIOLATION = "C00452D8"
Public Const ERR_SESSION_TIMEOUT = "C0045451"
Public Const ERR_EXCEPTION_THROWN = "C0045007"
Public Const ERR_NO_TABLES_EXIST = "C0045117"
Public Const ERR_WRONG_DBALIAS_DEFINITION = "C0045006"
Public Const ERR_WRONG_DBALIAS_USER_PASS ="C"
Public Const ERR_WRONG_DBALIAS_NAME = "C00452D6"
Public Const ERR_WRONG_TABLE_VERSION = "C0045123"
Public Const ERR_PROPERTY_NOT_DEFINED = "C0045112"
Public Const ERR_USER_NOT_EXIST = "C0045216"
Public Const ERR_VD_ALREADY_EXIST = "C0045124"
Public Const ERR_FOLDER_NOT_FOUND = "C0045560"
Public Const ERR_NO_ACTIVE_FOLDER = "C0045DF0"


Public Const URL_MISSING_PARAMETER = -2
Public Const ERR_ADDR_OPERATION = -5
Public Const ERR_SUBS_OPERATION = -8
Public Const ERR_XML_LOAD_FAILED = -11
Public Const ERR_MISSING_PORTAL_ADDRESS = -12
Public Const ERR_RETRIEVING_RESULTS = -13
Public Const ERR_INVALID_DEFAULT_OBJECT = -14
Public Const ERR_INVALID_DEFAULT_CHANNEL = -15
Public Const ERR_CACHE_CONTENT = -16
Public Const ERR_USERDEFAULT_NOTEXIST = -17
Public Const ERR_QUESTION_IN_SERVICE_DEF = -18
Public Const ERR_QUESTION_ALREADY_USED = -19
Public Const ERR_EMPTY_GUID = -19
Public Const ERR_INACTIVE_FOLDER_ANCESTOR = -20

Public Const ERR_LOGIN_BLANKS = 1
Public Const ERR_HINT_BLANK = 2
Public Const ERR_CONFIRM_PASSWORD = 4
Public Const ERR_DEFAULT_ADDRESS_INVALID = 8
Public Const ERR_ISLOGIN_BLANK = 16
Public Const ERR_ISLOGIN_ERROR = 32
Public Const ERR_ADDRESS_BLANKS = 1
Public Const ERR_ADDR_NAME_INVALID = 2
Public Const ERR_EMAIL_ADDR_INVALID = 4
Public Const ERR_NUMBER_ADDR_INVALID = 8

Public Const ERR_INVALID_NAME = 1

'*** CORE FUNCTIONS ***'
Private Const ERR_OBJECT_DOES_NOT_EXIST = -2147216373

'*** Error logging constants ***
Public Const ERRORLogLevel = 2

Public Const LogLevelError = 1
Public Const LogLevelTrace = 2
Public Const LogLevelInfo = 4
Public Const LogLevelWarning = 8

Public Const LogErrorOriginPortal = 1
Public Const LogErrorOriginAdmin = 2


Function DisplayError(szErrorHeader, sErrorMessage, sButtonCaption, szParentPage)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next

	Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
	Response.Write "<form METHOD=""GET"" ACTION=""" & szParentPage & """ id=form1 name=form1><TR>"
	Response.Write "<TD VALIGN=TOP WIDTH=""1%"">"
	Response.Write "<IMG SRC=""images/jobError.gif"" WIDTH=""55"" HEIGHT=""65"" BORDER=""0"" ALT="""">"
	Response.Write "</TD>"
	Response.Write "<TD VALIGN=TOP WIDTH=""99%"">"
	Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """><font color=""#cc0000""><b>" & szErrorHeader & "</b></font><BR><BR>"
	Response.Write sErrorMessage & "</font></TD>"
	Response.Write "<TR><TD></TD><TD><BR /><input TYPE=""SUBMIT"" CLASS=""buttonClass"" VALUE=""" & sButtonCaption & """ id=1 name=1>"
	Response.Write "</TD>"
	Response.Write "</TR></form>"
	Response.Write "</TABLE>"

	DisplayError = Err.number
End Function

Function DisplayLoginError(sErrorHeader, sErrorMessage)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next

	Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"
	Response.Write "<TR><TD BGCOLOR=""#000000""><IMG SRC=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD></TR>"
	Response.Write "<TR>"
	Response.Write "<TD VALIGN=TOP>"
	Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """><font color=""#cc0000""><b>" & sErrorHeader & "</b></font><BR>"
	Response.Write sErrorMessage & "</font>"
	Response.Write "</TD>"
	Response.Write "</TR>"
	Response.Write "<TR><TD BGCOLOR=""#000000""><IMG SRC=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD></TR>"
	Response.Write "</TABLE><BR>"

	DisplayLoginError = Err.number
End Function

Function LogError(sErrMsg)
'*******************************************************************************
'Purpose: Writes to the c:\temp\qeweb.log file sErrMsg
'Inputs:  sErrMsg
'*******************************************************************************
    Const FOR_APPENDING = 8
    Dim oFso
    Dim oTs
    Set oFso = Server.CreateObject("Scripting.FileSystemObject")
    If IsObject(oFso) Then
        Set oTs = oFso.OpenTextFile("c:\temp\qeweb.log", FOR_APPENDING, True)
        If IsObject(oTs) Then
            oTs.WriteLine DisplayDateAndTime(Date, Time) & Chr(9) & sErrMsg
            oTs.Close
            Set oTs = Nothing
        End If
        Set oFso = Nothing
    End If
End Function


Function LogMemory(sPage, sFunction, lStart, lEnd)
'*******************************************************************************
'Purpose: Writes to the c:\temp\qeweb.log file sErrMsg
'Inputs:  sErrMsg
'*******************************************************************************
    Const FOR_APPENDING = 8
    Dim oFso
    Dim oTs

    On Error Resume Next

    Set oFso = Server.CreateObject("Scripting.FileSystemObject")
    If IsObject(oFso) Then
        If strcomp(Right(Server.MapPath("."), 5), "Admin", vbTextCompare) = 0 Then
			Set oTs = oFso.OpenTextFile(Server.MapPath("Logs\mem.log"), FOR_APPENDING, True)
		Else
			Set oTs = oFso.OpenTextFile(Server.MapPath("Admin\Logs\mem.log"), FOR_APPENDING, True)
		End If

        If IsObject(oTs) Then
            oTs.WriteLine sPage & ";" & sFunction & ";" & lStart & ";" & lEnd & ";" & (lEnd - lStart)
            oTs.Close
            Set oTs = Nothing
        End If
        Set oFso = Nothing
    End If

End Function

Function LogErrorXML(aConnectionInfo, sErrID, sErrDesc, sErrSource, sFile, sASPFunc, sAPIFunc, sComments, iErrorLevel)
'*******************************************************************************
'Purpose: To log the different errors in XML format
'Inputs:  aConnectionInfo, sErrID, sErrDesc, sFile, sASPFunc, sAPIFunc, sComments, iErrorLevel
'*******************************************************************************
    On Error Resume Next
    Const FOR_APPENDING = 8
    Dim oFso
    Dim oTs
    Dim sLogFileName
    Dim sFolderPath
    Dim iErrOrigin

    If iErrorLevel <= ERRORLogLevel Then
		sLogFileName = Year(Date) & "-" & Month(Date) & "-" & Day(Date) & "err.log"
        Set oTS = Nothing

        Set oFso = Server.CreateObject("Scripting.FileSystemObject")
        If IsObject(oFso) Then
            If StrComp(Right(Server.MapPath("."), 5), "Admin", vbTextCompare) = 0 Then
				sFolderPath = "..\Logs"
			Else
				sFolderPath = "Logs"
			End If

			Set oTs = oFso.OpenTextFile(Server.MapPath(sFolderPath & "\" & sLogFileName), FOR_APPENDING, True)
			iErrOrigin = LogErrorOriginAdmin

            If Err.Number <> 0 Or oTS is Nothing Then
				'Check to see if the error is due to no "Logs" folder
				If Not (oFso.FolderExists(Server.MapPath(sFolderPath))) Then
					oFso.CreateFolder Server.MapPath(sFolderPath)
					Set oTs = oFso.OpenTextFile(Server.MapPath(sFolderPath & "\" & sLogFileName), FOR_APPENDING, True)
				End If
			End If

			oTs.WriteLine "<errMsg> <sortTime> " & NowAsUniqueString() & " </sortTime>"
			oTs.WriteLine " <time> " & DisplayDateAndTime(ConvertLocalTimeToUTC(Date), ConvertLocalTimeToUTC(Time)) & " </time>"
			oTs.WriteLine " <user> " & Server.HTMLEncode(aConnectionInfo(S_UID_CONNECTION) & " (" & aConnectionInfo(S_IP_ADDRESS_CONNECTION) & ")" ) & " </user>"
			oTs.WriteLine " <errID> " & Server.HTMLEncode(sErrID) & " </errID> "
			oTs.WriteLine " <errDesc> " & Server.HTMLEncode(sErrDesc) & " </errDesc> "
			oTs.WriteLine " <errSrc> " & Server.HTMLEncode(sErrSource) & " </errSrc> "
			oTs.WriteLine " <file> " & Server.HTMLEncode(sFile) & " </file> "
			oTs.WriteLine " <ASPFunc> " & Server.HTMLEncode(sASPFunc) & " </ASPFunc> "
			oTs.WriteLine " <APIFunc> " & Server.HTMLEncode(sAPIFunc) & " </APIFunc> "
			oTs.WriteLine " <comments> " & Server.HTMLEncode(sComments) & " </comments> "
			oTs.WriteLine " <errLevel> " & iErrorLevel & " </errLevel> "
			oTs.WriteLine " <errOrigin> " & iErrOrigin & " </errOrigin> "

			oTs.WriteLine " </errMsg>"
			oTs.Close
			Set oTs = Nothing
        End If
    End If

    Set oTs = Nothing
    Set oFso = Nothing

    Err.Clear
End Function

Function CleanString(sSource)
'*******************************************************************************
'Purpose: To replace &, <, and " in a string
'Inputs:  sSource
'Outputs: The new string
'*******************************************************************************
    CleanString = Replace(Replace(Replace(sSource, "&", "&#38;"), "<", "&#60;"), """", "&#34;")
End Function

Function NowAsUniqueString()
'*******************************************************************************
'Purpose: To return the date and time as a unique sortable string
'Outputs: A string with the date and time
'*******************************************************************************
	On Error Resume Next
    Dim n
    Dim sResult

    n = Now()
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
    NowAsUniqueString = sResult
End Function
%>