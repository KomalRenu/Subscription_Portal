<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/DeleteProfileCoLib.asp" -->

<%
const HYDRA_APIERROR_DELETE_PROFILE = "-1"

Function ParseRequestForDeleteProfile(oRequest, sSubGUID, sQOID, sPrefID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber

	lErrNumber = NO_ERR

	sSubGUID = ""
	sQOID = ""
	sPrefID = ""

	sSubGUID = Trim(CStr(oRequest("subGUID")))
	sQOID = Trim(CStr(oRequest("qoid")))
	sPrefID = Trim(CStr(oRequest("prefID")))

	If Err.number <> NO_ERR Then
	    lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "ParseRequestForPostPrompt", "", "Error setting variables equal to Request variables", LogLevelError)
	Else
	    If Len(sSubGUID)=0 OR Len(sQOID)=0 OR Len(sPrefID)=0 Then
	        lErrNumber = URL_MISSING_PARAMETER
	    End If
	End If

	ParseRequestForDeleteProfile = lErrNumber
	Err.Clear
End Function

Function ParseInfoFromCache(sCacheXML, sQOID, sISID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
	Dim lErrNumber
	Dim oCacheDOM
	Dim oCurrQO

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "DeleteProfileCuLib.asp", "ParseInfoFromCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		set oCurrQO = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@tp='" & TYPE_QUESTION & "' $and$ @id='" & sQOID & "']")
		sISID = oCurrQO.getAttribute("isid")
	End If

	set oCacheDOM = nothing
	set oCurrQO = nothing

	ParseInfoFromCache = lErrNumber
	Err.Clear
End Function

Function cu_DeleteProfile(sPreferenceObjectID, sQuestionObjectID, sInfoSourceID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim sSessionID
    Dim sDeleteProfileXML
    Dim bForceDelete

    lErrNumber = NO_ERR
    sSessionID = GetSessionID()

	bForceDelete = false

    lErrNumber = co_DeleteProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, bForceDelete, sDeleteProfileXML)
    If lErrNumber <> NO_ERR Then
        If lErrNumber = HYDRA_APIERROR_DELETE_PROFILE Then
        Else
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "cu_CreateProfile", "", "Error while calling co_DeleteProfile", LogLevelTrace)
		End If
    End If

    cu_DeleteProfile = lErrNumber
    Err.Clear
End Function

Function UpdateCacheXML_DeleteProfile(sCacheXML, sQOID, sPrefID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oProfiles
    Dim oProfileToDelete

    lErrNumber = NO_ERR

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
    oCacheDOM.async = False
    oCacheDOM.loadXML(sCacheXML)

    Set oProfiles = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@tp = '" & TYPE_QUESTION & "' and @id = '" & sQOID & "']/mi")

	If Not (oProfiles Is Nothing) Then
		set oProfileToDelete = oProfiles.selectSingleNode("oi[@tp = '" & TYPE_PROFILE & "' and @id = '" & sPrefID & "']")
		If Not (oProfileToDelete Is Nothing) Then
			oProfiles.removeChild(oProfileToDelete)
		End If
    End If

    sCacheXML = oCacheDOM.xml

    Set oCacheDOM = Nothing
    set oProfiles = Nothing
    set oProfileToDelete = Nothing

    UpdateCacheXML_DeleteProfile = lErrNumber
    Err.Clear
End Function
%>