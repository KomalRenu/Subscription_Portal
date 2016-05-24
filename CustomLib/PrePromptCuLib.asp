<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/PrePromptCoLib.asp" -->
<%

Function ParseRequestForPrePrompt(oRequest, sSubGUID, sQOID, sSRC, sFolderID, sPrefObjID)
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
	sSRC = ""
	sFolderID = ""
	sPrefObjID = ""

	sSubGUID = Trim(CStr(oRequest("subGUID")))
	sQOID = Trim(CStr(oRequest("qoid")))
	sSRC = Trim(CStr(oRequest("src")))
	sFolderID = Trim(CStr(oRequest("folderID")))
	sPrefObjID = Trim(CStr(oRequest("prefID")))

	If Err.number <> NO_ERR Then
	    lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", "ParseRequestForPrompt", "", "Error setting variables equal to Request variables", LogLevelError)
	Else
	    If Len(sSubGUID) = 0 Then
	    	lErrNumber = URL_MISSING_PARAMETER
	    End If
	End If

	ParseRequestForPrePrompt = lErrNumber
	Err.Clear
End Function


Function GetStatusFlag(sCacheXML, sStatusFlag)
'********************************************************
'*Purpose: Checks the cacheXML for the statusFlag, which
'          specifies new or edit
'*Inputs: sCacheXML
'*Outputs: sStatusFlag
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM

    lErrNumber = NO_ERR
    sStatusFlag = ""

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
    oCacheDOM.async = False
    If oCacheDOM.loadXML(sCacheXML) = False Then
    	lErrNumber = ERR_XML_LOAD_FAILED
    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", "GetStatusFlag", "", "Error loading sCacheXML", LogLevelError)
    Else
        sStatusFlag = CStr(oCacheDOM.selectSingleNode("/mi/sub").getAttribute("sf"))
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", "GetStatusFlag", "", "Error retrieving statusFlag", LogLevelError)
        End If
    End If

    Set oCacheDOM = Nothing

    GetStatusFlag = lErrNumber
    Err.Clear
End Function

Function AddProfileAnswerToCache(sCacheXML, sQOID, sPrefObjID, sGetPreferenceObjectsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim lTempErr
    Dim oCacheDOM
    Dim oPrefObjDOM
    Dim oQuestion
    Dim oAnswer
    Dim sProfileXML
    Dim oProfileDOM
    Dim oProfile
    Dim sProfileName
    Dim sProfileDesc

    lErrNumber = NO_ERR

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
    oCacheDOM.async = False
    If oCacheDOM.loadXML(sCacheXML) = False Then
    	lErrNumber = ERR_XML_LOAD_FAILED
    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", "AddProfileAnswerToCache", "", "Error loading sCacheXML", LogLevelError)
    Else
        Set oQuestion = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@tp = '" & TYPE_QUESTION & "' and @id = '" & sQOID & "']")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", "AddProfileAnswerToCache", "", "Error retrieving question oi node", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
        Set oPrefObjDOM = Server.CreateObject("Microsoft.XMLDOM")
        oPrefObjDOM.async = False
        If oPrefObjDOM.loadXML(sGetPreferenceObjectsXML) = False Then
        	lErrNumber = ERR_XML_LOAD_FAILED
        	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", "AddProfileAnswerToCache", "", "Error loading sGetPreferenceObjectsXML", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
        lTempErr = cu_GetProfile(sPrefObjID, sQOID, sProfileXML)

        Set oProfileDOM = Server.CreateObject("Microsoft.XMLDOM")
        oProfileDOM.async = False
        oProfileDOM.loadXML(sProfileXML)

        Set oProfile = oProfileDOM.selectSingleNode("/mi/in/oi")
        If Not (oProfile Is Nothing) Then
            sProfileName = oProfile.getAttribute("n")
            sProfileDesc = oProfile.getAttribute("des")
            If StrComp(sProfileDesc, "null", vbBinaryCompare) = 0 Then sProfileDesc = ""
        End If

        Set oProfileDOM = Nothing
        Set oProfile = Nothing
    End If

    If lErrNumber = NO_ERR Then
        Set oAnswer = oQuestion.selectSingleNode("answer")
        If Not (oAnswer Is Nothing) Then
            oQuestion.removeChild(oAnswer)
        End If
        Set oAnswer = oQuestion.appendChild(oCacheDOM.createElement("answer"))
        oAnswer.setAttribute "n", sProfileName
        'Answer.setAttribute "originaln", sProfileName
        oAnswer.setAttribute "desc", sProfileDesc
        oAnswer.setAttribute "prefID", sPrefObjID
        oAnswer.appendChild(oPrefObjDOM.selectSingleNode("/mi/in/oi[@id = '" & sPrefObjID & "']/*"))
    End If

    If lErrNumber = NO_ERR Then
        sCacheXML = oCacheDOM.xml
    End If

    Set oCacheDOM = Nothing
    Set oPrefObjDOM = Nothing
    Set oQuestion = Nothing
    Set oAnswer = Nothing

    AddProfileAnswerToCache = lErrNumber
    Err.Clear
End Function

Function GetQuestionProperty(sGetDetailsForQuestionsXML, sQOID, bHiddenQO, sISID)
'********************************************************
'*Purpose:	Check if this QO is hidden or not
'*Inputs:	sGetDetailsForQuestionsXML, sQOID
'*Outputs:	bHiddenQO, sISID
'********************************************************
	On Error Resume Next
	Dim oDetailsDOM
	Dim oCurrQO
	Dim oISMProgID

	bHiddenQO = False
	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetDetailsForQuestionsXML, oDetailsDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "GetQuestionProperty", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		set oCurrQO = oDetailsDOM.selectSingleNode("/mi/qos/oi[@tp='" & TYPE_QUESTION & "' $and$ @id='" & sQOID & "']")
		sISID = oCurrQO.getAttribute("isid")
		'set oISMProgID = oDetailsDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_INFORMATION_SOURCE & "' $and$ @id='" & oCurrQO.getAttribute("isid") & "']/prs/pr[@n='ISM_admin_progid']")
		'If oISMProgID is nothing Then
		'	lErrNumber = ERR_CACHE_CONTENT
		'Else
		'	If strcomp(oISMProgID.getAttribute("v"), "UserDetailsISM.cUserDetails", vbBinaryCompare) = 0 Then
		'		bHiddenQO = True
		'	End If
		'End If
		If Len(oCurrQO.getAttribute("hidden")) > 0 Then
			bHiddenQO = True
		End If
	End If

	GetQuestionProperty = lErrNumber
	Err.Clear
End Function

Function AnswerHiddenQO(sCacheXML, sQOID, sISID)
'********************************************************
'*Purpose:	build <answer> for hidden QO
'*Inputs:	sCacheXML, sQOID, sISID
'*Outputs:	sCacheXML
'********************************************************
	On Error Resume Next
	Dim oCacheDOM
	Dim oCurrQO
	Dim oAnswer
	Dim sPrefID
	Dim oDefaultProfile

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		'lErrNumber = GetUserDefaultPersonalizationForQO(sQOID, sPrefID)
		'If lErrNumber <> NO_ERR Then
		'	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
		'Else
			Set oCurrQO = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@tp = '" & TYPE_QUESTION & "' and @id = '" & sQOID & "']")
			Call oCurrQO.setAttribute("isid", sISID)
			Set oDefaultProfile = oCurrQO.selectSingleNode("mi/oi[@tp = '" & TYPE_PROFILE & "' and @def='1']")
			If oDefaultProfile is nothing Then
				lErrNumber = ERR_USERDEFAULT_NOTEXIST
			Else
				Set oAnswer = oCacheDOM.createElement("answer")
				Call oCurrQO.appendChild(oAnswer)
				Call oAnswer.setAttribute("n", oDefaultProfile.getAttribute("n"))
				Call oAnswer.setAttribute("prefID", oDefaultProfile.getAttribute("id"))
				sCacheXML = oCacheDOM.xml
			End If
		'End If
    End If

	Set oCacheDOM = nothing
	Set oCurrQO = nothing
	Set oAnswer = nothing
	Set oDefaultProfile = nothing

	AnswerHiddenQO = lErrNumber
	Err.clear
End Function

'not used any more
Function GetUserDefaultPersonalizationForQO(sQOID, sPrefID)
'********************************************************
'*Purpose:	get preferenceID for a QOID
'*Inputs:	sQOID
'*Outputs:	sPrefID
'********************************************************
	On Error Resume Next
	Dim sSessionID
	Dim oUserDefaultDOM
	Dim sUserDefaultPersonalizationXML
	Dim oCurrQO

	sSessionID = GetSessionID()
	lErrNumber = co_getUserDefaultPersonalization(sSessionID, sUserDefaultPersonalizationXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "GetUserDefaultPersonalizationForQO", "cu_getUserDefaultPersonalization", "Error calling cu_getUserDefaultPersonalization", LogLevelTrace)
	Else
		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sUserDefaultPersonalizationXML, oUserDefaultDOM)
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "GetUserDefaultPersonalizationForQO", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
		Else
			set oCurrQO = oUserDefaultDOM.selectSingleNode("/mi/qos/oi[@tp='" & TYPE_QUESTION & "' $and$ @id='" & sQOID & "']")
			If oCurrQO is nothing Then
				lErrNumber = ERR_USERDEFAULT_NOTEXIST
			Else
				sPrefID = oCurrQO.selectSingleNode("mi/in/oi[@tp='" & TYPE_PREFERENCEOBJECT & "']").getAttribute("id")
				lErrNumber = Err.Number
			End If
		End If
	End If

	Set oUserDefaultDOM = nothing
	Set oCurrQO = nothing

	GetUserDefaultPersonalizationForQO = lErrNumber
	err.clear
End Function
%>