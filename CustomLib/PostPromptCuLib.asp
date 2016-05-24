<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/PostPromptCoLib.asp" -->
<!--#include file="../CoreLib/SubscribeCoLib.asp" -->
<%
Function ParseRequestForPostPrompt(oRequest, sSubGUID, sQOID, sSource)
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
	sSource = ""

	sSubGUID = Trim(CStr(oRequest("subGUID")))
	sQOID = Trim(CStr(oRequest("qoid")))
	sSource = Trim(CStr(oRequest("src")))

	If Err.number <> NO_ERR Then
	    lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "ParseRequestForPostPrompt", "", "Error setting variables equal to Request variables", LogLevelError)
	Else
	    If Len(sSubGUID) = 0 Then
	        lErrNumber = URL_MISSING_PARAMETER
	    End If
	End If

	ParseRequestForPostPrompt = lErrNumber
	Err.Clear
End Function

Function UpdateCache_CreateProfile(sCacheXML, sQOID, sPrefID, sProfileName, sPrefDesc, bIsDefault)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oAnswer
    Dim oCurrQO
    Dim oProfilesMI
    Dim oNewProfile
    Dim oOldDefault

    lErrNumber = NO_ERR

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
    oCacheDOM.async = False
    If oCacheDOM.loadXML(sCacheXML) = False Then
    	lErrNumber = ERR_XML_LOAD_FAILED
    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "UpdateProfileNames", "", "Error loading sCacheXML", LogLevelError)
    Else
		Set oCurrQO = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@tp = '" & TYPE_QUESTION & "' and @id = '" & sQOID & "']")
        Set oAnswer = oCurrQO.selectSingleNode("answer")
        oAnswer.setAttribute "prefID", sPrefID
        'oAnswer.setAttribute "originaln", sPrefDesc
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "UpdateProfileNames", "", "Error updating original profile name", LogLevelError)
        End If

        Set oProfilesMI = oCurrQO.selectSingleNode("mi")
        If oProfilesMI Is Nothing Then
			Set oProfilesMI = oCacheDOM.createElement("mi")
			Call oCurrQO.appendChild(oProfilesMI)
        End If

        Set oNewProfile = oCacheDOM.createElement("oi")
        Call oNewProfile.setAttribute("tp", CStr(TYPE_PROFILE))
        Call oNewProfile.setAttribute("id", sPrefID)
        Call oNewProfile.setAttribute("n", sProfileName)
        Call oNewProfile.setAttribute("des", sPrefDesc)
        If bIsDefault Then
			Call oNewProfile.setAttribute("def", "1")
			Set oOldDefault = oProfilesMI.selectSingleNode("oi[@tp = '" & TYPE_PROFILE & "' and @def = '1']")
			If Not oOldDefault Is Nothing Then
				Call oOldDefault.setAttribute("def", "0")
			End If
        Else
        	Call oNewProfile.setAttribute("def", "0")
        End If
        Call oProfilesMI.appendChild(oNewProfile)
    End If

    sCacheXML = oCacheDOM.xml

    Set oCacheDOM = Nothing
    Set oAnswer = Nothing
    Set oProfilesMI = Nothing
    Set oNewProfile = Nothing
    Set oOldDefault = Nothing

    UpdateProfileNames = lErrNumber
    Err.Clear
End Function

Function CheckForProfile(sCacheXML, sServiceID, sQOID, sISID, sProfileName, sOriginalProfileName, sPrefID, sPrefDesc, bIsDefault, bIsExistingProfile, bHasPrefDef)
'********************************************************
'*Purpose:
'*Inputs:	sCacheXML, sQOID
'*Outputs:	sISID, sProfileName, sOriginalProfileName, sPrefID, sPrefDesc, bIsExistingProfile
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oAnswer
    Dim oCurrQO
    Dim oExistingProfile
    Dim oSub

    lErrNumber = NO_ERR
    sProfileName = ""
    sOriginalProfileName = ""
    sPrefID = ""
    bIsExistingProfile = False
    sISID = ""
	bHasPrefDef = False

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "DeleteProfileCuLib.asp", "ParseInfoFromCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		Set oSub = oCacheDOM.selectSingleNode("/mi/sub")
		sServiceID = oSub.getAttribute("svcid")

		set oCurrQO = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@tp='" & TYPE_QUESTION & "' $and$ @id='" & sQOID & "']")
		sISID = oCurrQO.getAttribute("isid")

		Set oAnswer = oCurrQO.selectSingleNode("answer")
        'sProfileName = Server.HTMLEncode(oAnswer.getAttribute("n"))
        sProfileName = oAnswer.getAttribute("n")
        'OriginalProfileName = Server.HTMLEncode(oAnswer.getAttribute("originaln"))
        sPrefID = oAnswer.getAttribute("prefID")
        sPrefDesc =  oAnswer.getAttribute("desc")
        'sPrefDesc =  Server.HTMLEncode(oAnswer.getAttribute("desc"))
        bIsDefault = (oAnswer.getAttribute("def") = "1")
        If not (oAnswer.selectSingleNode("*") Is Nothing) Then
            bHasPrefDef = True
        End If

        bIsExistingProfile = False
        Set oExistingProfile = Nothing
        'set oExistingProfile = oCurrQO.selectSingleNode("mi/oi[@tp=""" & TYPE_PROFILE & """ $and$ @n=""" & Server.HTMLEncode(sProfileName) & """]")
        'set oExistingProfile = oCurrQO.selectSingleNode("mi/oi[@tp=""" & TYPE_PROFILE & """ $and$ @n=""" & Replace(sProfileName, """", """""") & """]")
        'If not oExistingProfile is nothing Then
        '    bIsExistingProfile = True
        '    sPrefID = oExistingProfile.getAttribute("id")
        '    Call oAnswer.setAttribute("prefID", sPrefID)
        'End If
        For each oExistingProfile in oCurrQO.selectNodes("mi/oi[@tp=""" & TYPE_PROFILE & """]")
			If oExistingProfile.getAttribute("n") = sProfileName Then
			    bIsExistingProfile = True
			    sPrefID = oExistingProfile.getAttribute("id")
			    Call oAnswer.setAttribute("prefID", sPrefID)
			    Exit For
			End If
		Next

	    If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "CheckForProfile", "", "Error retrieving profile names", LogLevelError)
        End If
    End If

    sCacheXML = oCacheDOM.xml
    Set oCacheDOM = Nothing
    Set oAnswer = nothing
    Set oCurrQO = nothing
    Set oExistingProfile = nothing
    Set oSub = Nothing

    CheckForProfile = lErrNumber
    Err.Clear
End Function

Function GetPreviousQuestionObject(sQOID, sCacheXML, sNextQOID, sNextPrefID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oCurrentQuestion
    Dim bLoop

    lErrNumber = NO_ERR
    sNextPrefID = ""

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
    oCacheDOM.async = False
    If oCacheDOM.loadXML(sCacheXML) = False Then
    	lErrNumber = ERR_XML_LOAD_FAILED
    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "GetPreviousQuestionObject", "", "Error loading sCacheXML", LogLevelError)
    Else
        Set oCurrentQuestion = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@id = '" & sQOID & "']")
        If oCurrentQuestion Is Nothing Then
            'add error handling
        Else
			bLoop = True
   			While (bLoop)
				Set oCurrentQuestion = oCurrentQuestion.previousSibling
				If oCurrentQuestion Is Nothing Then
					sNextQOID = "first"
					bLoop = False
				ElseIf Strcomp(oCurrentQuestion.getAttribute("hidden"), "0", vbTextCompare) = 0 Then
					sNextQOID = oCurrentQuestion.getAttribute("id")
					If Not (oCurrentQuestion.selectSingleNode("answer") Is Nothing) Then
					    sNextPrefID = oCurrentQuestion.selectSingleNode("answer").getAttribute("prefID")
					End If
					bLoop = False
				End If
			WEnd
        End If
    End If

    Set oCacheDOM = Nothing
    Set oCurrentQuestion = Nothing

    GetPreviousQuestionObject = lErrNumber
    Err.Clear
End Function

Function GetNextQuestionObject(sQOID, sCacheXML, sNextQOID, sNextPrefID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oCurrentQuestion
    Dim bLoop

    lErrNumber = NO_ERR
    sNextPrefID = ""

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
    oCacheDOM.async = False
    If oCacheDOM.loadXML(sCacheXML) = False Then
    	lErrNumber = ERR_XML_LOAD_FAILED
    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "GetNextQuestionObject", "", "Error loading sCacheXML", LogLevelError)
    Else
        Set oCurrentQuestion = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@id = '" & sQOID & "']")
        If oCurrentQuestion Is Nothing Then
            'add error handling
        Else
			bLoop = True
			While (bLoop)
				Set oCurrentQuestion = oCurrentQuestion.nextSibling
				If oCurrentQuestion Is Nothing Then
					sNextQOID = "last"
					bLoop = False
				ElseIf Strcomp(oCurrentQuestion.getAttribute("hidden"), "0", vbTextCompare) = 0  Then
					sNextQOID = oCurrentQuestion.getAttribute("id")
					If Not (oCurrentQuestion.selectSingleNode("answer") Is Nothing) Then
					    sNextPrefID = oCurrentQuestion.selectSingleNode("answer").getAttribute("prefID")
					End If
					bLoop = False
				End If
			WEnd
        End If
    End If

    Set oCacheDOM = Nothing
    Set oCurrentQuestion = Nothing

    GetNextQuestionObject = lErrNumber
    Err.Clear
End Function

Function cu_CreateProfile(sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, sProfileName, sProfileDesc, bIsDefault)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "cu_CreateProfile"
    Dim lErrNumber
    Dim sSessionID

    lErrNumber = NO_ERR
    sSessionID = GetSessionID()
    sPreferenceObjectID = GetGUID()

    lErrNumber = co_CreateProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, sProfileName, sProfileDesc, bIsDefault)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_CreateProfile", LogLevelTrace)
    End If

    cu_CreateProfile = lErrNumber
    Err.Clear
End Function

Function cu_UpdatePreferenceObjects(sCacheXML, sQOID, sPrefID, sServiceID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: add error handling
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim sSessionID
    Dim oCacheDOM
    Dim oAnswer
    Dim sNewGUID
    Dim sPrefDataXML

    lErrNumber = NO_ERR
    sSessionID = GetSessionID()
    sPrefDataXML = ""

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "UpdatePreferenceObject_Profile", "", "Error while calling co_UpdatePreferenceObjects", LogLevelTrace)
    Else
        Set oAnswer = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@tp = '" & TYPE_QUESTION & "' $and$ @id = '" & sQOID & "']/answer")

        sPrefDataXML = sPrefDataXML & "<updatePreferenceObjects>"
        sPrefDataXML = sPrefDataXML & "<subsetID>" & oCacheDOM.selectSingleNode("/mi/sub").getAttribute("sbstid") & "</subsetID>"
        sPrefDataXML = sPrefDataXML & "<serviceID>" & sServiceID & "</serviceID>"
        sPrefDataXML = sPrefDataXML & "<sessionID>" & sSessionID & "</sessionID>"
        sPrefDataXML = sPrefDataXML & "<channelID>" & sChannel & "</channelID>"
        sPrefDataXML = sPrefDataXML & "<personalization>"
        sPrefDataXML = sPrefDataXML & "<qo id='" & oAnswer.parentNode.getAttribute("id") & "'>"
        sPrefDataXML = sPrefDataXML & "<INFO_SOURCE_ID>" & oAnswer.parentNode.getAttribute("isid") & "</INFO_SOURCE_ID>"
        sPrefDataXML = sPrefDataXML & "<PREFERENCE_ID>" & sPrefID & "</PREFERENCE_ID>"
        sPrefDataXML = sPrefDataXML & "<PROMPT_ANSWER>"
        sPrefDataXML = sPrefDataXML & Server.HTMLEncode(CStr(oAnswer.selectSingleNode("*").xml))
        sPrefDataXML = sPrefDataXML & "</PROMPT_ANSWER>"
        sPrefDataXML = sPrefDataXML & "</qo>"
        sPrefDataXML = sPrefDataXML & "</personalization>"
        sPrefDataXML = sPrefDataXML & "</updatePreferenceObjects>"

        Set oAnswer = Nothing
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_UpdatePreferenceObjects(sPrefDataXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "UpdatePreferenceObject_Profile", "", "Error while calling co_UpdatePreferenceObjects", LogLevelTrace)
        End If
    End If

    Set oCacheDOM = Nothing

    cu_UpdatePreferenceObjects = lErrNumber
    Err.Clear
End Function

Function cu_UpdateProfile(sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, sProfileName, sProfileDesc, bIsDefault)
'********************************************************
'*Purpose:	Update Profile Definition
'*Inputs:	sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, sProfileName, sProfileDesc, bIsDefault
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "cu_UpdateProfile"
    Dim lErrNumber
    Dim sSessionID

    lErrNumber = NO_ERR
    sSessionID = GetSessionID()

    lErrNumber = co_UpdateProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sInfoSourceID, sProfileName, sProfileDesc, bIsDefault)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_UpdateProfile", LogLevelTrace)
    End If

    cu_UpdateProfile = lErrNumber
    Err.Clear
End Function
%>