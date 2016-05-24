<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/SubscribeCoLib.asp" -->
<%
	Function ParseRequestForModifySubscription(oRequest, sSubGUID, sStatusFlag)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'********************************************************
		On Error Resume Next
        Dim lErrNumber

        lErrNumber = NO_ERR

        sSubGUID = ""
        sStatusFlag = ""

        sSubGUID = Trim(CStr(oRequest("subGUID")))
		If Trim(CStr(oRequest("action"))) = "delete" Then
		    sStatusFlag = "2"
		End If

		If Err.number <> NO_ERR Then
		    lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifySubscriptionCuLib.asp", "ParseRequestForModifySubscription", "", "Error setting variables equal to Request variables", LogLevelError)
		Else
		    If Len(sSubGUID) = 0 And StrComp(sStatusFlag, "2", vbBinaryCompare) <> 0 Then
		    	lErrNumber = URL_MISSING_PARAMETER
		    End If
		End If

		ParseRequestForModifySubscription = lErrNumber
		Err.Clear
	End Function

Function ReadSubscriptionProperties(sCacheXML, iNumQuestions, sStatusFlag, sSubsSetID, sServiceID, sAddressID, sFolderID, sSubsEnabled, sPublicationID, sOriginalPublicationID, sOriginalSubsSetID, sTransPropsID, sSubID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oSub

    lErrNumber = NO_ERR

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
	oCacheDOM.async = False
    oCacheDOM.loadXML(sCacheXML)

    iNumQuestions = CInt(oCacheDOM.selectNodes("//oi[@tp = '" & TYPE_QUESTION & "' $and$ @hidden='0']").length)
    Set oSub = oCacheDOM.selectSingleNode("/mi/sub")
    sStatusFlag = oSub.getAttribute("sf")
    sSubsSetID = oSub.getAttribute("sbstid")
    sServiceID = oSub.getAttribute("svcid")
    sAddressID = oSub.getAttribute("adid")
    sFolderID = oSub.getAttribute("fid")
    sSubsEnabled = oSub.getAttribute("enf")
    sPublicationID = oSub.getAttribute("pubid")
    sOriginalPublicationID = oSub.getAttribute("epubid")
    sOriginalSubsSetID = oSub.getAttribute("esbstid")
    sTransPropsID = oSub.getAttribute("trps")
    sSubID = oSub.getAttribute("subid")

    Set oCacheDOM = Nothing
    Set oSub = Nothing

    ReadSubscriptionProperties = lErrNumber
    Err.Clear
End Function

Function cu_AddSubscription(sSubsSetID, sServiceID, sAddressID, sTransPropsID, sSubsGUID, sSubsEnabled, sCacheXML)
'********************************************************
'*Purpose:
'*Inputs: sSubsSetID, sAddressID, sTransPropsID, sSubsGUID, sSubsEnabled
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_AddSubscription"
	Dim lErrNumber
	Dim sSubDataXML
	Dim sSessionID
    Dim oCacheDOM
    Dim oQuestions
    Dim oCurrentQuestion
    Dim sPreferenceID
	Dim oAnswer

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifySubscriptionCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
    Else
        Set oQuestions = oCacheDOM.selectNodes("/mi/qos/mi/in/oi[@tp = '" & TYPE_QUESTION & "']")
    End If

	If lErrNumber = NO_ERR Then
	    sSubDataXML = "<addSubscription>"
	    sSubDataXML = sSubDataXML & "<subsetID>" & sSubsSetID & "</subsetID>"
	    sSubDataXML = sSubDataXML & "<sessionID>" & sSessionID & "</sessionID>"
	    sSubDataXML = sSubDataXML & "<serviceID>" & sServiceID & "</serviceID>"
	    sSubDataXML = sSubDataXML & "<channelID>" & GetCurrentChannel() & "</channelID>" 'channelID
	    sSubDataXML = sSubDataXML & "<subscription>"
	        sSubDataXML = sSubDataXML & "<SUBSCRIPTION_ID></SUBSCRIPTION_ID>"
	        sSubDataXML = sSubDataXML & "<SUBSCRIPTION_GUID>" & sSubsGUID & "</SUBSCRIPTION_GUID>"
	        sSubDataXML = sSubDataXML & "<SUBSCRIPTION_SET_ID>" & sSubsSetID & "</SUBSCRIPTION_SET_ID>"
	        sSubDataXML = sSubDataXML & "<ADDRESS_ID>" & sAddressID & "</ADDRESS_ID>"
	        sSubDataXML = sSubDataXML & "<ACCOUNT_ID></ACCOUNT_ID>"
	        sSubDataXML = sSubDataXML & "<TRANS_PROPS_ID>" & sTransPropsID & "</TRANS_PROPS_ID>"
	        sSubDataXML = sSubDataXML & "<ADD_TRANS_PROPS>1</ADD_TRANS_PROPS>"
	        sSubDataXML = sSubDataXML & "<STATUS>" & sSubsEnabled & "</STATUS>"
	    sSubDataXML = sSubDataXML & "</subscription>"
	    sSubDataXML = sSubDataXML & "<personalization>"

	    For Each oCurrentQuestion In oQuestions
			Set oAnswer = oCurrentQuestion.selectSingleNode("answer")
			If Len(oAnswer.getAttribute("prefID")) > 0 Then
				sPreferenceID = oAnswer.getAttribute("prefID")
			Else
				sPreferenceID = GetGUID()
				Call oAnswer.setAttribute("prefID", sPreferenceID)
			End If

	        sSubDataXML = sSubDataXML & "<qo id='" & oCurrentQuestion.getAttribute("id") & "'>"
	        sSubDataXML = sSubDataXML & "<INFO_SOURCE_ID>" & oCurrentQuestion.getAttribute("isid") & "</INFO_SOURCE_ID>"

	        'If Strcomp(oCurrentQuestion.getAttribute("hidden"), "0", vbTextCompare) = 0  Then
			If CLng(oCurrentQuestion.getAttribute("qtp")) <> QO_TYPE_SLICING Then
				sSubDataXML = sSubDataXML & "<PREFERENCE_ID>"
			    sSubDataXML = sSubDataXML & sPreferenceID
				sSubDataXML = sSubDataXML & "</PREFERENCE_ID>"
				If (Len(CStr(oAnswer.getAttribute("n"))) > 0) And (oAnswer.selectSingleNode("*") Is Nothing) Then
				    sSubDataXML = sSubDataXML & "<PROFILE>1</PROFILE>"
				Else
				    sSubDataXML = sSubDataXML & "<PROMPT_ANSWER>"
				    'Escaped prompt answer XML
				    sSubDataXML = sSubDataXML & Server.HTMLEncode(oAnswer.selectSingleNode("*").xml)
				    sSubDataXML = sSubDataXML & "</PROMPT_ANSWER>"
				End If
			End If
	        sSubDataXML = sSubDataXML & "<QUESTION_TYPE>" & oCurrentQuestion.getAttribute("qtp") & "</QUESTION_TYPE>"
	        sSubDataXML = sSubDataXML & "</qo>"
	    Next

	    sSubDataXML = sSubDataXML & "</personalization>"
	    sSubDataXML = sSubDataXML & "</addSubscription>"
	End If

	If lErrNumber = NO_ERR Then
		sCacheXML = oCacheDOM.xml
	    lErrNumber = co_AddSubscription(sSubDataXML)
	    If lErrNumber <> NO_ERR Then
	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifySubscriptionCuLib.asp", PROCEDURE_NAME, "", "Error calling co_AddSubscription", LogLevelTrace)
	    End If
	End If

    Set oCacheDOM = Nothing
    Set oQuestions = Nothing
    Set oCurrentQuestion = Nothing

	cu_AddSubscription = lErrNumber
	Err.Clear
End Function

Function cu_EditSubscription(sSubsSetID, sServiceID, sAddressID, sTransPropsID, sSubsGUID, sSubsEnabled, sSubID, sCacheXML)
'********************************************************
'*Purpose:
'*Inputs: sSubsSetID, sAddressID, sTransPropsID, sSubsGUID, sSubsEnabled
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_EditSubscription"
	Dim lErrNumber
	Dim sSubDataXML
	Dim sSessionID
    Dim oCacheDOM
    Dim oQuestions
    Dim oCurrentQuestion
    Dim asSubscriptionID(0)
	Dim oAnswer
	Dim sPreferenceID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
	asSubscriptionID(0) = sSubsGUID

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifySubscriptionCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
    Else
        Set oQuestions = oCacheDOM.selectNodes("/mi/qos/mi/in/oi[@tp = '" & TYPE_QUESTION & "']")
    End If

	'Expire all subscriptions before editing. Although only
	'subscriptions that are in the document repository must get expired,
	'we expire all and don't worry about the results.
    If lErrNumber = NO_ERR Then
        lErrNumber = co_ExpireSubscription(sSessionID, asSubscriptionID)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifySubscriptionCuLib.asp", PROCEDURE_NAME, "", "Error calling co_ExpireSubscription", LogLevelTrace)
        End If
    End If

	If lErrNumber = NO_ERR Then
	    sSubDataXML = "<editSubscription>"
	    sSubDataXML = sSubDataXML & "<subsetID>" & sSubsSetID & "</subsetID>"
	    sSubDataXML = sSubDataXML & "<sessionID>" & sSessionID & "</sessionID>"
	    sSubDataXML = sSubDataXML & "<serviceID>" & sServiceID & "</serviceID>"
	    sSubDataXML = sSubDataXML & "<channelID>" & sChannel & "</channelID>" 'channelID
	    sSubDataXML = sSubDataXML & "<subscription>"
	        sSubDataXML = sSubDataXML & "<SUBSCRIPTION_ID>" & sSubID & "</SUBSCRIPTION_ID>"
	        sSubDataXML = sSubDataXML & "<SUBSCRIPTION_GUID>" & sSubsGUID & "</SUBSCRIPTION_GUID>"
	        sSubDataXML = sSubDataXML & "<SUBSCRIPTION_SET_ID>" & sSubsSetID & "</SUBSCRIPTION_SET_ID>"
	        sSubDataXML = sSubDataXML & "<ADDRESS_ID>" & sAddressID & "</ADDRESS_ID>"
	        sSubDataXML = sSubDataXML & "<ACCOUNT_ID></ACCOUNT_ID>"
	        sSubDataXML = sSubDataXML & "<TRANS_PROPS_ID>" & sTransPropsID & "</TRANS_PROPS_ID>"
	        sSubDataXML = sSubDataXML & "<ADD_TRANS_PROPS>1</ADD_TRANS_PROPS>"
	        sSubDataXML = sSubDataXML & "<STATUS>" & sSubsEnabled & "</STATUS>"
	    sSubDataXML = sSubDataXML & "</subscription>"
	    sSubDataXML = sSubDataXML & "<personalization>"

	    For Each oCurrentQuestion In oQuestions
	        Set oAnswer = oCurrentQuestion.selectSingleNode("answer")
			If Len(oAnswer.getAttribute("prefID")) > 0 Then
				sPreferenceID = oAnswer.getAttribute("prefID")
			Else
				sPreferenceID = GetGUID()
				Call oAnswer.setAttribute("prefID", sPreferenceID)
			End If

			sSubDataXML = sSubDataXML & "<qo id='" & oCurrentQuestion.getAttribute("id") & "'>"
	        sSubDataXML = sSubDataXML & "<INFO_SOURCE_ID>" & oCurrentQuestion.getAttribute("isid") & "</INFO_SOURCE_ID>"

	        If CLng(oCurrentQuestion.getAttribute("qtp")) <> QO_TYPE_SLICING Then
				sSubDataXML = sSubDataXML & "<PREFERENCE_ID>" & sPreferenceID & "</PREFERENCE_ID>"

				If (Len(CStr(oAnswer.getAttribute("n"))) > 0) And (oAnswer.selectSingleNode("*") Is Nothing) Then
				    sSubDataXML = sSubDataXML & "<PROFILE>1</PROFILE>"
				Else
				    sSubDataXML = sSubDataXML & "<PROMPT_ANSWER>"
				    'Escaped prompt answer XML
				    sSubDataXML = sSubDataXML & Server.HTMLEncode(oAnswer.selectSingleNode("*").xml)
				    sSubDataXML = sSubDataXML & "</PROMPT_ANSWER>"
				End If
			End If

            sSubDataXML = sSubDataXML & "<QUESTION_TYPE>" & oCurrentQuestion.getAttribute("qtp") & "</QUESTION_TYPE>"
	        sSubDataXML = sSubDataXML & "</qo>"
	    Next

	    sSubDataXML = sSubDataXML & "</personalization>"
	    sSubDataXML = sSubDataXML & "</editSubscription>"
	End If

	If lErrNumber = NO_ERR Then
	    sCacheXML = oCacheDOM.xml
	    lErrNumber = co_EditSubscription(sSubDataXML)
	    If lErrNumber <> NO_ERR Then
	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifySubscriptionCuLib.asp", PROCEDURE_NAME, "", "Error calling co_EditSubscription", LogLevelTrace)
	    End If
	End If

    Set oCacheDOM = Nothing
    Set oQuestions = Nothing
    Set oCurrentQuestion = Nothing

	cu_EditSubscription = lErrNumber
	Err.Clear
End Function

Function cu_DeleteSubscriptions()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Talk to Gunther about using oRequest, and not passing DelSubsGUID
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_DeleteSubscriptions"
	Dim lErrNumber
	Dim i
	Dim sSubSetID
	Dim sSubGUID
	Dim asSubscriptionID()
	Dim sSessionID
	Dim sDelSubDataXML
	Dim oXMLDOM
	Dim oCurrentSubSet
	Dim oNewNode
	Dim iSeparator
	Dim sServiceID
	Dim temArray

	lErrNumber = NO_ERR
    sSessionID = GetSessionID()

	Redim asSubscriptionID(oRequest("delSubsGUID").Count - 1)

    If lErrNumber = NO_ERR Then
        sDelSubDataXML = "<deleteSubscription>"
        sDelSubDataXML = sDelSubDataXML & "<sessionID>" & sSessionID & "</sessionID>"
        sDelSubDataXML = sDelSubDataXML & "<channelID>" & sChannel & "</channelID>"
        sDelSubDataXML = sDelSubDataXML & "</deleteSubscription>"

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sDelSubDataXML, oXMLDOM)

	    For i=1 To oRequest("delSubsGUID").Count
	        temArray = Split(CStr(oRequest("delSubsGUID")(i)), ";", -1, vbBinaryCompare)
            sSubSetID = CStr(temArray(0))
            sSubGUID = CStr(temArray(1))
            sServiceID = CStr(temArray(2))
            asSubscriptionID(i-1) = sSubGUID

	        Set oCurrentSubSet = oXMLDOM.selectSingleNode("/deleteSubscription/subscription[subsetID = '" & sSubSetID & "']")
            If (oCurrentSubSet Is Nothing) Then
                Set oCurrentSubSet = oXMLDOM.selectSingleNode("deleteSubscription").appendChild(oXMLDOM.createElement("subscription"))
                Set oNewNode = oCurrentSubSet.appendChild(oXMLDOM.createElement("subsetID"))
                oNewNode.text = CStr(sSubSetID)
				Set oNewNode = oCurrentSubSet.appendChild(oXMLDOM.createElement("serviceID"))
				oNewNode.text = CStr(sServiceID)
            End If
            Set oNewNode = oCurrentSubSet.appendChild(oXMLDOM.createElement("SUBSCRIPTION_GUID"))
            oNewNode.text = CStr(sSubGUID)
	    Next
    End If

	'Expire all subscriptions before deleting. Although only
	'subscriptions that are in the document repository must get expired,
	'we expire all and don't worry about the results.
    If lErrNumber = NO_ERR Then
        lErrNumber = co_ExpireSubscription(sSessionID, asSubscriptionID)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifySubscriptionCuLib.asp", PROCEDURE_NAME, "", "Error calling co_ExpireSubscription", LogLevelTrace)
        End If
    End If

    'Now delete the subscription:
    If lErrNumber = NO_ERR Then
        lErrNumber = co_DeleteSubscriptions(CStr(oXMLDOM.xml))
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifySubscriptionCuLib.asp", PROCEDURE_NAME, "", "Error calling co_DeleteSubscriptions", LogLevelTrace)
        End If
    End If

    Set oXMLDOM = Nothing
    Set oCurrentSubSet = Nothing
    Set oNewNode = Nothing

	cu_DeleteSubscriptions = lErrNumber
	Err.Clear
End Function
%>