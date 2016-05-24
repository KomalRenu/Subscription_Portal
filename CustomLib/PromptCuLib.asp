<%'** Copyright © 2000-2012 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!-- #include file="PromptConstCuLib.asp"-->
<!-- #include file="PromptDisplayCuLib.asp"-->
<!-- #include file="PromptProcessCuLib.asp"-->
<!-- #include file="PromptSearchCuLib.asp"-->
<!-- #include file="PromptCommonCuLib.asp"-->
<!-- #include file="FilterDetailsCuLib.asp"-->
<!-- #include file="PersonalizeCuLib.asp"-->

<%
'Browsing Variables
	Dim oObjectsNodeTBrowsing
	Dim oRefNodeTBrowsing
	Dim vItemTBrowsing
	Dim sDescriptionTBrowsing, sTimeTBrowsing
	Dim asObjectsTBrowsing(1)
	Dim iTBrowsing
	Dim sIconTBrowsing
	Dim bWebDeleteTBrowsing
	Dim iNodeCounter

'NOTE: Fix for SP1 was made on function CheckAllClosedPrompts. Need to validate it with SDK in Boyd so
'	   this fix can be reviewed and validated.

Function LoadPromptQuestionsXMLForJob(aConnectionInfo, aPromptGeneralInfo, oSession, aPromptInfo, oRequest, sErrDescription)
'***********************************************************************************************
'Purpose: Get aPromptGeneralInfo(PROMPT_O_QUESTIONSXML) from file or from XML API
'Inputs:  aConnectionInfo, aPromptGeneralInfo(PROMPT_S_MSGID), aPromptGeneralInfo(PROMPT_B_ISDOC), oSession
'Outputs: aPromptGeneralInfo(PROMPT_L_MAXPIN), aPromptInfo, sErrDescription
'***********************************************************************************************
    On Error Resume Next
    Dim sFileName
    Dim oFS
    Dim oTFile
    Dim oFile
    Dim sPromptQuestionsXML
    Dim bAllClosed
    Dim bFileExist
    Dim lErrNumber
	Dim oAllPrompts
	Dim oExpression
	Dim oNode

	'Hydra - Could be already created
    If aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT) Is Nothing Then

		'create Helper Object
		lErrNumber = CreatePromptsHelperObject(aConnectionInfo, oAllPrompts, sErrDescription)
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptCuLib.asp", "LoadPromptQuestionsXMLForJob", "", "Error in call to CreatePromptsHelperObject", LogLevelTrace)
		Else
			oAllPrompts.MessageID = CStr(aPromptGeneralInfo(PROMPT_S_MSGID))

			If Len(oRequest("index")) > 0 Then
				oAllPrompts.StateID = CLng(oRequest("index"))
			Else
				oAllPrompts.StateID = -1
			End If

			If aPromptGeneralInfo(PROMPT_B_ISDOC) Then
				oAllPrompts.DocumentOrigin = True
			End If
			oAllPrompts.Locale = CLng(GetServerLanguage(aConnectionInfo, Application.Value("iSourcePerm")))
			oAllPrompts.ResultFlags = DssXmlResultNoDerivedPromptXML Or DssXmlResultGrid
			oAllPrompts.ExecutionFlags = DssXmlExecutionUpdateCache Or DssXmlExecutionUseCache

			oAllPrompts.Init
			If Err.number=0 Then
				set aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT) = oAllPrompts
				If Len(aPromptGeneralInfo(PROMPT_S_ANSWERSXML)) > 0 Then
					aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).ShortAnswerXML = aPromptGeneralInfo(PROMPT_S_ANSWERSXML)
					If Err.number <> NO_ERR Then
						Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.Description), Err.source, "PromptCuLib.asp", "LoadPromptQuestionsXMLForJob", "", "Error in call to set the ShortAnswerXML", LogLevelTrace)
					End If
				End If
			else
				'If Err.number=-2147205114 Then
				'	Dim aRefreshInfo()
				'	lErrNumber = NO_ERR
				'	Redim Preserve aRefreshInfo(B_IS_PROMPTED_REFRESH + 1)
				'	aRefreshInfo(B_IS_PROMPTED_REFRESH) = True
				'	lErrNumber = RefreshObject(oRequest, aConnectionInfo, aRefreshInfo, aPageInfo, sErrDescription)
				'	Response.Redirect aRefreshInfo(S_REDIRECT_URL_REFRESH)
				'Else
					lErrNumber = ERR_GET_PROMPTQUESTION_FROMSERVER
					sErrDescription = asDescriptors(825) 'Descriptor: Error reading prompt definition from MicroStrategy Server. Please execute the report or document again.
					Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.Description), Err.source, "PromptCuLib.asp", "LoadPromptQuestionsXMLForJob", "", "Error in call to CreatePromptsHelperObject", LogLevelTrace)
				'End If
			End If
		End If
	End If

	If lErrNumber=NO_ERR Then
		lErrNumber = GetXMLDOM(aConnectionInfo, aPromptGeneralInfo(PROMPT_O_QUESTIONSXML), sErrDescription)
		If IsObject(aPromptGeneralInfo(PROMPT_O_QUESTIONSXML)) Then
		    aPromptGeneralInfo(PROMPT_O_QUESTIONSXML).loadXML (aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).QuestionXML)
		    if Err.Number <> 0 then
				lErrNumber = ERR_GET_PROMPTQUESTION_FROMSERVER
				Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptCuLib.asp", "LoadPromptQuestionsXMLForJob", "", "", LogLevelTrace)
				sErrDescription = asDescriptors(825) 'Descriptor: Error reading prompt definition from MicroStrategy Server. Please execute the report or document again.
			End If
		End If
	End If

	If lErrNumber=NO_ERR then
		if aPromptGeneralInfo(PROMPT_B_ISDOC) Then
		    Set aPromptGeneralInfo(PROMPT_O_QUESTIONPIFS) = aPromptGeneralInfo(PROMPT_O_QUESTIONSXML).selectNodes("/mi/in/oi[@tp='10']/mi/pif")
		Else
		    Set aPromptGeneralInfo(PROMPT_O_QUESTIONPIFS) = aPromptGeneralInfo(PROMPT_O_QUESTIONSXML).selectNodes("/mi/rit/rsl/mi/in/oi[@tp='10']/mi/pif")
		End If

		If aPromptGeneralInfo(PROMPT_O_QUESTIONPIFS).length > 0 Then	'go back in document, we get HTML execution result, no PIF
			lErrNumber = CreatePromptInfoArray(aConnectionInfo, aPromptGeneralInfo, aPromptInfo)
			If lErrNumber <> NO_ERR Then
			    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptCuLib.asp", "LoadPromptQuestionsXMLForJob", "", "Error in call to CreatePromptInfoArray", LogLevelTrace)
			End If
		End If
	End If

	If lErrNumber = NO_ERR Then	'handle go back to prompt page after execution
		lErrNumber = CheckAllClosedPrompts(aConnectionInfo, aPromptGeneralInfo, bAllClosed)
		If lErrNumber <> NO_ERR Then
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptCuLib.asp", "LoadPromptQuestionsXMLForJob", "", "Error in call to CheckAllClosedPrompts", LogLevelTrace)
		ElseIf bAllClosed Then
		    lErrNumber = ERR_NO_OPEN_PROMPTS
		    sErrDescription = asDescriptors(483) & " " &  asDescriptors(1112) 'Decriptor: The prompts for this report could not be displayed because they have already been answered. | If you wish to answer these prompts again, please click the Re-prompt link next to the report.
		End If
	End If

	LoadPromptQuestionsXMLForJob = lErrNumber
    Err.Clear
End Function

Function GetTempPromptAnswersXMLForJob(aConnectionInfo, aPromptGeneralInfo, aPromptInfo)
'***********************************************************************************************
'Purpose: Load the promptAnswersXML that we've created so far from a disk file into DOM object.
'Inputs:  aConnectionInfo, aPromptGeneralInfo(PROMPT_S_MSGID), aPromptInfo, aPromptGeneralInfo(PROMPT_O_QUESTIONSXML)
'Outputs: aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML), sErrDescription
'***********************************************************************************************
    On Error Resume Next
    Dim lPin
    Dim oRootXML
    Dim oRootMI
    Dim oSinglePrompt
    Dim oSinglePromptTempXML
    Dim sPromptsTempXML
	Dim lErrNumber

    Call GetXMLDOM(aConnectionInfo, aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML), sErrDescription)

    sPromptsTempXML = decodeXML(oRequest("nuXML_tempanswerxml"))
    If Len(sPromptsTempXML) > 0 Then
		Call aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML).loadXML(sPromptsTempXML)
		For lPin = 1 to aPromptGeneralInfo(PROMPT_L_MAXPIN)
			Set oSinglePromptTempXML = aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML).selectSingleNode("./mi/pif[@pin='" & CStr(lPin) & "']")
			Set aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER) = oSinglePromptTempXML
		Next

	Else
		'Create TempXML Frame
		Set oRootXML = aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML).selectSingleNode("/")
    	Set oRootMI = oRootXML.createElement("mi")
		Call aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML).appendChild(oRootMI)
	    For lPin = 1 to aPromptGeneralInfo(PROMPT_L_MAXPIN)
			Set oSinglePromptTempXML = oRootXML.createElement("pif")
			Call oSinglePromptTempXML.setAttribute("pin", CStr(lPin))
			Call oRootMI.appendChild(oSinglePromptTempXML)
			Set aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER) = oSinglePromptTempXML

			Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)
			If oSinglePrompt.PromptType = DssXmlPromptExpression Then
				Call TestUnknownDefault(aConnectionInfo, aPromptInfo, lPin, oSinglePrompt)
			End If
		Next
	End If

    GetTempPromptAnswersXMLForJob = lErrNumber
    Err.Clear
End Function

Function ProcessCancelButton(aConnectionInfo, aObjectInfo, oSession, aPromptGeneralInfo, oRequest, sErrDescription)
'***********************************************************************************************
'Purpose: Process cancel button on prompt page
'Inputs:  aConnectionInfo, aPromptGeneralInfo, oRequest, oSession
'Outputs: aPromptGeneralInfo, sErrDescription
'***********************************************************************************************
    On Error Resume Next

    Call DeleteCache(aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID), CStr(GetSessionID()))
    Call CloseCastorSession(aConnectionInfo)
    Response.Redirect "services.asp?folderID=" & aPromptGeneralInfo(PROMPT_S_FOLDERID)

    ProcessCancelButton = lErrNumber
    Err.Clear
End Function

Function CheckNeedProcess(aConnectionInfo, oRequest, aPromptGeneralInfo)
'***********************************************************************************************
'Purpose: Check from oRequest to see if need processing later
'Inputs:  aConnectionInfo, oRequest
'Outputs: aPromptGeneralInfo(PROMPT_B_NEEDPROCESS)
'History
'Author          Date		  Description
'Gregorio Parra  02/08/2001	  Removed code that was making the page to be processed even though
'							  user hadn't answered the prompts.
'***********************************************************************************************
	On Error Resume Next
	Dim oItem


	'Setting flag
	aPromptGeneralInfo(PROMPT_B_NEEDPROCESS) = bMFCBlaster

	'Process prompt if user has answer them
	If Not bMFCBlaster Then
		If Request.Form.Count > 0 Then
			aPromptGeneralInfo(PROMPT_B_NEEDPROCESS) = True
		End If
	End If

	If aPromptGeneralInfo(PROMPT_B_BACK) Then
		aPromptGeneralInfo(PROMPT_B_NEEDPROCESS) = False
	End If
	CheckNeedProcess = Err.number
	Err.Clear
End Function

Function LogTimeInfo(aConnectionInfo, sFile, sASPFunc, sAPIFunc, sComment)
'***********************************************************************************************
'Purpose: To log the different errors in XML format
'Inputs:  aConnectionInfo, sFile, sASPFunc, sAPIFunc, sComment
'Outputs: Err.Number
'***********************************************************************************************
    Dim oFso
    Dim oTs

	Set oFso = Server.CreateObject("Scripting.FileSystemObject")
    If IsObject(oFso) Then
        Set oTs = oFso.OpenTextFile("c:\temp\time.txt", 8, True)
        If IsObject(oTs) Then
			Dim x
			Call oTimeRecorder.gettime(x)
            oTs.WriteLine CStr(Now()) & ": " & x & ": " & sComment & " ( " & sFile & ":" & sASPFunc & ":" & sAPIFunc & " )"
            oTs.Close
            Set oTs = Nothing
        End If
        Set oFso = Nothing
    End If

	Set oTs = Nothing
    Set oFso = Nothing
End Function

Function GetPrompt(oRequest, aConnectionInfo, oSession, oObjServer, aObjectInfo, aFolderInfo, aPromptGeneralInfo, sErrDescription)
'***********************************************************************************************
' Purpose: Load all prompt information before displaying
' Inputs:  oRequest, aConnectionInfo
' Outputs: oSession, oObjServer, aObjectInfo, aFolderInfo, aPromptGeneralInfo, sErrDescription
'***********************************************************************************************
	On Error Resume Next

	Dim sURL
	Dim lErrNumber
	Dim sRedirectURL
	Dim aNCSubscriptionInfo
	Dim sReexecutedReport
	Dim sDisableSaveAs

	'Hydra specific variables
    Dim sPrefIDName
    Dim iPos

	'Clear error setting before loading prompt
	lErrNumber = NO_ERR
	sErrDescription = ""

	If Len(oRequest("HydraNext1")) > 0 Or Len(oRequest("HydraBack1")) > 0 Or Len(oRequest("HydraFinish1")) > 0 Then
        lErrNumber = AnswerPromptByProfile(aConnectionInfo, aPromptGeneralInfo, oRequest)
        If lErrNumber = NO_ERR Then
            If Len(oRequest("HydraNext1")) > 0 Then
                Response.Redirect "PostPrompt.asp?action=next" & "&subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & "&qoid=" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) & "&src=" & oRequest("src")
            ElseIf Len(oRequest("HydraBack1")) > 0 Then
                Response.Redirect "PostPrompt.asp?action=back" & "&subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & "&qoid=" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) & "&src=" & oRequest("src")
            ElseIf Len(oRequest("HydraFinish1")) > 0 Then
				If (aPromptGeneralInfo(PROMPT_B_NOSUBS) = "1") Then
					Response.Redirect "testDisplay.asp?action=finish&subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID)
				Else
					Response.Redirect "Modify_Subscription.asp?action=finish&subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID)
				End If
            End If
        Else
            'LogErrorXML( )
            'Exit Function
        End If
    ElseIf Len(oRequest("ProfileEdit")) > 0 Then
        sPrefIDName = CStr(oRequest("ProfileList"))
        iPos = InStr(1, sPrefIDName, ":", vbBinaryCompare)
        Response.Redirect "PrePrompt.asp?subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & "&qoid=" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) & "&prefID=" & Left(sPrefIDName, iPos - 1)
    ElseIf Len(oRequest("ProfileDelete")) > 0 Then
        sPrefIDName = CStr(oRequest("ProfileList"))
        iPos = InStr(1, sPrefIDName, ":", vbBinaryCompare)
        Response.Redirect "DeleteProfile.asp?subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & "&qoid=" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) & "&prefID=" & Left(sPrefIDName, iPos - 1)
    End If

	If aPromptGeneralInfo(B_TRIGGERS_ONLY_PROMPT) then
		'cancel button
		If aPromptGeneralInfo(PROMPT_B_CANCEL) And Len(aPromptGeneralInfo(PROMPT_S_MSGID)) > 0 Then
		    lErrNumber = ProcessCancelButton(aConnectionInfo, aObjectInfo, oSession, aPromptGeneralInfo, oRequest, sErrDescription)
		    If lErrNumber <> NO_ERR Then
		        Call ShowErrorMessage(lErrNumber)
		    End If
		End If
	Else
		lErrNumber = LoadPromptQuestionsXMLForJob(aConnectionInfo, aPromptGeneralInfo, oSession, aPromptInfo, oRequest, sErrDescription)
		If lErrNumber <> NO_ERR Then
		    If lErrNumber = ERR_GET_PROMPTQUESTION_FROMSERVER then
		    	sURL = GetGeneralParasinURL(oRequest)
		        Response.Redirect "JobError.asp?ErrNum=" & lErrNumber & "&ErrDesc=" & Server.URLEncode(CleanErrorMessage(sErrDescription)) & "&" & sURL
			ElseIf lErrNumber = ERR_NO_OPEN_PROMPTS Then
		        Call AnswerPromptByWidget(aConnectionInfo, oRequest, aPromptGeneralInfo)
                Call CloseCastorSession(aConnectionInfo)
                Response.Redirect "PostPrompt.asp?action=next" & "&subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & "&qoid=" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID)
		    End If
		End If
		If aPromptGeneralInfo(PROMPT_B_XML) Then
		    Response.Write "<!-- aPromptGeneralInfo(PROMPT_O_QUESTIONSXML).xml: " & aPromptGeneralInfo(PROMPT_O_QUESTIONSXML).xml & " -->"
		End If

		'Cancel this request
		If aPromptGeneralInfo(PROMPT_B_CANCEL) And Len(aPromptGeneralInfo(PROMPT_S_MSGID)) > 0 Then
		    lErrNumber = ProcessCancelButton(aConnectionInfo, aObjectInfo, oSession, aPromptGeneralInfo, oRequest, sErrDescription)
		    If lErrNumber <> NO_ERR Then
		        Call ShowErrorMessage(lErrNumber)
		    End If
		End If

		If lErrNumber = NO_ERR Then
			lErrNumber = GetTempPromptAnswersXMLForJob(aConnectionInfo, aPromptGeneralInfo, aPromptInfo)
		End If

		If lErrNumber = NO_ERR Then
			If aPromptGeneralInfo(PROMPT_B_NEEDPROCESS) Then
		        lErrNumber = ProcessAllPromptSelectionsForJob(aConnectionInfo, oSession, aPromptGeneralInfo, aPromptInfo, oRequest)
			End If
		End If


		If lErrNumber = NO_ERR Then
		    If aPromptGeneralInfo(PROMPT_B_SENDANSWER) Then
			    lErrNumber = SendPromptAnswersXMLForJob(aConnectionInfo, oRequest, aPromptGeneralInfo, aPromptInfo, oSession, sErrDescription)

				If lErrNumber <> NO_ERR Then
					If lErrNumber = ERR_ANSWERPROMPT Then
						lErrNumber = NO_ERR
						If aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) Then
							lErrNumber = LoadPromptQuestionsXMLForJob(aConnectionInfo, aPromptGeneralInfo, oSession, aPromptInfo, oRequest, sErrDescription)
						End If
					End If
				Else

					Dim oReport
					Set oReport = Server.CreateObject("WebAPIHelper.DSSXMLResultSet")
					oReport.SessionID = aConnectionInfo(S_TOKEN_CONNECTION)
					oReport.MessageID = aPromptGeneralInfo(PROMPT_S_MSGID)
					oReport.ExecutionFlags = DssXmlExecutionResolve or DssXmlDocExecutionResolve
					oReport.GetResults

					Dim oPrompts
					Set oPrompts = oReport.PromptsObject

					'Hydra
					Response.Clear
                    If Len(oRequest("HydraNext2")) > 0 Then
						If isObject(oPrompts) Then
							Response.Redirect "Prompt.asp?msgsaved=1&subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & "&qoid=" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) & "&sessionID=" & oReport.SessionID & "&msgid=" & oReport.MessageID
						Else
							Call CloseCastorSession(aConnectionInfo)
							Response.Redirect "PostPrompt.asp?action=next" & "&subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & "&qoid=" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID)
                    	End If
                    ElseIf Len(oRequest("HydraBack2")) > 0 Then
                        Call CloseCastorSession(aConnectionInfo)
                        Response.Redirect "PostPrompt.asp?action=back" & "&subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & "&qoid=" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID)
                    ElseIf Len(oRequest("HydraFinish2")) > 0 Then
                        Response.Redirect "PostPrompt.asp?action=finish&subGUID=" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & "&qoid=" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID)
                    End If

				End If
		    End If
		End If

		If Not aPromptGeneralInfo(PROMPT_B_SUMMARY) And (Len(Cstr(oRequest("subscribego"))) = 0 Or lErrNumber <> NO_ERR) Then
			lErrNumber = CreateDisplayXMLForAllPrompts(aConnectionInfo, oSession, aPromptInfo, aPromptGeneralInfo, oRequest, sErrDescription)
		End If

		Call CO_GetAnyPromptError(aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML), aPromptGeneralInfo(PROMPT_B_ANYERROR))

	End If

	GetPrompt = lErrNumber
	Err.Clear
End Function


Function Clean(aConnectionInfo)
'***********************************************************************************************
'Purpose: Clean up memory
'Inputs:  aConnectionInfo
'***********************************************************************************************
	On Error Resume Next

	Call CleanMem(oRequest, oSession, aPromptGeneralInfo, aiQuestions)
	Set aObjectInfo(O_CONTENTS_XML_OBJECT) = Nothing
	Set aFolderInfo(O_CONTENTS_XML_OBJECT) = Nothing
	Erase aObjectInfo
	Erase aFolderInfo

End function

Function ReceivePromptRequest(oRequest, aObjectInfo, aFolderInfo, aPromptGeneralInfo)
'***********************************************************************************************
'Purpose: Receive request parameters from oRequest
'Inputs:  oRequest
'Outputs: aPromptGeneralInfo, aObjectInfo, aFolderInfo
'***********************************************************************************************
    On Error Resume Next
    Dim sPos
    Dim oObjInbox
    Dim oAllPrompts
    Dim bSummary
    Dim bIndexButton
    Dim sReq
    Dim sTempPromptAnswersXML
    Dim sIndex
    Dim bFromSaveAsPage

	ReDim Preserve aObjectInfo(MAX_OBJECT_INFO)
	ReDim Preserve aFolderInfo(MAX_OBJECT_INFO)
	ReDim Preserve aPromptGeneralInfo(MAX_PROMPTGENERAL_INFO)


	'If oRequest.Exists("Split") Then
	'	Call ModifyRequestForSplitInputs(oRequest)
	'End If

	aPromptGeneralInfo(PROMPT_B_XML) = False

	If IsEmpty(oRequest("displaypin")) Then
		aPromptGeneralInfo(PROMPT_S_CURORDER) = "1"
	Else
		aPromptGeneralInfo(PROMPT_S_CURORDER) = CStr(oRequest("displaypin"))
	End If

	If not IsEmpty(oRequest("noSubs")) Then
		aPromptGeneralInfo(PROMPT_B_NOSUBS) = "1"
	Else
		aPromptGeneralInfo(PROMPT_B_NOSUBS) = "0"
	End If

	aPromptGeneralInfo(PROMPT_B_SAVE) = false
	aPromptGeneralInfo(PROMPT_B_SENDANSWER) = false
	aPromptGeneralInfo(PROMPT_S_BETWEENSEPERATOR) = ";"
	aPromptGeneralInfo(PROMPT_S_INSEPERATOR) = ";"
	aPromptGeneralInfo(PROMPT_B_REDIRECTTOREBUILD) = False

	If Len(oRequest("xsl")) > 0 And Len(oRequest("css")) > 0 Then
		aPromptGeneralInfo(PROMPT_B_REDIRECTTOREBUILD)	= True
	End If

	aPromptGeneralInfo(PROMPT_S_MSGID) = CStr(oRequest("msgid"))

    aPromptGeneralInfo(PROMPT_B_ISDOC) = (Len(CStr(oRequest("documentid"))) > 0)

    aPromptGeneralInfo(PROMPT_S_REPORTID) = CStr(oRequest("reportid"))
    aPromptGeneralInfo(PROMPT_S_FILTERID) = CStr(oRequest("filterid"))
    aPromptGeneralInfo(PROMPT_S_TEMPLATEID) = CStr(oRequest("templateid"))
    If (Len(aPromptGeneralInfo(PROMPT_S_REPORTID)) = 0) And (Len(aPromptGeneralInfo(PROMPT_S_FILTERID)) > 0) And (Len(aPromptGeneralInfo(PROMPT_S_TEMPLATEID)) > 0) Then
		aPromptGeneralInfo(PROMPT_S_REPORTID) = "-1"
    End If
	aPromptGeneralInfo(PROMPT_S_DOCUMENTID) = CStr(oRequest("documentid"))
    aPromptGeneralInfo(PROMPT_S_VIEWMODE) = CStr(oRequest("view"))

    If Not IsEmpty(oRequest("msgsaved")) Then
		aPromptGeneralInfo(PROMPT_B_MESSAGESAVED) = True
	Else
	'	Call CreateInboxHelperObject(aConnectionInfo, oObjInbox, sErrDescription)
	'	oObjInbox.MessageID = aPromptGeneralInfo(PROMPT_S_MSGID)
	'	aPromptGeneralInfo(PROMPT_B_MESSAGESAVED) = oObjInbox.IsSaved
	aPromptGeneralInfo(PROMPT_B_MESSAGESAVED) = False
	End If

	aPromptGeneralInfo(PROMPT_B_XML) = Not IsEmpty(oRequest("xml"))

	bSummary = Not (IsEmpty(oRequest("summary")))

	bIndexButton = False
	For each sReq in oRequest
		If Not IsEmpty(oRequest(sReq)) Then
			If StrComp(Left(sReq, 11), "PromptCurr_", vbTextCompare) = 0 Then
				bIndexButton = True
				sIndex = Mid(sReq, 12, Len(sReq)-13)
			End If
		End If
	Next

	If bSummary And (Not(IsEmpty(oRequest("promptback"))) Or bIndexButton) Then
		aPromptGeneralInfo(PROMPT_B_BACK) = true
		If bIndexButton Then
			aPromptGeneralInfo(PROMPT_S_CURORDER) = sIndex
		End If
	Else
		aPromptGeneralInfo(PROMPT_B_BACK) = false
	End If
	Call CheckNeedProcess(aConnectionInfo, oRequest, aPromptGeneralInfo)

	bFromSaveAsPage = Not IsEmpty(oRequest("FromSaveAsPage"))
	If aPromptGeneralInfo(PROMPT_B_NEEDPROCESS) Then
		If ((aPageInfo(N_ALIAS_PAGE) = DssXmlFolderNameTemplateReports) Or bFromSaveAsPage)And StrComp(CStr(oRequest("saveasbtn")), "", vbTextCompare) <> 0 Then
			aPromptGeneralInfo(PROMPT_B_SAVE) = True
		End If
		'If Len(CStr(oRequest("promptgo")))> 0 Or Len(CStr(oRequest("saveasbtn"))) > 0 Or Len(CStr(oRequest("subscribego"))) > 0  Then
		If Len(CStr(oRequest("promptGO"))) > 0 Or Len(CStr(oRequest("SaveAsBtn"))) > 0 Or Len(CStr(oRequest("subscribego"))) > 0 Or Len(CStr(oRequest("HydraBack2"))) > 0 Or Len(CStr(oRequest("HydraNext2"))) > 0 Or Len(CStr(oRequest("HydraFinish2"))) > 0 Then
			aPromptGeneralInfo(PROMPT_B_SENDANSWER) = True
			aPromptGeneralInfo(PROMPT_B_VALIDATE) = True
		End If
	End If

	aPromptGeneralInfo(PROMPT_B_REPROMPT) = (StrComp(CStr(oRequest("reprompt")), "1", vbBinaryCompare) = 0)
	aPromptGeneralInfo(PROMPT_B_CANCEL) = (Len(CStr(oRequest("cancel"))) > 0)
	aPromptGeneralInfo(PROMPT_B_EXECUTE) = (Len(CStr(oRequest("promptgo")))>0)

	'set DHTML option
'	aPromptGeneralInfo(PROMPT_B_DHTML) = True
'	If Not aPageInfo(B_USE_DHTML_PAGE) Then
'		aPromptGeneralInfo(PROMPT_B_DHTML) = False
'	End If

	If Not(IsEmpty(oRequest("userselect"))) Then
		aPromptGeneralInfo(PROMPT_B_DHTML) = True
	End If

	'set prompts display order option
	aPromptGeneralInfo(PROMPT_B_REQUIREDFIRST) = (ReadUserOption(REQUIRED_PROMPTS_FIRST_OPTION) = "CHECKED")

	'set Display All Prompts in One page
	'aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) = (ReadUserOption(PROMPTS_ON_ONE_PAGE_OPTION) = "1" )

	aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) = True

	'Hydra
	Erase asDescriptors
	asDescriptors = asHydraDescriptors
    lErrNumber = GetHydraPrompt(aPromptGeneralInfo)

    Erase asDescriptors
	asDescriptors = asWebDescriptors

    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.Source, "PromptCuLib.asp", "ReceivePromptRequest", "", "Error after calling GetHydraPrompt", LogLevelTrace)
    End If

	If lErrNumber = NO_ERR Then
		If Not IsEmpty(oRequest("page")) Then
			aPageInfo(N_ALIAS_PAGE) = CStr(oRequest("page"))
			sPos = InStr(1, aPageInfo(N_ALIAS_PAGE), "#", vbBinaryCompare)
			If sPos > 0 Then
			    aPageInfo(N_ALIAS_PAGE) = Left(aPageInfo(N_ALIAS_PAGE), sPos - 1)
			End If
			aPageInfo(N_ALIAS_PAGE) = CLng(aPageInfo(N_ALIAS_PAGE))
		End If

	    If aPromptGeneralInfo(PROMPT_B_ISDOC)  Then
			aObjectInfo(S_OBJECT_ID_OBJECT) = aPromptGeneralInfo(PROMPT_S_DOCUMENTID)
			aObjectInfo(L_TYPE_OBJECT) = DssXmlTypeDocumentDefinition
		Else
			aObjectInfo(S_OBJECT_ID_OBJECT) = aPromptGeneralInfo(PROMPT_S_REPORTID)
			aObjectInfo(L_TYPE_OBJECT) = DssXmlTypeReportDefinition
		End If
		aObjectInfo(S_TARGET_PAGE_OBJECT) = "Folder.asp"
		aFolderInfo(S_TARGET_PAGE_OBJECT) = "Folder.asp"
		aFolderInfo(N_NUMBER_OF_FOLDERS_TO_SHOW_OBJECT) = 0

		aPromptGeneralInfo(PROMPT_B_SUMMARY) = (Len(CStr(oRequest("promptsummary.x"))) > 0 Or Len(CStr(oRequest("promptsummary.y"))) > 0)


		If Not(IsEmpty(oRequest("nuXML_AnswerXml"))) Then
			aPromptGeneralInfo(PROMPT_S_ANSWERSXML) = DecodeXML(oRequest("nuXML_AnswerXml"))
		Else
			aPromptGeneralInfo(PROMPT_S_ANSWERSXML) = ""
		End If

		If IsEmpty(aPromptGeneralInfo(B_ADD_SUBSCRIPTION_PROMPT)) Then
			aPromptGeneralInfo(B_ADD_SUBSCRIPTION_PROMPT) = Not IsEmpty(oRequest("addsubscription"))
		End If
		If IsEmpty(aPromptGeneralInfo(B_EDIT_SUBSCRIPTION_PROMPT)) Then
			aPromptGeneralInfo(B_EDIT_SUBSCRIPTION_PROMPT) = Not IsEmpty(oRequest("editsubscription"))
			if IsEmpty(oRequest("oldtriggerid")) then
				aPromptGeneralInfo(S_OLD_TRIGGER_ID_PROMPT) = Cstr(oRequest("triggerid"))
			else
				aPromptGeneralInfo(S_OLD_TRIGGER_ID_PROMPT) = Cstr(oRequest("oldtriggerid"))
			End If
		End If
		If IsEmpty(aPromptGeneralInfo(B_TRIGGERS_ONLY_PROMPT)) Then
			aPromptGeneralInfo(B_TRIGGERS_ONLY_PROMPT) =  Not IsEmpty(oRequest("triggersonly"))
		End If
		If IsEmpty(aPromptGeneralInfo(S_TRIGGER_ID_PROMPT)) Then
			aPromptGeneralInfo(S_TRIGGER_ID_PROMPT) = oRequest("triggerid")
		End If

		'Narrowcast Integration:

		aPromptGeneralInfo(PROMPT_B_USE_NC) = Not IsEmpty(oRequest("NCact"))

	'	aPromptGeneralInfo(B_DISPLAY_TRIGGER_PROMPT) = aPromptGeneralInfo(PROMPT_B_USE_NC)
	'	aPromptGeneralInfo(B_DISPLAY_SUBSCRIBE_BUTTON) = aPromptGeneralInfo(PROMPT_B_USE_NC)
	'	If (aPromptGeneralInfo(B_ADD_SUBSCRIPTION_PROMPT) Or aPromptGeneralInfo(B_EDIT_SUBSCRIPTION_PROMPT)) Then
	'		aPromptGeneralInfo(B_DISPLAY_TRIGGER_PROMPT) = true
	'		aPromptGeneralInfo(B_DISPLAY_SUBSCRIBE_BUTTON) = True
	'	End If

		aPromptGeneralInfo(B_DISPLAY_TRIGGER_PROMPT) = False
		aPromptGeneralInfo(B_DISPLAY_SUBSCRIBE_BUTTON) = False

		If IsEmpty(aPromptGeneralInfo(PROMPT_B_REEXECUTED)) Then
			aPromptGeneralInfo(PROMPT_B_REEXECUTED) = Not IsEmpty(oRequest("Reexecuted"))
		End If
		If IsEmpty(aPromptGeneralInfo(PROMPT_B_DISABLE_SAVE)) Then
			If Not IsEmpty(oRequest("PromptGO")) Then
				aPromptGeneralInfo(PROMPT_B_DISABLE_SAVE) = True
			Else
				aPromptGeneralInfo(PROMPT_B_DISABLE_SAVE) = Not IsEmpty(oRequest("DisableSave"))
			End If
		End If
	End If
	ReceivePromptRequest = lErrNumber
    Err.Clear
End Function

Function CancelRequest(sMsgID, aConnectionInfo, oSession, sErrDescription)
'******************************************************************************
'Purpose: To cancel the given message
'Inputs:  sMsgID, aConnectionInfo, oSession
'Outputs: sErrDescription, Err.number
'******************************************************************************
    On Error Resume Next
	Dim oInboxObject
	Dim lErrNumber

	lErrNumber = CreateInboxHelperObject(aConnectionInfo, oInboxObject, sErrDescription)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "PromptCuLib.asp", "CancelRequest", "", "Error after calling CreateInboxHelperObject", LogLevelTrace)
	Else
		oInboxObject.MessageID = sMsgID
		Call oInboxObject.Remove()
		If lErrNumber <> NO_ERR Then
			lErrNumber = Err.number
			sErrDescription = asDescriptors(ERR_UNEXPECTED) & " " & CStr(Err.description) 'Descriptor: MicroStrategy Server error:
			Call LogErrorXML(aConnectionInfo, Cstr(lErrNumber), sErrDescription, Err.source, "PromptCuLib.asp", "CancelRequest", "InboxObject.Remove", "Error in call to InboxObject.Remove() function", LogLevelError)
		End If
	End If

	set oInboxObject = nothing
	CancelRequest = Err.number
	Err.Clear
End Function

Function TestUnknownDefault(aConnectionInfo, aPromptInfo, lPin, oSinglePrompt)
'***************************************************************************************************
'Purpose:   Add default expression prompt answer (from oSinglePromptQuestionXML) to oSinglePromptAnswerOI
'Inputs:    aConnectionInfo, oSinglePromptQuestionXML
'Outputs:   oSinglePromptAnswerOI
'***************************************************************************************************
    On Error Resume Next
    Dim oDefaultEXP
    Dim oRootND
    Dim oNode
    Dim bUnknownOperator
	Dim lResult
	Dim oEXP
	Dim oOP
	Dim lIndex
	Dim bUnknownDef
	Dim lDimensionality

	Set oDefaultEXP = oSinglePrompt.ExpressionObject
	bUnknownDef = False
	lDimensionality = 1

	If not (oDefaultEXP is nothing) then
		set oRootND = oDefaultEXP.RootNode
		if oRootND.ExpressionType = DssXmlFilterBranchQual And oRootND.ChildCount = 1 then
			set oRootND = oRootND.Child(1)
		End If

		Select Case oRootND.ExpressionType
		Case DssXmlFilterMetricExpression, _
			DssXmlFilterSingleBaseFormExpression, _
			DssXmlFilterJointFormQual,	 _
			DssXmlFilterJointListFormQual, _
			DssXmlFilterJointListQual, _
			DssXmlFilterMultiBaseFormQual, _
			DssXmlFilterMultiMetricQual
				bUnknownDef = True		'un-supported
		Case DssXmlFilterBranchQual
			if aPromptInfo(lPin, PROMPTINFO_B_ISCART)  then
				for lIndex = 1 to oRootND.ChildCount
					Set oNode = oRootND.Child(lIndex)
					If oNode.ExpressionType = DssXmlFilterBranchQual Then	'mixed AND/OR
						bUnknownDef = True
						exit for
					Else
						Call CO_TestUnknownOperator(oSinglePrompt.ExpressionType, oNode, bUnknownOperator)
						if bUnknownOperator Then
							bUnknownDef = True
							lDimensionality = oNode.DimensionalityType
							exit for
						End If
					End If
				next
			Else
				If oRootND.ChildCount > 0 Then
					bUnknownDef = True
				End If
			End If
		Case Else
			Call CO_TestUnknownOperator(oSinglePrompt.ExpressionType, oRootND, bUnknownOperator)
			if bUnknownOperator Then
				bUnknownDef = True
			End If
		End Select
    End If

	Call CO_SetbUnknownDef(aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER), bUnknownDef)
	if bUnknownDef then
		Call CO_SetDisplayUnknownDef(aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER), "1")
		If lDimensionality  = DssXmlNodeDimtyOutputLevel Then
			If oNode.DimtyObject.Count = 0 Then
				Call oSinglePrompt.SetPreviousAsAnswer()
			End If
		ElseIf oSinglePrompt.ExpressionType <> DssXmlFilterAttributeIDQual Then
			Call oSinglePrompt.SetDefaultAsAnswer()
		End If
	else
		Call CO_SetDisplayUnknownDef(aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER), "0")
	End If

    set oDefaultEXP = nothing
    set oRootND = nothing
    set oNode = nothing

    TestUnknownDefault = Err.Number
    Err.Clear
End Function

Function CO_TestUnknownOperator(lExpType, oNode, bUnknownOperator)
'***************************************************************************************************
'Purpose:   check unknown operator in node expression
'Inputs:    lExpType, oNode
'Outputs:   bUnknownOperator
'***************************************************************************************************
    On Error Resume Next
    Dim lOP
    Dim lOPType
    Dim lMRPOP

    lOP = oNode.Operator
    lMRPOP = oNode.MRPOperator
    lOPType = oNode.OperatorType

	bUnknownOperator = True

	If lOPType = DssXmlOperatorGeneric Then
		select case lOP
		case DssXmlFunctionBetween, DssXmlFunctionNotBetween, DssXmlFunctionEquals, DssXmlFunctionNotEqual, _
			 DssXmlFunctionGreater, DssXmlFunctionGreaterEqual, DssXmlFunctionLess, DssXmlFunctionLessEqual, _
			 DssXmlFunctionLike, DssXmlFunctionNotLike
			 If oNode.DimensionalityType  = DssXmlNodeDimtyOutputLevel Then
				If oNode.DimtyObject.Count = 0 Then
					bUnknownOperator = False
				End If
			 Else
				bUnknownOperator = False
			 End If
		case DssXmlFunctionIn, DssXmlFunctionNotIn
			if lExpType = DssXmlFilterAllAttributeQual then	'HI prompt
				bUnknownOperator = False
			Else
				bUnknownOperator = (oNode.ExpressionType = DssXmlFilterListQual)		'AQ/MQ
			End If
		end select
	Else
		If oNode.DimensionalityType  = DssXmlNodeDimtyOutputLevel Then
			If oNode.DimtyObject.Count = 0 Then
				bUnknownOperator = False
			End If
		Else
			Select Case lMRPOP
				Case DssXmlMRPFunctionTop, DssXmlMRPFunctionBottom
					bUnknownOperator = False
				Case Else
					bUnknownOperator = True
			End Select
		End If
	End If

	CO_TestUnknownOperator = Err.number
	Err.Clear
End Function


Function CheckAllClosedPrompts(aConnectionInfo, aPromptGeneralInfo, bAllClosed)
'***************************************************************************************************
'Purpose:   check If all prompts in aPromptGeneralInfo(PROMPT_O_QUESTIONSXML) are closed
'Inputs:    aConnectionInfo, aPromptGeneralInfo(PROMPT_B_ISDOC), aPromptGeneralInfo(PROMPT_O_QUESTIONPIFS)
'Outputs:   bAllClosed
'***************************************************************************************************
    On Error Resume Next
    Dim oSinglePrompt
    Dim lPin

	aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) = 0

	For lPin = 1 to aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Count
		set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)
		If oSinglePrompt.Used And not oSinglePrompt.Closed then
			aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) = aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) + 1
		End If
	Next
	bAllClosed = Not (aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) > 0)

    CheckAllClosedPrompts = Err.Number
    Err.Clear
End Function

Function ClosePromptsinAnswerXML(aConnectionInfo, oAnswer)
'***************************************************************************************************
'Purpose:   close all open prompts in oAnswer, And Set <pa ip="1"> to <pa ia="1">
'Inputs:    aConnectionInfo, oAnswer
'Outputs:   oAnswer
'***************************************************************************************************
    On Error Resume Next
    Dim oPif
    Dim oPA

	for each oPif in oAnswer.selectNodes("in/oi[@tp='10']/mi/pif")
		Call oPif.setAttribute("cl", "1")
		Set oPA = oPif.selectSingleNode("pa[@ip='1']")
		If Not(oPA is Nothing) Then
			Call oPA.removeAttribute("ip")
			Call oPA.setAttribute("ia", "1")
		End If
	next

	Set oPif = Nothing
	Set oPA = Nothing
	ClosePromptsinAnswerXML = Err.number
	Err.Clear
End Function

Function ClosePromptsinAnswerXMLString(aConnectionInfo, sAnswerXML)
'***************************************************************************************************
'Purpose:   close all open prompts in sAnswer, Set <pa ip="1"> to <pa ia="1">
'Inputs:    aConnectionInfo, sAnswer
'Outputs:   oAnswer
'***************************************************************************************************
    On Error Resume Next
    Dim oPif
	Dim oAnswer
	Dim lErrNumber

	lErrNumber = GetXMLDOM(aConnectionInfo, oAnswer, sErrDescription)
	If lErrNumber = NO_ERR Then
		oAnswer.loadXML (sAnswerXML)
		For Each oPif In oAnswer.selectNodes("mi/in/oi[@tp='10']/mi/pif")
			Call oPif.setAttribute("cl", "1")
		Next
		sAnswerXML = oAnswer.xml
	Else
	    Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptCulib.asp", "ClosePromptsinAnswerXMLString", "", "Couldn't create XMLDOM Object", LogLevelError)
	End If

	Set oPif = Nothing
	set oAnswer = nothing

	ClosePromptsinAnswerXMLString = Err.number
	Err.Clear
End Function

Function CleanMem(oRequest, oSession, aPromptGeneralInfo, aiQuestions)
'***************************************************************************************************
'Purpose:   Set each object be Nothing
'Inputs:    oRequest, oSession, aPromptGeneralInfo(PROMPT_O_QUESTIONSXML), aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML)
'Outputs:
'***************************************************************************************************
    On Error Resume Next

    Set oRequest = Nothing
    Set oSession = Nothing
    Set aPromptGeneralInfo(PROMPT_O_QUESTIONSXML) = Nothing
    Set aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML)= Nothing
	Erase aiQuestions

    CleanMem = Err.Number
    Err.Clear
End Function

Function GetStepandMsg(aConnectionInfo, oSinglePromptQuestionXML, sPin, lType, lExpType, sMin, sMax, bRequired, sStep, sMsg)
'***************************************************************************************************
'Purpose:  create <increfetch> part of display XML for element prompt from oSinglePromptQuestionXML
'Input:     aConnectionInfo, oSinglePromptQuestionXML, sPin, lType, lMin, lMax, bRequired
'Output:    sStep, sMsg
'***************************************************************************************************
    On Error Resume Next
	Dim lMax
	Dim lMin


	Select Case lType
	Case DssXmlPromptString
		lMin = CLng(sMin)
		lMax = CLng(sMax)
		If (lMin<>0) And (lMax<>-1) Then
			If (lMin = lMax) Then
				sMsg = asDescriptors(1980) 'Descriptor: This prompt requires a value of exactly ## characters.
				sMsg = replace(sMsg, "##", CStr(lMin))
			Else
				sMsg = asDescriptors(1977) 'Descriptor: This prompt requires a value between ## and ### characters.
				sMsg = replace(sMsg, "###", CStr(lMax))
				sMsg = replace(sMsg, "##", CStr(lMin))
			End If
		ElseIf (lMin<>0) Then
			sMsg = asDescriptors(1978) 'Descriptor: This prompt requires a value of no less than ## characters.
			sMsg = replace(sMsg, "##", CStr(lMin))
		ElseIf (lMax<>-1) Then
			sMsg = asDescriptors(1979) 'Descriptor: This prompt requires a value of no more than ## characters.
			sMsg = replace(sMsg, "##", CStr(lMax))
		Else
			sMsg = ""
		End If
	Case DssXmlPromptLong
		lMin = CLng(sMin)
		lMax = CLng(sMax)
		If (lMin<>0) And (lMax<>-1) Then
			sMsg = asDescriptors(731) 'Descriptor: This prompt requires a value between ## and ###.
			sMsg = replace(sMsg, "###", CStr(lMax))
			sMsg = replace(sMsg, "##", CStr(lMin))
		ElseIf (lMin<>0) Then
			sMsg = asDescriptors(732) 'Descriptor: This prompt requires a value no less than ##.
			sMsg = replace(sMsg, "##", CStr(lMin))
		ElseIf (lMax<>-1) Then
			sMsg = asDescriptors(733) 'Descriptor: This prompt requires a value no more than ##.
			sMsg = replace(sMsg, "##", CStr(lMax))
		Else
			sMsg = ""
		End If
	Case DssXmlPromptDouble
		If (sMin<>"0") And (sMax<>"-1") Then
			sMsg = asDescriptors(731) 'Descriptor: This prompt requires a value between ## and ###.
			sMsg = replace(sMsg, "###", sMax)
			sMsg = replace(sMsg, "##", sMin)
		ElseIf (sMin<>"0") Then
			sMsg = asDescriptors(732) 'Descriptor: This prompt requires a value no less than ##.
			sMsg = replace(sMsg, "##", sMin)
		ElseIf (sMax<>"-1") Then
			sMsg = asDescriptors(733) 'Descriptor: This prompt requires a value no more than ##.
			sMsg = replace(sMsg, "##", sMax)
		Else
			sMsg = ""
		End If
	Case DssXmlPromptDate
		If (sMin<>"0") And (sMax<>"-1") Then
			sMsg = asDescriptors(731) 'Descriptor: This prompt requires a value between ## and ###.
			sMsg = replace(sMsg, "###", sMax)
			sMsg = replace(sMsg, "##", sMin)
		ElseIf (sMin<>"0") Then
			sMsg = asDescriptors(735) 'Descriptor: This prompt requires a value no earlier than ##.
			sMsg = replace(sMsg, "##", sMin)
		ElseIf (sMax<>"-1") Then
			sMsg = asDescriptors(734) 'Descriptor: This prompt requires a value no later than ##.
			sMsg = replace(sMsg, "##", sMax)
		Else
			sMsg = ""
		End If
		sMin = "-1"
		sMax = "-1"
	Case Else
		lMin = CLng(sMin)
		lMax = CLng(sMax)
		If (lMin = -1) And (lMax = -1) And Not(bRequired) then
			sMsg = ""
		ElseIf bRequired And (lMin < 1) then
			lMin = 1
		End If

		If (lMin = -1) And (lMax <> -1) Then
			If lMax = 1 Then
				sMsg = asDescriptors(988) 'Descriptor: This prompt cannot accept more than 1 selection.
			Else
				sMsg = asDescriptors(918) 'Descriptor: This prompt cannot accept more than ## selections.
				sMsg = replace(sMsg, "##", CStr(lMax))
			End If
		ElseIf (lMax = -1) And (lMin <> -1) Then
			If lMin = 1 Then
				sMsg = asDescriptors(987) 'Descriptor: This prompt requires at least 1 selection.
			ElseIf lMin = 0 Then
				sMsg = asDescriptors(2456) 'Descriptor: No answer is required for this Prompt.
			Else
				sMsg = asDescriptors(917) 'Descriptor: This prompt requires at least ## selections.
				sMsg = replace(sMsg, "##", CStr(lMin))
			End If
		ElseIf (lMax <> -1) And (lMin <> -1) And (lMin < lMax) Then
			sMsg = asDescriptors(916) 'Descriptor: This prompt requires between ## and ### selections.
			sMsg = replace(sMsg, "###", CStr(lMax))
			sMsg = replace(sMsg, "##", CStr(lMin))
		ElseIf (lMax <> -1) And (lMin <> -1) And (lMin = lMax) Then
			If lMin = 1 Then
				sMsg = asDescriptors(968) 'Descriptor: This prompt requires only 1 selection.
			Else
				sMsg = asDescriptors(936) 'Descriptor: This prompt requires exactly ## selections.
				sMsg = replace(sMsg, "##", CStr(lMin))
			End If
		End If

		If lExpType = DssXmlFilterAllAttributeQual And Len(sMsg)>0 then
			sMsg = sMsg & " " & asDescriptors(1002) 'Descriptor: An expression or a group of elements from one attribute is equivalent to one selection.
		End If

	End Select

    'create step information
    sStep = ""
    If bRequired or lMin > 0 Then
		sStep = " (" & asDescriptors(661) & ")" 'Descriptor: Required
    End If

   	GetStepandMsg = Err.Number
    Err.Clear
End Function

Function CreatePromptInfoArray(aConnectionInfo, aPromptGeneralInfo, aPromptInfo)
'***************************************************************************************************
' Purpose:	Create prompt info array to be used throughout prompt page
' Inputs:	aConnectionInfo, aPromptGeneralInfo(PROMPT_O_QUESTIONPIFS)
' Outputs:	aPromptGeneralInfo(PROMPT_L_MAXPIN), aPromptInfo
'***************************************************************************************************
	On Error Resume Next
	Dim oSinglePromptQuestionXML
	Dim oLastPIF
    Dim sPin
    Dim lPin
    Dim sStyle
    Dim sXSL
    Dim lType
    Dim lExpType
    Dim oRes
    Dim bRequired
    Dim oMin
    Dim oMax
    Dim sMin
    Dim sMax
    Dim sStep
    Dim sMsg
    Dim sErrCode
	Dim bIsCart
	Dim oORDM
	Dim oSinglePrompt

	Set oLastPIF = aPromptGeneralInfo(PROMPT_O_QUESTIONPIFS)(aPromptGeneralInfo(PROMPT_O_QUESTIONPIFS).length-1)
	aPromptGeneralInfo(PROMPT_L_MAXPIN) = CLng(oLastPIF.getAttribute("pin"))
	if isempty(aPromptInfo) then
		Redim aPromptInfo(aPromptGeneralInfo(PROMPT_L_MAXPIN), MAX_PROMPTINFO_S_INDEX)
	else
		Redim Preserve aPromptInfo(aPromptGeneralInfo(PROMPT_L_MAXPIN), MAX_PROMPTINFO_S_INDEX)
	End If

	aPromptGeneralInfo(PROMPT_B_ANY_TEXTFILE) = False

	For Each oSinglePromptQuestionXML In aPromptGeneralInfo(PROMPT_O_QUESTIONPIFS)
		sPin = oSinglePromptQuestionXML.getAttribute("pin")
		lPin = CLng(sPin)
		Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)

		If Not (oSinglePrompt.HasAnswer) Then
			If oSinglePrompt.HasPreviousAnswer Then
				Call oSinglePrompt.setPreviousAsAnswer()
			ElseIf oSinglePrompt.HasDefaultAnswer Then
				Call oSinglePrompt.setDefaultAsAnswer()
			End If
		End If

		lType = oSinglePrompt.PromptType
		lExpType = oSinglePrompt.ExpressionType

		Call CO_GetPromptStyleFile(oSinglePrompt, oSinglePromptQuestionXML, sXSL)

		Call CO_GetPromptCartProperty(oSinglePrompt, oSinglePromptQuestionXML, bIsCart)

		Call GetStepandMsg(aConnectionInfo, oSinglePromptQuestionXML, sPin, lType, lExpType, oSinglePrompt.Min, oSinglePrompt.Max, oSinglePrompt.Required, sStep, sMsg)

		'Hydra
        If lType = DssXmlPromptElements And aPromptGeneralInfo(PROMPT_B_CHANGE_STYLE) Then
			Select Case lCase(sXSL)
			case "promptelement_checkbox.xsl"
				sXSL = "PromptElement_radio.xsl"
			case "promptelement_cart.xsl", "promptelement_multiselect_listbox.xsl"
				sXSL = "PromptElement_SingleSelect_listbox.xsl"
			End Select
			bIsCart = False
		End If

		Set aPromptInfo(lPin, PROMPTINFO_O_QUESTION) = oSinglePromptQuestionXML
		aPromptInfo(lPin, PROMPTINFO_S_INDEX) = sPin
		aPromptInfo(lPin, PROMPTINFO_S_XSLFILE) = sXSL
		aPromptInfo(lPin, PROMPTINFO_B_ISCART) = bIsCart
		aPromptInfo(lPin, PROMPTINFO_S_STEP) = sStep
		aPromptInfo(lPin, PROMPTINFO_S_MSG) = sMsg

		If strcomp(sXSL, "promptexpression_textfile.xsl", vbTextCompare) = 0 Then
			aPromptGeneralInfo(PROMPT_B_ANY_TEXTFILE) = True
		End If

		If lExpType = DssXmlFilterAllAttributeQual Then
			Set oORDM = oSinglePromptQuestionXML.selectSingleNode("./or/dm")
			aPromptInfo(lPin, PROMPTINFO_B_ISALLDIMENSION) = (oORDM Is Nothing)      'prompt on all dimension
		End If
	Next

	set oORDM = nothing

	CreatePromptInfoArray = Err.number
	Err.Clear
End Function


Function SendPromptAnswersXMLForJob(aConnectionInfo, oRequest, aPromptGeneralInfo, aPromptInfo, oSession, sErrDescription)
'***************************************************************************************************
'Purpose:   Validate and Answer the prompts
'Inputs:    aConnectionInfo, oRequest, aPromptGeneralInfo, aPromptInfo, oSession
'Outputs:   sErrDescription
'***************************************************************************************************
    On Error Resume Next
    Dim lErrNumber

	lErrNumber = ValidatePromptAnswers(aConnectionInfo, aPromptGeneralInfo, aPromptInfo, sErrDescription)
	If lErrNumber<>NO_ERR then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptCuLib.asp", "SendPromptAnswersXMLForJob", "ValidatePromptAnswers", "Error in call to ValidatePromptAnswers", LogLevelTrace)
	Else
		If aPromptGeneralInfo(PROMPT_L_ISM_TYPE) = ISM_TYPE_CASTOR Then   'Hydra
			Call aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Answer()
			If Err.number<>NO_ERR then
				lErrNumber = ERR_ANSWERPROMPT
				sErrDescription = asDescriptors(899)	'Err.Description 'Error when answering the prompt. Please review your answer(s).
				Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptCuLib.asp", "SendPromptAnswersXMLForJob", "aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Answer()", "Error in call to oRepServer.AnswerPrompt", LogLevelError)
			End If
		End If
		If lErrNumber = NO_ERR Then		'Hydra
			Erase asDescriptors
			asDescriptors = asHydraDescriptors

			lErrNumber = AnswerPromptByWidget(aConnectionInfo, oRequest, aPromptGeneralInfo)

			Erase asDescriptors
			asDescriptors = asWebDescriptors

			If lErrNumber <> NO_ERR then
			   Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptCoLib.asp", "SendPromptAnswersXMLForJob", "", "Error in call to AnswerPromptByWidget", LogLevelTrace)
			End If
		End If
	End If

    SendPromptAnswersXMLForJob = lErrNumber
    Err.Clear
End Function

Function ValidatePromptAnswers(aConnectionInfo, aPromptGeneralInfo, aPromptInfo, sErrDescription)
'*************************************************************************************************************
'Purpose:	validate prompt answers
'Inputs:	aConnectionInfo, aPromptGeneralInfo, sErrDescription
'Outputs:	sErrDescription
'*************************************************************************************************************
    On Error Resume Next
	Dim bSthWrong
	Dim sPin
	Dim sErrCode
	Dim lPin
	Dim oSinglePrompt
	Dim lType
	Dim lOrder
	Dim bFirst
	Dim oSinglePromptTempXML
	Dim lErrNumber

	bFirst = True
	bSthWrong = False
	Call CO_RemoveCurrentPin(aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML))

	For lOrder = 1 to aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT)
		Call GetPinbyOrder(aPromptGeneralInfo, aPromptInfo, lOrder, lPin)
		Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)
		if oSinglePrompt.Used and not oSinglePrompt.Closed then
			lType = oSinglePrompt.PromptType
			Set oSinglePromptTempXML = aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER)
			Call CO_GetPromptError(oSinglePromptTempXML, sErrCode)
			If sErrCode<>"0" Then		'If there is sth wrong in process, display it
				bSthWrong = true
			Else
				call oSinglePrompt.Validate
				lErrNumber = Err.number
				if lErrNumber<>0 then
					bSthWrong = true
					If bFirst Then
						aPromptGeneralInfo(PROMPT_S_CURORDER) = CStr(lOrder)
						bFirst = False
					End If
					Call CO_SetCurrentPin(oSinglePromptTempXML)

					Select case lType
						Case DssXmlPromptLong, DssXmlPromptString, DssXmlPromptDouble, DssXmlPromptDate
						Select case lErrNumber
						case HELPERERROR_PROMPT_REQUIRED
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_REQUIRED_PROMPT)
						case HELPERERROR_PROMPT_TOOMANY
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOLONG_TEXT_CONSTANTPROMPT)
						case HELPERERROR_PROMPT_TOOFEW
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOSHORT_TEXT_CONSTANTPROMPT)
					End Select

					Case DssXmlPromptObjects
						Select case lErrNumber
						case HELPERERROR_PROMPT_REQUIRED
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_REQUIRED_PROMPT)
						case HELPERERROR_PROMPT_TOOMANY
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOMANY_SELECTIONS_OBJECTPROMPT)
						case HELPERERROR_PROMPT_TOOFEW
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOFEW_SELECTIONS_OBJECTPROMPT)
						End Select

					Case DssXmlPromptElements
						Select case lErrNumber
						case HELPERERROR_PROMPT_REQUIRED
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_REQUIRED_PROMPT)
						case HELPERERROR_PROMPT_TOOMANY
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOMANY_SELECTIONS_ELEMENTPROMPT)
						case HELPERERROR_PROMPT_TOOFEW
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOFEW_SELECTIONS_ELEMENTPROMPT)
						End Select

					Case DssXmlPromptExpression
						Select case lErrNumber
						case HELPERERROR_PROMPT_REQUIRED
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_REQUIRED_PROMPT)
						case HELPERERROR_PROMPT_TOOMANY
							if oSinglePrompt.ExpressionType = DssXmlFilterAllAttributeQual then
								Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOMANY_SELECTIONS_HIPROMPT)
							else
								Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOMANY_SELECTIONS_EXPRESSIONPROMPT)
							End If
						case HELPERERROR_PROMPT_TOOFEW
							if oSinglePrompt.ExpressionType = DssXmlFilterAllAttributeQual then
								Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOFEW_SELECTIONS_HIPROMPT)
							else
								Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOFEW_SELECTIONS_EXPPROMPT)
							End If
						End Select

					Case DssXmlPromptDimty
						Select case lErrNumber
						case HELPERERROR_PROMPT_REQUIRED
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_REQUIRED_PROMPT)
						case HELPERERROR_PROMPT_TOOMANY
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOMANY_SELECTIONS_LEVELPROMPT)
						case HELPERERROR_PROMPT_TOOFEW
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOFEW_SELECTIONS_LEVELPROMPT)
						End Select

					Case Else
						Err.Raise ERR_CUSTOM_UNKNOWN_PROMPT_TYPE
						Call LogErrorXML(aConnectionInfo, Cstr(Err.number), CStr(Err.description), Err.source, "PromptCuLib.asp", "ValidatePromptAnswers", "", "Unknown Prompt Type", LogLevelError)
					End Select
				End If
			End If
		End If
	next

	If bSthWrong Then
		ValidatePromptAnswers = ERR_VALIDATION_FAILED
	Else
		ValidatePromptAnswers = Err.number
	End If

	Err.Clear
End Function

Function MapDTtoDDT(sDT, sDDT)
'***************************************************************
'Purpose:   Get Exp Item Text for AQ / MQ
'Inputs:    aConnectionInfo, oExpItem, sFlag
'Outputs:   sExpItemText
'***************************************************************
On Error Resume Next
	Dim lDDT

	Select Case Clng(sDT)
		case DssXmlBaseFormDateTime
			lDDT = DssXmlDataTypeTimeStamp
		case DssXmlBaseFormDate
			lDDT = DssXmlDataTypeDate
		case DssXmlBaseFormTime
			lDDT = DssXmlDataTypeTime
		Case DssXmlBaseFormNumber
			lDDT = DssXmlDataTypeReal
		Case DssXmlBaseFormText
			lDDT = DssXmlDataTypeChar
		Case Else
			lDDT = DssXmlDataTypeChar
	End Select
	sDDT = CStr(lDDT)

    MapDTtoDDT = Err.number
    Err.Clear
End Function

Function SetDefaultArrayforJS(aPromptGeneralInfo, aPromptInfo)
'***************************************************************
'Purpose:   Set bDefault Array values
'Inputs:
'Outputs:
'***************************************************************
	On Error Resume Next
	Dim i
	Dim oSinglePrompt
	Dim oSinglePromptTempXML
	Dim bUnknownDef

	For i = 1 to aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT)
		Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(i)
		if oSinglePrompt.Used and not oSinglePrompt.Closed then
			set oSinglePromptTempXML = aPromptInfo(i, PROMPTINFO_O_TEMPANSWER)
			Call CO_GetbUnknownDef(oSinglePromptTempXML, bUnknownDef)
			if bUnknownDef then
				Response.Write "bDefault[" & i & "] = true; "
			else
				Response.Write "bDefault[" & i & "] = false; "
			End If
		End If
	Next

	SetDefaultArrayforJS = Err.number
	Err.Clear
End Function

Function CheckFileAgainstAdminPreferences(aConnectionInfo, oRequest, sInputFileName, sErrDescription)
	On Error Resume Next
	Dim bValidFile
	Dim lErrNumber
	Dim sFileExtension

	If InStr(1, oRequest(sInputFileName & "_ext"), ".",  vbTextCompare)> 0 Then
		sFileExtension = Mid(oRequest(sInputFileName & "_ext"), Len(oRequest(sInputFileName & "_ext"))-2)
	Else
		sFileExtension = oRequest(sInputFileName & "_ext")
	End If

	bValidFile = False
	If InStr(1, "," & Replace(Replace(ReadUserOption(ALLOWED_FILE_EXTENSION_OPTION), " ", ""), ".", "") & ",", "," & sFileExtension & ",", vbTextCompare) = 0 Then
		Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.description), Err.source, "PromptCuLib.asp", "CheckFileAgainstAdminPreferences", "", "The extension of the file specified by the user was not allowed by the Adminstrator", LogLevelError)
		sErrDescription = Replace(asDescriptors(1261), "##", ReadUserOption(ALLOWED_FILE_EXTENSION_OPTION)) 'Descriptor: The file contains an invalid extension. The file extensions allowed are: ##
	Else
		If Len(oRequest(sInputFileName)) > (CInt(ReadUserOption(MAXIMUM_FILE_SIZE_TO_UPLOAD_OPTION)) * 1024) Then
			Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.description), Err.source, "PromptCuLib.asp", "CheckFileAgainstAdminPreferences", "", "The file specified by the user exceeded the maximum size allowed by the Administrator", LogLevelError)
			sErrDescription = Replace(asDescriptors(1312), "##", ReadUserOption(MAXIMUM_FILE_SIZE_TO_UPLOAD_OPTION)) 'Descriptor: You specified a file that exceeds the maximum size (## Kb) allowed by the Administrator.
		Else
			If Len(oRequest(sInputFileName)) = 0 Then
				Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.description), Err.source, "PromptCuLib.asp", "CheckFileAgainstAdminPreferences", "", "There user did not submit a file to the Web Server", LogLevelError)
				sErrDescription = asDescriptors(1259) 'Descriptor: There was not a file submitted to the Web Server.
			Else
				bValidFile = True
			End If
		End If
	End If

	CheckFileAgainstAdminPreferences = bValidFile
	Err.Clear
End Function

Function MaxNumberOfJobs(lErrNumber)
'*******************************************************************************
'Purpose: To indicate if an error number is due to have exceeded the max number of messages
'Inputs:  lErrNumber
'Outputs: True/False
'*******************************************************************************
	On Error Resume Next
	MaxNumberOfJobs = (lErrNumber = ERR_MAX_JOBS_PER_USER_EXCEEDED Or lErrNumber = ERR_MAX_JOBS_PER_PROJECT_EXCEEDED)
	Err.Clear
End Function

Function ModifyRequestForSplitInputs(oRequest)
'*******************************************************************************
'Purpose: To modify the Request Dictionary Object for Split Inputs
'Inputs:  oRequest
'Outputs: None
'*******************************************************************************
	On Error Resume Next

	Dim aString,i,j
	Dim aNewString(),lTotalItems
	Dim sItem,sFormName, lLength, lSeparator
	Dim sEntireString
	Dim lTotalStrings

	aString = SplitRequest(oRequest("split"))

	lTotalStrings = 0

	lTotalStrings = Ubound(aString)
	For i = 0 to lTotalStrings

		sItem = aString(i)
		lLength = Len(sItem)
		lSeparator = Instr(1,sItem,"|")
		lTotalItems = CLng(Right(sItem,lLength - lSeparator))

		sFormName = Mid(sItem,1,lSeparator - 1)

		Redim aNewString(lTotalItems + 1)

		If Instr(1,sFormName,"nuXML_",vbTextCompare) > 0 Then
			For j = 0 to lTotalItems
				aNewString(j) = oRequest("nuXML_split_" & j & "_" & sFormName)
				Call oRequest.Remove("nuXML_split_" & j & "_" & sFormName)
			Next
		Else
			For j = 0 to lTotalItems
				aNewString(j) = oRequest("split_" & j & "_" & sFormName)
				Call oRequest.Remove("split_" & j & "_" & sFormName)
			Next
		End If
		sEntireString = Join(aNewString,"")
		Call oRequest.Remove(sFormName)
		If Err.number <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.description), Err.source, "PromptCuLib.asp", "ModifyRequestForSplitInputs", "", "Error in Removing from the Request Object", LogLevelError)
		End If
		Call oRequest.Add(sFormName,sEntireString)
		If Err.number <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.description), Err.source, "PromptCuLib.asp", "ModifyRequestForSplitInputs", "", "Error in Adding to the Request Object", LogLevelError)
		End If
		Erase aNewString
	Next

	Call oRequest.Remove("split")

	If Err.number <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.description), Err.source, "PromptCuLib.asp", "ModifyRequestForSplitInputs", "", "Error in Recreating the Split elements in the Request", LogLevelError)
	End If

	Err.Clear
End Function

Function GetHydraPrompt(aPromptGeneralInfo)
'******************************************************************************
'Purpose: Get Current Prompt info
'Inputs:  aPromptGeneralInfo
'Outputs: Err.Number
'******************************************************************************
    On Error Resume Next
    'Dim oHydraPrompts
    Dim oCurrQO
    Dim oReport
    Dim sHydraPromptsXML
    Dim bFreshPrompt
    Dim oInfoSource
    Dim sProjectID
    Dim sInfoSourceID
    Dim oAnswerContent
    Dim oAnswer
    Dim oSecurity
    Dim oPrompts
    Dim oElements
    Dim oSub
    Dim oISOI
    Dim oUserAuth
    Dim oUserAuthPR
    Dim sUserAuth
    Dim iBeginPos
    Dim iEndPos
    Dim oUserSecurity
    Dim oUserSecurityPR
    Dim sUserSecurity
    Dim bUserLevel
    Dim oDefProfile
    Dim oISMProgID
	Dim oQODefinition
	Dim lPin
	Dim oSinglePrompt
	Dim sUserID
	Dim sElementIDs
	Dim i

    'set DHTML option
    aPromptGeneralInfo(PROMPT_B_DHTML) = (CStr(GetJavaScriptSetting()) = "1")

    aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) = CStr(oRequest("subGUID"))
    aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) = CStr(oRequest("qoid"))
    aPromptGeneralInfo(PROMPT_S_SRC) = CStr(oRequest("src"))

    aConnectionInfo(S_IP_ADDRESS_CONNECTION) = CStr(Request.ServerVariables("REMOTE_ADDR"))

    lErrNumber = ReadPromptQuestionFromCache(aConnectionInfo, aPromptGeneralInfo)
    If lErrNumber <> NO_ERR Then
        sErrDescription = "Error reading cache file"
        lErrNumber = ERR_GET_HYDRA_PROMPT
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErrDescription), Err.Source, "PromptCuLib.asp", "GetHydraPrompt", "", "Error reading cache file", LogLevelTrace)
    End If

    If lErrNumber = NO_ERR Then
        Set oSub = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).selectSingleNode("/mi/sub")
        aPromptGeneralInfo(PROMPT_S_FOLDERID) = oSub.getAttribute("fid")
        aPromptGeneralInfo(PROMPT_S_STATUS_FLAG) = oSub.getAttribute("sf")

        Set oCurrQO = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).selectSingleNode("/mi/qos/mi/in/oi[@tp='" & TYPE_QUESTION & "' $and$ @id='" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) & "']")

		aPromptGeneralInfo(PROMPT_B_ALLOW_PROFILE) = True
		If CLng(oCurrQO.getAttribute("qtp")) = QO_TYPE_CUSTOM And ( CLng(oCurrQO.getAttribute("disp")) = QO_TYPE_CUSTOM_MAPWITHSUBINFO OR CLng(oCurrQO.getAttribute("disp")) = QO_TYPE_CUSTOM_NOMAPPING ) Then
			aPromptGeneralInfo(PROMPT_B_ALLOW_PROFILE) = False
		End If

        aPromptGeneralInfo(PROMPT_B_CHANGE_STYLE) = False
		If CLng(oCurrQO.getAttribute("qtp")) = QO_TYPE_CUSTOM And CLng(oCurrQO.getAttribute("disp")) = QO_TYPE_CUSTOM_NOMAPPING Then
			aPromptGeneralInfo(PROMPT_B_CHANGE_STYLE) = True
		End If

		aPromptGeneralInfo(PROMPT_S_INFORMATIONSOURCE_ID) = oCurrQO.getAttribute("isid")
        Set oISOI = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).selectSingleNode("/mi/in/oi[@tp='" & TYPE_INFORMATION_SOURCE & "' $and$ @id='" & aPromptGeneralInfo(PROMPT_S_INFORMATIONSOURCE_ID) & "']")
		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, oISOI.selectSingleNode("prs/pr[@n='connInfo']").text, oInfoSource)
		Set oInfoSource = oInfoSource.selectSingleNode("/info_source_props")

		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, oCurrQO.selectSingleNode("./prs/pr[@n='definition']").text, oQODefinition)

		Set oISMProgID = oISOI.selectSingleNode("prs/pr[@n='ISM_admin_progid']")
		Select Case oISMProgID.getAttribute("v")
		Case "CastorISM.cCastorISM"
			aPromptGeneralInfo(PROMPT_L_ISM_TYPE) = ISM_TYPE_CASTOR
		Case "UserDetailsISM.cUserDetails"
			aPromptGeneralInfo(PROMPT_L_ISM_TYPE) = ISM_TYPE_USERDETAIL
		Case Else
			aPromptGeneralInfo(PROMPT_L_ISM_TYPE) = ISM_TYPE_CUSTOM
		End Select

        aPromptGeneralInfo(PROMPT_S_SECURITY_FILTERID) = ""

        lErrNumber = cu_getUserSecurityObjects(sUserSecurity)
		If lErrNumber <> NO_ERR Then
			sErrDescription = "Error getting user security object"
			lErrNumber = ERR_GET_HYDRA_PROMPT
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErrDescription), Err.Source, "PromptCuLib.asp", "GetHydraPrompt", "", "Error reading cache file", LogLevelTrace)
		End If

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sUserSecurity, oUserSecurity)
        Set oUserSecurityPR = oUserSecurity.selectSingleNode("/mi/in/oi[@id='" & aPromptGeneralInfo(PROMPT_S_INFORMATIONSOURCE_ID) &"']")

        bUserLevel = True
        If oUserSecurityPR Is Nothing Then
            bUserLevel = False
        ElseIf Len(oUserSecurityPR.getAttribute("v")) = 0 Then
            bUserLevel = False
        End If

        If Not bUserLevel Then
            Set oSecurity = oInfoSource.selectSingleNode("default/security")
            If Not oSecurity Is Nothing Then
                aPromptGeneralInfo(PROMPT_S_SECURITY_PROMPTID) = oInfoSource.selectSingleNode("project").getAttribute("security_prompt")
                aPromptGeneralInfo(PROMPT_S_SECURITY_FILTERID) = oSecurity.getAttribute("id")
            End If
        Else
            sUserSecurity = oUserSecurityPR.getAttribute("v")
            iBeginPos = InStr(1, sUserSecurity, "SecurityObject=""", vbTextCompare) + Len("SecurityObject=""")
            iEndPos = InStr(iBeginPos, sUserSecurity, """", vbTextCompare)

            aPromptGeneralInfo(PROMPT_S_SECURITY_FILTERID) = Mid(sUserSecurity, iBeginPos, iEndPos - iBeginPos)
            aPromptGeneralInfo(PROMPT_S_SECURITY_PROMPTID) = oInfoSource.selectSingleNode("project").getAttribute("security_prompt")
        End If

        aPromptGeneralInfo(PROMPT_S_PROFILE_ORIGINAL_NAME) = ""
        aPromptGeneralInfo(PROMPT_B_USER_DEFAULT) = False

        Set oAnswer = oCurrQO.selectSingleNode("answer")
        If Not oAnswer Is Nothing Then
            'PromptGeneralInfo(PROMPT_S_PROFILE_ORIGINAL_NAME) = oAnswer.getAttribute("originaln")
            aPromptGeneralInfo(PROMPT_S_PROFILE_NAME) = oAnswer.getAttribute("n")
            aPromptGeneralInfo(PROMPT_S_PROFILE_DESC) = oAnswer.getAttribute("desc")
            aPromptGeneralInfo(PROMPT_S_PREF_ID) = oAnswer.getAttribute("prefID")
            Set oDefProfile = oCurrQO.selectSingleNode("mi/oi[@tp='" & TYPE_PROFILE & "' $and$ @def='1']")
            If Not oDefProfile Is Nothing Then
                If aPromptGeneralInfo(PROMPT_S_PREF_ID) = oDefProfile.getAttribute("id") Then
                    aPromptGeneralInfo(PROMPT_B_USER_DEFAULT) = True
                End If
            End If
        End If

        bFreshPrompt = True
        If aPromptGeneralInfo(PROMPT_B_NEEDPROCESS) Then
			bFreshPrompt = False
		End If
    Else
        lErrNumber = ERR_GET_HYDRA_PROMPT
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErrDescription), Err.Source, "PromptCuLib.asp", "GetHydraPrompt", "", "Error processing cache file", LogLevelTrace)
    End If

	If Not bFreshPrompt Then
		aConnectionInfo(S_SERVER_NAME_CONNECTION) = CStr(oRequest("Server"))
		aConnectionInfo(S_PROJECT_CONNECTION) = CStr(oRequest("Project"))
		aConnectionInfo(S_UID_CONNECTION) = CStr(oRequest("Uid"))
		aConnectionInfo(S_TOKEN_CONNECTION) = CStr(oRequest("sToken"))
		aConnectionInfo(N_PORT_CONNECTION) = CLng(oRequest("Port"))

		aPromptGeneralInfo(PROMPT_S_MSGID) = CStr(oRequest("MsgID"))
		aPromptGeneralInfo(PROMPT_S_DOCUMENTID) = CStr(oRequest("DocumentID"))
		aPromptGeneralInfo(PROMPT_S_REPORTID) = CStr(oRequest("ReportID"))
		aPromptGeneralInfo(PROMPT_S_INFORMATIONSOURCE_ID) = CStr(oRequest("InformationSourceID"))
		'aPromptGeneralInfo(PROMPT_S_PROFILE_NAME) = Server.HTMLEncode(CStr(oRequest("ProfileName")))
		'aPromptGeneralInfo(PROMPT_S_PROFILE_DESC) = Server.HTMLEncode(CStr(oRequest("ProfileDesc")))
		aPromptGeneralInfo(PROMPT_S_PROFILE_NAME) = CStr(oRequest("ProfileName"))
		aPromptGeneralInfo(PROMPT_S_PROFILE_DESC) = CStr(oRequest("ProfileDesc"))

		Call GetDSSSession(aConnectionInfo, oSession, sErrDescription)
		If aPromptGeneralInfo(PROMPT_L_ISM_TYPE) = ISM_TYPE_CUSTOM Then
			lErrNumber = CreatePromptsHelperObject(aConnectionInfo, aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT) , sErrDescription)
			aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Locale = CLng(GetLng())
			aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Init(oQODefinition.selectSingleNode("./Question_Object_Definition/*").xml)
		End If
    Else
    	If aPromptGeneralInfo(PROMPT_L_ISM_TYPE) = ISM_TYPE_CUSTOM Then
			lErrNumber = CreatePromptsHelperObject(aConnectionInfo, aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT) , sErrDescription)
			aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Locale = CLng(GetLng())
			aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Init(oQODefinition.selectSingleNode("./Question_Object_Definition/*").xml)
		Else
			Set oReport = Server.CreateObject("WebAPIHelper.DSSXMLResultSet")
			If Not aPromptGeneralInfo(PROMPT_B_MESSAGESAVED) then
				aConnectionInfo(S_SERVER_NAME_CONNECTION) = oInfoSource.selectSingleNode("server/primary").getAttribute("name")
				aConnectionInfo(N_PORT_CONNECTION) = CLng(oInfoSource.selectSingleNode("server/primary").getAttribute("port"))
				Call GetDSSSession(aConnectionInfo, oSession, sErrDescription)

				sProjectID = oInfoSource.selectSingleNode("project").getAttribute("id")
				lErrNumber = MapProjectIDToName(oSession, sProjectID, aConnectionInfo(S_PROJECT_CONNECTION))
				If lErrNumber <> NO_ERR Then
					lErrNumber = ERR_GET_HYDRA_PROMPT
					Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErrDescription), Err.Source, "PromptCuLib.asp", "GetHydraPrompt", "", "Error calling MapProjectIDToName", LogLevelTrace)
				Else
					lErrNumber = cu_GetUserAuthenticationObjects(sUserAuth)
					If lErrNumber <> NO_ERR Then
						sErrDescription = "Error getting user authentication object"
						lErrNumber = ERR_GET_HYDRA_PROMPT
						Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErrDescription), Err.Source, "PromptCuLib.asp", "GetHydraPrompt", "", "Error reading cache file", LogLevelTrace)
					End If
					lErrNumber = LoadXMLDOMFromString(aConnectionInfo,sUserAuth, oUserAuth)

					Set oUserAuthPR = oUserAuth.selectSingleNode("/mi/in/oi[@id='" & aPromptGeneralInfo(PROMPT_S_INFORMATIONSOURCE_ID) & "']")

					bUserLevel = True
					If oUserAuthPR Is Nothing Then
						bUserLevel = False
					ElseIf Len(oUserAuthPR.getAttribute("v")) = 0 Then
						bUserLevel = False
					End If

					If Not bUserLevel Then
						aConnectionInfo(S_UID_CONNECTION) = oInfoSource.selectSingleNode("default/authentication").getAttribute("name")
						aConnectionInfo(S_PWD_CONNECTION) = Decrypt(oInfoSource.selectSingleNode("default/authentication").getAttribute("pwd"))
					Else
						sUserAuth = oUserAuthPR.getAttribute("v")
						Call ParseAuthenticationObject(sUserAuth, aConnectionInfo(S_UID_CONNECTION), aConnectionInfo(S_PWD_CONNECTION), sUserID)
					End If

					'If the user is logged in using IServer NT authentication, pass in the correct auth mode
					If Session("AuthMode") = "2" and Session("castorUserID") = sUserID  then
						aConnectionInfo(S_TOKEN_CONNECTION) = oSession.CreateSession(aConnectionInfo(S_UID_CONNECTION), aConnectionInfo(S_PWD_CONNECTION), , aConnectionInfo(S_PROJECT_CONNECTION), GetLng() , , 2)
					Else
						aConnectionInfo(S_TOKEN_CONNECTION) = oSession.CreateSession(aConnectionInfo(S_UID_CONNECTION), aConnectionInfo(S_PWD_CONNECTION), , aConnectionInfo(S_PROJECT_CONNECTION), GetLng())
					End If
				End If

				If Err.Number <> 0 Then
					If Err.Number = ERR_API_NO_PROJECT_ACCESS Then
						lErrNumber = ERR_API_NO_PROJECT_ACCESS
					Else
						lErrNumber = ERR_GET_HYDRA_PROMPT
					End If
					Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), Err.Source, "PromptCuLib.asp", "GetHydraPrompt", "", "Error creating castor session", LogLevelError)
				Else
					oReport.SessionID = aConnectionInfo(S_TOKEN_CONNECTION)
					oReport.ExecutionFlags = DssXmlExecutionResolve or DssXmlDocExecutionResolve
					If Not oAnswer Is Nothing Then
						oReport.PromptAnswer = oAnswer.selectSingleNode("*").xml
					End If

					If StrComp(oQODefinition.selectSingleNode("/Question_Object_Definition").getAttribute("qra_type") , DssXmlTypeDocumentDefinition) = 0 Then
						'document
						oReport.DocumentOrigin = true
						aPromptGeneralInfo(PROMPT_B_ISDOC) = true
						aPromptGeneralInfo(PROMPT_S_DOCUMENTID) = oQODefinition.selectSingleNode("/Question_Object_Definition").getAttribute("id")
						oReport.Submit (aPromptGeneralInfo(PROMPT_S_DOCUMENTID))
					Else
						'report (blank or DssXmlTypeReportDefinition)
						oReport.DocumentOrigin = false
						aPromptGeneralInfo(PROMPT_B_ISDOC) =false
						aPromptGeneralInfo(PROMPT_S_REPORTID) = oQODefinition.selectSingleNode("/Question_Object_Definition").getAttribute("id")
						oReport.Submit (aPromptGeneralInfo(PROMPT_S_REPORTID))
					End If

					If Err.number = 0 Then
						aPromptGeneralInfo(PROMPT_S_MSGID) = oReport.MessageID
						Call oReport.GetResults
						If oReport.Status = 1 then
							'TQMS: 713924
							'2013.01.31 - vgarcia - we need additional execution flags for this
							oReport.ExecutionFlags = oReport.ExecutionFlags or &h1000
							oReport.Refresh(true)
							Call oReport.GetResults
						End If
						Set aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT) = oReport.PromptsObject
						aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Locale = CLng(GetLng())
						lErrNumber = Err.Number
					Else
						If Err.number = ERR_API_NO_WRITE_ACCESS Then
							lErrNumber = ERR_API_NO_WRITE_ACCESS
							Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), Err.Source, "PromptCuLib.asp", "GetHydraPrompt", "", "Error submiting castor execute request", LogLevelError)
						Else
							lErrNumber = Err.Number
						End If
					End If
				End If
			Else
				aConnectionInfo(S_TOKEN_CONNECTION) = oRequest("sessionID")
				oReport.SessionID = aConnectionInfo(S_TOKEN_CONNECTION)
				oReport.ExecutionFlags = DssXmlExecutionResolve or DssXmlDocExecutionResolve
				oReport.MessageID = aPromptGeneralInfo(PROMPT_S_MSGID)
				Set aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT) = oReport.PromptsObject
				If StrComp(oQODefinition.selectSingleNode("/Question_Object_Definition").getAttribute("qra_type") , DssXmlTypeDocumentDefinition) = 0 Then
					'document
					oReport.DocumentOrigin = true
					aPromptGeneralInfo(PROMPT_B_ISDOC) = true
					aPromptGeneralInfo(PROMPT_S_DOCUMENTID) = oQODefinition.selectSingleNode("/Question_Object_Definition").getAttribute("id")
				Else
					'report (blank or DssXmlTypeReportDefinition)
					oReport.DocumentOrigin = false
					aPromptGeneralInfo(PROMPT_B_ISDOC) =false
					aPromptGeneralInfo(PROMPT_S_REPORTID) = oQODefinition.selectSingleNode("/Question_Object_Definition").getAttribute("id")
				End If
			End If


			If Len(aPromptGeneralInfo(PROMPT_S_SECURITY_FILTERID)) > 0 And lErrNumber = NO_ERR Then
		        Set oPrompts = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT)
		        For lPin = 1 To oPrompts.Count
				    Set oSinglePrompt = oPrompts.Item(lPin)
					If oSinglePrompt.Info.ID = aPromptGeneralInfo(PROMPT_S_SECURITY_PROMPTID) Then
						Set oElements = oSinglePrompt.ElementsObject
						Call oElements.Clear

						'07/11/05 epolo, TQMS 160565: Allow more than one element in the security object prompt.
						'The new security object definition consists of concatenated elements separated by the '#' character.
						sElementIDs = Split(aPromptGeneralInfo(PROMPT_S_SECURITY_FILTERID),"#")

						For i = LBound(sElementIDs) to UBound(sElementIDs)
							Call oElements.Add(sElementIDs(i))
						Next
						Call oSinglePrompt.Answer
					End If
				Next

		        Call oReport.GetResults
		        Set oPrompts = oReport.PromptsObject

		        For lPin = 1 To oPrompts.Count
					Set oSinglePrompt = oPrompts.Item(lPin)
					If oSinglePrompt.Info.ID <> aPromptGeneralInfo(PROMPT_S_SECURITY_PROMPTID) Then
						If oSinglePrompt.PromptType = DssPromptElements Then
							Call FilterDefaultPromptAnswer(oSinglePrompt, aPromptGeneralInfo(PROMPT_S_SECURITY_FILTERID), aConnectionInfo(S_TOKEN_CONNECTION))
						End If
					End If
				Next
		        Set aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT) = oPrompts
		    End If
        End If

		If lErrNumber <> NO_ERR Then
			If lErrNumber <> ERR_API_NO_PROJECT_ACCESS AND lErrNumber <> ERR_API_NO_WRITE_ACCESS Then
				lErrNumber = ERR_GET_HYDRA_PROMPT
			End IF
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErrDescription), Err.Source, "PromptCuLib.asp", "GetHydraPrompt", "", "Error setting security or previous answer", LogLevelTrace)
		End If

    End If

    Set oCurrQO = Nothing
    Set oReport = Nothing
    Set oInfoSource = Nothing
    Set oAnswerContent = Nothing
    Set oSecurity = Nothing
    Set oPrompts = Nothing
    Set oElements = Nothing
    Set oSub = Nothing
    Set oUserAuthPR = Nothing
    Set oISMProgID = Nothing

    GetHydraPrompt = lErrNumber
    Err.Clear
End Function


Function FilterDefaultPromptAnswer(oPrompt, sElementID, sSessionID)
	Dim sAttributeID
	Dim oExpression
	Dim oOperatorNode
	Dim oShortCutNode
	Dim oElementListNode
	Dim oElementSource
	Dim sElements
	Dim oAnswerElement
	Dim i
	Dim sElementIDs

	If isObject(oPrompt.ElementsObject) Then

		sAttributeID = left(sElementID, 32)

		'07/11/05 epolo, TQMS 160565: Allow more than one element in the security object prompt.
		'The new security object definition consists of concatenated elements separated by the '#' character.
		sElementIDs = Split(sElementID,"#")

		Set oExpression = Server.CreateObject("WebAPIHelper.DSSXMLExpression")
		Set oOperatorNode = oExpression.CreateOperatorNode(DssXmlFilterListQual, DssXmlFunctionIn)

		Set oShortCutNode = oExpression.CreateShortcutNode(sAttributeID, DssXmlTypeAttribute, oOperatorNode)
		Set oElementListNode = oExpression.CreateElementListNode(sAttributeId, oOperatorNode)

		For i = LBound(sElementIDs) To UBound(sElementIDs)
			oElementListNode.ElementsObject.Add sElementIDs(i)
		Next

		Set oElementSource = Server.CreateObject("WebAPIHelper.DSSXMLElementSource")
		oElementSource.sessionID = sSessionID
		set oElementSource.ExpressionObject = oExpression
		oElementSource.AttributeID = oPrompt.ElementSourceObject.AttributeID

		sElements = oElementSource.getElements()

		For i = 1 to oPrompt.ElementsObject.Count
			Set oAnswerElement = oPrompt.ElementsObject(i)
			If instr(1, sElements, """" & oAnswerElement.ElementID & """") < 1 Then
				oPrompt.ElementsObject.Remove i
			End If
		Next

	End If

End Function


Function ReadPromptQuestionFromCache(aConnectionInfo, aPromptGeneralInfo)
'******************************************************************************
'Purpose: Read Prompt Quesiton XML To Cache For Hydra
'Inputs:  aConnectionInfo, aPromptGeneralInfo
'Outputs:
'******************************************************************************
    On Error Resume Next
    Dim sCacheXML

    lErrNumber = ReadCache(aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID), CStr(GetSessionID()), sCacheXML)
    If lErrNumber = NO_ERR Then
        'lErrNumber = LoadXMLDOMFromString(aConnectionInfo, Replace(Replace(Replace(sCacheXML, "&gt;", ">"), "&lt;", "<"), "&amp;", "&"), aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS))
        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS))
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptCuLib.asp", "ReadPromptQuestionFromCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
        End If
    Else
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptCuLib.asp", "ReadPromptQuestionFromCache", "", "Error in call to ReadCache", LogLevelTrace)
    End If

    ReadPromptQuestionFromCache = lErrNumber
    Err.Clear
End Function

Function SavePromptAnswerToCache(aConnectionInfo, aPromptGeneralInfo)
'******************************************************************************
'Purpose: Read Prompt Quesiton XML To Cache For Hydra
'Inputs:  aConnectionInfo, aPromptGeneralInfo
'Outputs:
'******************************************************************************
    On Error Resume Next

    Call WriteCache(aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID), CStr(GetSessionID()), aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).xml)

    SavePromptAnswerToCache = Err.Number
    Err.Clear
End Function

Function AnswerPromptByWidget(aConnectionInfo, oRequest, aPromptGeneralInfo)
'******************************************************************************
'Purpose: Save Prompt AnswerXML To Cache for Hydra
'Inputs:  aConnectionInfo, aPromptGeneralInfo
'Outputs:
'******************************************************************************
    On Error Resume Next
    Dim sPreferenceID
    Dim oAnswerXML
    Dim oAnswer
    Dim oTemp
    Dim oTempXML
    Dim oFile
    Dim oCurrQO
    Dim i
    Dim oSinglePrompt

    sPreferenceID = ""

    Set oCurrQO = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).selectSingleNode("/mi/qos/mi/in/oi[@tp='" & TYPE_QUESTION & "' $and$ @id='" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) & "']")

    '<answer> node
    Set oAnswer = oCurrQO.selectSingleNode("answer")
    If Not oAnswer Is Nothing Then
        sPreferenceID = oAnswer.getAttribute("prefID")
        Call oCurrQO.RemoveChild(oAnswer)
    End If
    Set oAnswer = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).createElement("answer")
    Call oCurrQO.appendChild(oAnswer)
    Call oAnswer.setAttribute("n", "")
    Call oAnswer.setAttribute("desc", "")
    Call oAnswer.setAttribute("prefID", "")
    Call oAnswer.setAttribute("def", "")

    'prefID and name
	If Len(sPreferenceID) > 0 And Len(oAnswer.getAttribute("n")) = 0 Then	'edit from preference to preference
		Call oAnswer.setAttribute("prefID", sPreferenceID)
	Else	'new, or, edit from profile to preference
    	Call oAnswer.setAttribute("prefID", "")
	End If
    Call oAnswer.setAttribute("n", aPromptGeneralInfo(PROMPT_S_PROFILE_NAME))

    'desc and def
    Call oAnswer.setAttribute("desc", aPromptGeneralInfo(PROMPT_S_PROFILE_DESC))
    'If Len(CStr(oRequest("userDef"))) > 0 Then
    '    Call oAnswer.setAttribute("def", "1")
    'Else
    '    Call oAnswer.setAttribute("def", "0")
    'End If
    Call oAnswer.setAttribute("def", "1")		'set it as default always

    'open security filter to include it in answerXML
    If Len(aPromptGeneralInfo(PROMPT_S_SECURITY_PROMPTID)) > 0 Then
        For i = 1 To aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Count
            Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(i)
            If oSinglePrompt.Info.ID = aPromptGeneralInfo(PROMPT_S_SECURITY_PROMPTID) Then
                Call oSinglePrompt.ReopenPrompt
                Exit For
            End If
        Next
    End If

    Call aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).ReOpenPrompts()
    aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).AnswerFormat = DssXmlAnswerFormatFlat
    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).AnswerXML, oAnswerXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, lErrNumber, Err.Description, Err.Source, "PromptCuLib.asp", "AnswerPromptByWidget", "LoadXMLDOMFromString()", "", LogLevelTrace)
    End If

    If lErrNumber = NO_ERR Then
        Call oAnswer.appendChild(oAnswerXML.selectSingleNode("*"))

        'save temp XML
        Set oTemp = oCurrQO.selectSingleNode("temp")
        If Not oTemp Is Nothing Then
            Call oCurrQO.RemoveChild(oTemp)
        End If

        Set oTemp = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).createElement("temp")
        Call oCurrQO.appendChild(oTemp)

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML).xml, oTempXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, Err.Description, Err.Source, "PromptCuLib.asp", "AnswerPromptByWidget", "LoadXMLDOMFromString()", "", LogLevelTrace)
        Else
            Call oTemp.appendChild(oTempXML.selectSingleNode("*"))
            Call SavePromptAnswerToCache(aConnectionInfo, aPromptGeneralInfo)
        End If
    End If

    Set oAnswerXML = Nothing
    Set oAnswer = Nothing
    Set oTemp = Nothing
    Set oTempXML = Nothing
    Set oFile = Nothing
    Set oCurrQO = Nothing
    Set oSinglePrompt = Nothing

    AnswerPromptByWidget = Err.Number
    Err.Clear
End Function

Function GetCustomQuestionObject(sQuestionXML)
'******************************************************************************
'Purpose: Get Custom QuestionXML from File
'Inputs:
'Outputs: sQuestionXML
'******************************************************************************
    On Error Resume Next
    Dim oFS
    Dim oFile

    Set oFS = Server.CreateObject("Scripting.FileSystemObject")
    If IsObject(oFS) Then
        Set oFile = oFS.OpenTextFile(APP_CACHE_FOLDER & GetSessionID() & "\" & "sample.xml")
        If IsObject(oFile) Then
            sQuestionXML = oFile.ReadAll()
        End If
    End If

    Set oFS = Nothing
    Set oFile = Nothing

    GetCustomQuestionObject = Err.Number
    Err.Clear
End Function

Function BuildSecurityFilter(aConnectionInfo, oSession, aPromptGeneralInfo, oSinglePrompt, oFilterExp)
'******************************************************************************
'Purpose: Build Expression for Security Filter
'Inputs:
'Outputs: oFilterExp
'******************************************************************************
    On Error Resume Next
    Dim oOperatorNode
    Dim oElementList
    Dim oElements
    Dim sEI
    Dim sATDID
    Dim iPos
    Dim sElementIDs
    Dim i

    sEI = aPromptGeneralInfo(PROMPT_S_SECURITY_FILTERID)
    iPos = InStr(1, sEI, ":", vbBinaryCompare)
    sATDID = Left(sEI, iPos - 1)

    sElementIDs = Split(sEI,"#")

	If oFilterExp Is Nothing Or IsNull(oFilterExp) Or IsEmpty(oFilterExp) Then
		Set oFilterExp = Server.CreateObject("WebAPIHelper.DSSXMLExpression")
	End If

	oFilterExp.RootNode.Operator = DssXmlFunctionAnd

    Set oOperatorNode = oFilterExp.CreateOperatorNode(DssXmlFilterListQual, DssXmlFunctionIn)
    Call oFilterExp.CreateShortCutNode(sATDID, DssXmlTypeAttribute, oOperatorNode)
    Set oElementList = oFilterExp.CreateElementListNode(sATDID, oOperatorNode)
    Set oElements = oElementList.ElementsObject

    For i = LBound(sElementIDs) To UBound(sElementIDs)
    	Call oElements.Add(sElementIDs(i))
    Next

    Set oOperatorNode = Nothing
    Set oElementList = Nothing
    Set oElements = Nothing

    BuildSecurityFilter = Err.Number
    Err.Clear
End Function

'not used any more
Function RenderQOList(aConnectionInfo, aPromptGeneralInfo)
'******************************************************************************
'Purpose: Build Expression for Security Filter
'Inputs:
'Outputs: oFilterExp
'******************************************************************************
    On Error Resume Next
    Dim oHydraPrompts

    lErrNumber = ReadPromptQuestionFromCache(aConnectionInfo, aPromptGeneralInfo)

    Call RenderQuestions_Personalize(aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).selectSingleNode("/mi/qos/mi").xml, aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID), CStr(oRequest("folderID")))

    RenderQOList = Err.Number
    Err.Clear
End Function


'not used any more
Function MapProjectIDToName_old(oSession, sProjectID, sProjectName)
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
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErrDescription, Err.Source, "CommonHeaderCuLib.asp", "GetProjectsFromServer", "oServerSession.GetProjects", "Error getting projects from server", LogLevelError)
        Select Case lErrNumber
            Case API_ERR_PROJECT_OFFLINE
            Case API_ERR_LOGIN_PASSWORD_INVALID
            Case API_ERR_SERVER_NOT_FOUND
                sErrDescription = asDescriptors(160)  'Descriptor: The MicroStrategy Server you are trying to connect to was not found. Please try again later.
            Case API_ERR_USER_PRIVILEGES
            'case AUTHEN_E_LOGIN_FAILED_NEW_PASSWORD_REQD
            Case AUTHEN_E_ACCOUNT_DISABLED
            Case AUTHEN_E_LOGIN_FAIL_EXPIRED_PWD
            Case API_ERR_NT_NOT_LINKED

            'case
                    '-2147467259            "FindClass failed. (Success)"
                    'XXX                    "JVM Can't created"
        End Select
    Else
        Set oProjectsXML = Server.CreateObject("Microsoft.XMLDOM")
        Call oProjectsXML.loadXML(sProjectsXML)
        Set oProject = oProjectsXML.selectSingleNode("/mi/srps/sp[@ps = '0' $and$ @pgd= '" & sProjectID & "']")
        If oProject Is Nothing Then
            lErrNumber = ERR_PROJECT_NAME_NOT_EXIST
            sErrDescription = asDescriptors(551) 'Descriptor: The Project ## could not be found in the server ###
        Else
            sProjectName = oProject.getAttribute("pn")
        End If
    End If

    Set oProjectsXML = Nothing
    Set oProject = Nothing

    MapProjectIDToName = lErrNumber
    Err.Clear
End Function

Function AnswerPromptByProfile(aConnectionInfo, aPromptGeneralInfo, oRequest)
'******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'******************************************************************************
    On Error Resume Next
    Dim oCurrQO
    Dim oAnswer
    Dim sTemp
    Dim temArray
    Dim oAnswerContent

    Set oCurrQO = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).selectSingleNode("/mi/qos/mi/in/oi[@tp='" & TYPE_QUESTION & "' $and$ @id='" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) & "']")

    Set oAnswer = oCurrQO.selectSingleNode("answer")
    If oAnswer Is Nothing Then
        Set oAnswer = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).createElement("answer")
        Call oCurrQO.appendChild(oAnswer)
    Else
        Set oAnswerContent = oAnswer.selectSingleNode("*")
        If Not oAnswerContent Is Nothing Then
            Call oAnswer.RemoveChild(oAnswerContent)
        End If
    End If

    Call oAnswer.setAttribute("def", "")
    Call oAnswer.setAttribute("desc", "")

    If oRequest("ProfileList").Count > 0 Then
        sTemp = CStr(oRequest("ProfileList"))
        If strcomp(sTemp, "-none-", vbTextCompare) = 0 Then
			lErrNumber = ERR_NEED_PROFILE_ANSWER
			sErrDescription = asDescriptors(777)	'Please pick a profile for this question object
		Else
			temArray = Split(sTemp, ":", -1, vbBinaryCompare)
			Call oAnswer.setAttribute("prefID", temArray(0))
			Call oAnswer.setAttribute("n", temArray(1))
			'all oAnswer.setAttribute("originaln", "")
		End If
	Else
		lErrNumber = ERR_NEED_PROFILE_ANSWER
		sErrDescription = asDescriptors(777)	'Please pick a profile for this question object
    End If

    Call SavePromptAnswerToCache(aConnectionInfo, aPromptGeneralInfo)

    Set oCurrQO = Nothing
    Set oAnswer = Nothing
    Set oAnswerContent = Nothing

    AnswerPromptByProfile = lErrNumber
    Err.Clear
End Function


Function DisplayProfileList(aConnectionInfo, aPromptGeneralInfo)
'******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'******************************************************************************
    On Error Resume Next
    Dim oCurrQO
    Dim oProfiles
    Dim oCurrentProfile
    Dim sDisplay

    Set oCurrQO = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).selectSingleNode("/mi/qos/mi/in/oi[@tp='" & TYPE_QUESTION & "' $and$ @id='" & aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID) & "']")

    Response.Write "<FONT FACE=""Verdana,Arial,Helvetica"" SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """>"
    Response.Write "<B>" & Server.HTMLEncode(oCurrQO.getAttribute("n")) & "<BR /></B>"
    Response.Write "<FONT SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & oCurrQO.getAttribute("des") & "<BR /><BR /></FONT>"

    Set oProfiles = oCurrQO.selectNodes("mi/oi[@tp='" & TYPE_PROFILE & "']")

    If oProfiles.length > 0 Then
        Response.Write "<TABLE BORDER=""0"" COLS=""2"">"
        Response.Write "    <TR>"
        Response.Write "        <TD WIDTH=""13""><IMG SRC=""images/1arrow_right.gif"" WIDTH=""13"" HEIGHT=""13"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "        <TD>"
        Response.Write "        <FONT FACE=""Verdana,Arial,Helvetica"" SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """><B>" & asDescriptors(550) & "</B></FONT>"   'Descriptor: Apply these selections.
		Response.Write "        </TD>"
        Response.Write "    </TR>"
        Response.Write "</TABLE>"

        Response.Write "<TABLE BORDER=""0"" COLS=""3"" WIDTH=""100%"" CELLPADDING=""0"">"
        Response.Write "    <TR>"
        Response.Write "        <TD WIDTH=""24"" ROWSPAN=""3""><IMG SRC=""images/1ptrans.gif"" WIDTH=""24"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "        <TD COLSPAN=""2"">"
        Response.Write "        <FONT FACE=""Verdana,Arial,Helvetica"" SIZE=""1"">" & asDescriptors(548) & "<BR /></FONT>"   'Descriptor: Choose from this list of previously saved selections

        Response.Write "        <NOBR><select name=""ProfileList"" class=""pullDownClass"">"

        Response.Write "        <option value=""" & "-none-" & """> ** " & asDescriptors(549) & " ** </option>"   'Descriptor: Choose one

        For Each oCurrentProfile In oProfiles
            sDisplay = oCurrentProfile.getAttribute("n")
            'If oCurrentProfile.getAttribute("def") = "1" Then
            '    sDisplay = sDisplay & " (d)"
            'End If
			sDisplay = Server.HTMLEncode(sDisplay)

            If aPromptGeneralInfo(PROMPT_S_PREF_ID) = oCurrentProfile.getAttribute("id") Then
                Response.Write "<option selected=""1"" value=""" & oCurrentProfile.getAttribute("id") & ":" & Server.HTMLEncode(oCurrentProfile.getAttribute("n")) & """>" & sDisplay & "</option>"
            Else
                If Len(aPromptGeneralInfo(PROMPT_S_PREF_ID)) > 0 Then
                    Response.Write "<option value=""" & oCurrentProfile.getAttribute("id") & ":" & Server.HTMLEncode(oCurrentProfile.getAttribute("n")) & """>" & sDisplay & "</option>"
                Else
                    If StrComp(oCurrentProfile.getAttribute("def"), "1", vbBinaryCompare) = 0 Then
                        Response.Write "<option selected=""1"" value=""" & oCurrentProfile.getAttribute("id") & ":" & Server.HTMLEncode(oCurrentProfile.getAttribute("n")) & """>" & sDisplay & "</option>"
                    Else
                        Response.Write "<option value=""" & oCurrentProfile.getAttribute("id") & ":" & Server.HTMLEncode(oCurrentProfile.getAttribute("n")) & """>" & sDisplay & "</option>"
                    End If
                End If
            End If

        Next

        Response.Write "        </select></NOBR>"
        Response.Write "        <IMG WIDTH=""6"" HEIGHT=""1"" ALT="""" BORDER=""0"" SRC=""images/1ptrans.gif"" />"
        Response.Write "        <INPUT TYPE=""SUBMIT"" CLASS=""buttonClass"" NAME=""ProfileEdit"" VALUE=""" & asDescriptors(353) & """ />"        'Descriptor: Edit
        Response.Write "        <INPUT TYPE=""SUBMIT"" CLASS=""buttonClass"" NAME=""ProfileDelete"" VALUE=""" & asDescriptors(249) & """ />"        'Descriptor: Delete

        Response.Write "        </TD>"
        Response.Write "    </TR>"
        Response.Write "    <TR>"
        Response.Write "        <TD COLSPAN=""2"" BGCOLOR=""#cccccc""><IMG SRC=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "    </TR>"
        Response.Write "    <TR>"
        Response.Write "        <TD ALIGN=""LEFT"" NOWRAP=""1"">"
        Response.Write "        <INPUT TYPE=""SUBMIT"" CLASS=""buttonClass"" NAME=""HydraBack1"" VALUE=""" & asDescriptors(149) & """/>"    'Descriptor: Back
        Response.Write "        <INPUT TYPE=""SUBMIT"" CLASS=""buttonClass"" NAME=""HydraNext1"" VALUE=""" & asDescriptors(335) & """/>" 'Descriptor: Next
        Response.Write "        </TD>"
        Response.Write "        <TD ALIGN=""RIGHT"" NOWRAP=""1"">"
        If aPromptGeneralInfo(PROMPT_B_FINISH_ENABLED) Then
            Response.Write "            <INPUT TYPE=""SUBMIT"" CLASS=""buttonClass"" NAME=""HydraFinish1"" VALUE=""" & asDescriptors(442) & """/>"      'Descriptor: Finish
        'Else
        '    Response.Write "            <INPUT TYPE=""Button"" CLASS=""disabledButton"" NAME=""HydraFinish1"" VALUE=""" & asDescriptors(442) & """/>"   'Descriptor: Finish
        End If
        Response.Write "            <INPUT TYPE=""SUBMIT"" CLASS=""buttonClass"" NAME=""cancel"" VALUE=""" & asDescriptors(120) & """/>"  'Descriptor: Cancel
        Response.Write "        </TD>"
        Response.Write "    </TR>"
        Response.Write "</TABLE>"

        Response.Write "<TABLE BORDER=""0"" COLS=""2"" WIDTH=""100%"">"
        Response.Write "    <TR>"
        Response.Write "        <TD WIDTH=""13""><IMG SRC=""images/1ptrans.gif"" WIDTH=""13"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "        <TD>"
        Response.Write "        <FONT FACE=""Verdana,Arial,Helvetica"" SIZE=""2""><B><BR /><BR />" & asDescriptors(339) & "<BR /><BR /></B></FONT>"   'Descriptor: OR
        Response.Write "        </TD>"
        Response.Write "    </TR>"
        Response.Write "    <TR>"
        Response.Write "        <TD WIDTH=""13""><IMG SRC=""images/arrow_down.gif"" WIDTH=""13"" HEIGHT=""13"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "        <TD>"
        Response.Write "        <FONT FACE=""Verdana,Arial,Helvetica"" SIZE=""2""><B>" & asDescriptors(547) & "</B></FONT>"       'Descriptor: Make new selections
        Response.Write "        </TD>"
        Response.Write "    </TR>"
        Response.Write "</TABLE>"
    End If

    DisplayProfileForm = lErrNumber
    Err.Clear
End Function

Function DisplayProfileNameDesc(aConnectionInfo, aPromptGeneralInfo)
'******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'******************************************************************************
    On Error Resume Next

    Response.Write "<!-- BEGIN: Hydra Profile Name box -->"
    Response.Write "<TABLE BORDER=""0"" COLS=""3"" WIDTH=""100%"" CELLPADDING=""0"">"
    Response.Write "    <TR>"
    Response.Write "    <TD WIDTH=""1%""><IMG SRC=""images/1ptrans.gif"" WIDTH=""20"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
    Response.Write "    <TD ALIGN=""LEFT"">"	'from LEFT to RIGHT
    'Response.Write "        <FONT FACE=""Verdana,Arial,Helvetica"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(546) & "</B><BR />Name:<BR /></FONT>"   'Descriptor: Save Selections As (optional)
    Response.Write "        <FONT FACE=""Verdana,Arial,Helvetica"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(546) & ":<BR /></FONT>"   'Descriptor: Save Selections As (optional)
    Response.Write "        <INPUT TYPE=""TEXT"" SIZE=""22"" MAXLENGTH=""25"" NAME=""ProfileName"" VALUE=""" & Server.HTMLEncode(aPromptGeneralInfo(PROMPT_S_PROFILE_NAME)) & """ />"
    Response.Write "    </TD>"
    'Response.Write "    <TD ALIGN=""LEFT"">"
    'Response.Write "        <FONT FACE=""Verdana,Arial,Helvetica"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"
    'If aPromptGeneralInfo(PROMPT_B_USER_DEFAULT) Then
    '    Response.Write "        <INPUT TYPE=""CHECKBOX"" NAME=""userDef"" VALUE=""1"" CHECKED=""1"">" & asDescriptors(695) & "</INPUT>" 'Descriptor: Use these selections by default
    'Else
    '    Response.Write "        <INPUT TYPE=""CHECKBOX"" NAME=""userDef"" VALUE=""1"">" & asDescriptors(695) & "</INPUT>" 'Descriptor: Use these selections by default
    'End If
    'Response.Write "        <BR />" & asDescriptors(22) & " :<BR /></FONT>"            'Descriptor: Description
    'Response.Write "        <INPUT TYPE=""TEXT"" SIZE=""15"" MAXLENGTH=""100"" NAME=""ProfileDesc"" VALUE=""" & aPromptGeneralInfo(PROMPT_S_PROFILE_DESC) & """ />"
    'Response.Write "    </TD>"
    Response.Write "    </TR>"
    Response.Write "</TABLE>"
    Response.Write "<!-- END: Hydra Profile Name box -->"

    DisplayProfileNameDesc = Err.Number
    Err.Clear
End Function

Function CreateResultSetHelperObject(aConnectionInfo, oObjResultSet, sErrDescription)
'*******************************************************************************
'Purpose:   Gets the Result Set helper object
'Inputs:	aConnectionInfo
'Outputs:	oObjResultSet, lErrNumber, sErrDescription
'*******************************************************************************
    On Error Resume Next
    Dim lErrNumber
    lErrNumber = 0
    Set oObjResultSet = Server.CreateObject("WebAPIHelper.DSSXMLResultSet")
    lErrNumber = Err.number
    If lErrNumber <> NO_ERR Then
		sErrDescription = Err.description
		Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "CommonLib.asp", "CreateResultSetHelperObject", "", "Error after calling Server.CreateObject(""WebAPIHelper.DSSXMLResultSet"")", LogLevelError)
    Else
		oObjResultSet.SessionID = aConnectionInfo(S_TOKEN_CONNECTION)
		oObjResultSet.DocumentOrigin = False
    End If
    CreateResultSetHelperObject = lErrNumber
    Err.Clear
End Function


%>
