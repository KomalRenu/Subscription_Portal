<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Private Const lBlockSize = 50000

Function CreateDisplayXMLForAllPrompts(aConnectionInfo, oSession, aPromptInfo, aPromptGeneralInfo, oRequest, sErrDescription)
'*************************************************************************************************************
'Purpose:   display all the prompts for the job.
'Inputs:    aConnectionInfo, oSession, aPromptInfo, aPromptGeneralInfo, oRequest
'Outputs:   sErrDescription
'*************************************************************************************************************
    On Error Resume Next
    Dim oSinglePromptQuestionXML
    Dim oSinglePromptTempXML
    Dim sPin
    Dim lPin
    Dim bAnyPromptSucceeded
    Dim lErrNumber
    Dim lPType
    Dim oInputs
    Dim oDisplayAllPromptsXML
    Dim oDisplayXML
    Dim oNewDisplayAllXML
    Dim bFirst
    Dim lOrder
    Dim oSinglePrompt

    bFirst = True
    If (aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE)) Then
		lOrder = 1
	Else
		lOrder = CLng(aPromptGeneralInfo(PROMPT_S_CURORDER))
	End If

    bAnyPromptSucceeded = False
    Call GetXMLDOM(aConnectionInfo, oDisplayAllPromptsXML, sErrDescription)
    Set oNewDisplayAllXML = oDisplayAllPromptsXML.createElement("mi")
    Call oDisplayAllPromptsXML.appendChild(oNewDisplayAllXML)
	set aPromptGeneralInfo(PROMPT_O_DISPLAYXML) = oNewDisplayAllXML

    If lErrNumber = 0 Then
		While ((lOrder <= aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT)) And bFirst)
			Call GetPinbyOrder(aPromptGeneralInfo, aPromptInfo, lOrder, lPin)
			Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)
			If oSinglePrompt.Used And Not(oSinglePrompt.Closed) Then
				If Not(aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE)) Then
					bFirst = False
				End If
				lPType = oSinglePrompt.PromptType
				set oSinglePromptQuestionXML = aPromptInfo(lPin, PROMPTINFO_O_QUESTION)
				sPin = CStr(lPin)
				'Set oSinglePromptTempXML = aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML).selectSingleNode("./mi/pif[@pin='" & sPin & "']")
				Set oSinglePromptTempXML = aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER)

                lErrNumber = CreateDisplayXMLForSinglePrompt(aConnectionInfo, oSinglePrompt, lOrder, oSession, aPromptInfo, sPin, lPType, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oInputs, oDisplayXML, sErrDescription)
                If lErrNumber <> 0 Then
					'Handling error for issue 129585
					If lErrNumber = -2147207171 Then
						Call aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).inboxObject.RemoveMessages()
					End If

					Call oDisplayXML.setAttribute("errnumber", CStr(lErrNumber))
					Call oDisplayXML.setAttribute("errdescription", CStr(sErrDescription))
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), Err.Description, Err.source, "PromptDisplayCuLib.asp", "DisplayAllPromptsForJob", "", "Error in call to " & "CreateDisplayXMLForSinglePrompt", LogLevelTrace)
                    If lErrNumber <> ERR_UNSUPPORTED_PROMPTS Then lErrNumber = ERR_DISPLAY_FAILED
                    'sErrDescription = ""
                Else
					Call oDisplayXML.setAttribute("errnumber", "0")
                    bAnyPromptSucceeded = True
                End If

                'If lErrNumber = 0 Then
				Call oNewDisplayAllXML.appendChild(oDisplayXML)
                'End If
				aPromptGeneralInfo(PROMPT_S_ANSWERSXML) = oSinglePrompt.ShortAnswerXML
            End If
            lOrder = lOrder + 1
        Wend
    Else
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "DisplayAllPromptsForJob", "", "Error in call to BuildInputNode", LogLevelTrace)
    End If

	If aPromptGeneralInfo(PROMPT_B_XML) Then
		Response.Write "<!-- oDisplayXML.xml: " & aPromptGeneralInfo(PROMPT_O_DISPLAYXML).xml & "-->"       'test only
	End If


	If lErrNumber = NO_ERR Then
		aPromptGeneralInfo(PROMPT_B_SPECIAL_FORM) = False
		'Hydra - Cannot support Text File in HTML
		If aPromptGeneralInfo(PROMPT_B_DHTML) Then
			If aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) Then
				If aPromptGeneralInfo(PROMPT_B_ANY_TEXTFILE) Then
					aPromptGeneralInfo(PROMPT_B_SPECIAL_FORM) = True
				End if
			Else
				If StrComp(aPromptInfo(lPin, PROMPTINFO_S_XSLFILE), "promptexpression_textfile.xsl", vbTextCompare) = 0 Then
					aPromptGeneralInfo(PROMPT_B_SPECIAL_FORM) = True
				End if
			End if
		End If
	End If

	set oSinglePromptQuestionXML = nothing
	set	oSinglePromptTempXML = nothing
	set oInputs = nothing
    set oDisplayAllPromptsXML = nothing
    set oDisplayXML = nothing
    set oNewDisplayAllXML = nothing

	CreateDisplayXMLForAllPrompts = lErrNumber
	Err.Clear
End Function

Function DisplayAllPrompts(aConnectionInfo, aPromptInfo, aPromptGeneralInfo, sErrDescription)
'*************************************************************************************************************
'Purpose:   display all the prompts for the job.
'Inputs:    aConnectionInfo, aPromptInfo, aPromptGeneralInfo
'Outputs:   sErrDescription
'*************************************************************************************************************
    On Error Resume Next
	Dim lErrNumber
	Dim oDisplays
	Dim oDisplay
	Dim sPin
	Dim oInputs
	Dim lOrder

	lErrNumber = BuildInputNode(aConnectionInfo, aPromptGeneralInfo, oInputs)
    If lErrNumber <> 0 Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "DisplayAllPromptsForJob", "", "Error in call to BuildInputNode", LogLevelTrace)
        lErrNumber = ERR_DISPLAY_FAILED
    End If

	Set oDisplays = aPromptGeneralInfo(PROMPT_O_DISPLAYXML).selectNodes("/mi/pif")

    lOrder = 0
    If oDisplays.length = 0 Then
		Response.Write asDescriptors(826) & asDescriptors(456)
    Else
		For Each oDisplay In oDisplays
			sPin = oDisplay.getAttribute("pin")
			lErrNumber = DisplaySinglePrompt(aConnectionInfo, aPromptInfo, aPromptGeneralInfo, oDisplay, sPin, oInputs, sErrDescription)

			If Not (aPromptGeneralInfo(PROMPT_B_REPROMPT) And StrComp(aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE), "PromptExpression_textbox.xsl") = 0 And _
			   aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).item(CLng(sPin)).ExpressionType <> DssXmlFilterSingleMetricQual) Then
				Call DisplaySinglePromptDefaultMeaning(aPromptGeneralInfo, aPromptInfo, CLng(sPin))
			End If
		Next
	End If

    set oDisplays = nothing
	set oDisplay = nothing
	set oInputs = nothing

	DisplayAllPrompts = Err.number
	Err.Clear
End Function

Function DisplaySinglePrompt(aConnectionInfo, aPromptInfo, aPromptGeneralInfo, oDisplay, sPin, oInputs, sErrDescription)
'*************************************************************************************************************
'Purpose:   display all the prompts for the job.
'Inputs:    aConnectionInfo, oSession, aPromptInfo, aPromptGeneralInfo, oRequest
'Outputs:   sErrDescription
'*************************************************************************************************************
    On Error Resume Next
	Dim lErrNumber
	Dim sXSL
	Dim oSinglePromptXSL
	Dim sURL

	lErrNumber = CLng(oDisplay.getAttribute("errnumber"))
	sErrDescription = oDisplay.getAttribute("errdescription")
	sXSL = aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE)	'"PromptExpression_textfile.xsl"'

	If StrComp(sXSL,"PromptExpression_HierTree.xsl", vbTextCompare) = 0 Then sXSL = "PromptExpression_HierCart_OPTsearch_drill_NOqual.xsl"


	If lErrNumber = 0 Then

		If len(ReadUserOption(ACCESSIBILITY_OPTION))>0 then
			If StrComp(sXSL, "PromptElement_MultiSelect_listbox.xsl", vbTextCompare) = 0 then
				sXSL = "PromptElement_cart.xsl"
			ElseIf StrComp(sXSL, "PromptElement_SingleSelect_listbox.xsl", vbTextCompare) = 0 then
					sXSL = "PromptElement_pulldown.xsl"
				ElseIf StrComp(sXSL, "PromptExpression_SingleSelect_listbox.xsl", vbTextCompare) = 0 then
						sXSL = "PromptExpression_pulldown.xsl"
					ElseIf StrComp(sXSL, "PromptObject_MultiSelect_listbox.xsl", vbTextCompare) = 0 then
							sXSL = "PromptObject_cart.xsl"
						ElseIf StrComp(sXSL, "PromptObject_SingleSelect_listbox.xsl", vbTextCompare) = 0 then
							sXSL = "PromptObject_pulldown.xsl"
			End If
		End If

	    lErrNumber = MapPromptXSL(aConnectionInfo, sXSL, oSinglePromptXSL)

		if lErrNumber<>0 then
	        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForSinglePrompt", "", "Error loading XSL file", LogLevelTrace)
	    End If

		If lErrNumber = 0 Then
		    lErrNumber = AddInputsForPrompts(aConnectionInfo, sPin, aPromptInfo, oInputs, oDisplay)
		    If lErrNumber <> 0 Then
		        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForSinglePrompt", "", "Error in call to AddInputsForPrompts", LogLevelTrace)
		    End If
		End If

 		If lErrNumber = 0 Then
 			If StrComp(sXSL, "PromptExpression_Cart.xsl", vbTextCompare) = 0 And _
 			   aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).item(CLng(sPin)).ExpressionType = DssXmlFilterSingleMetricQual Then
 				oDisplay.ownerDocument.selectSingleNode("./mi/inputs/DHTML").text = ""
			End If

		    Response.Write oDisplay.transformNode(oSinglePromptXSL)
		    lErrNumber = Err.number
		    If lErrNumber <> 0 Then
		        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForSinglePrompt", "", "Error oXMLRoot.transformNode", LogLevelError)
		    End If
		End If
	ElseIf lErrNumber = ERR_API_REQUEST_TIMED_OUT Then
		Call WritePromptError(sPin, aPromptInfo, asDescriptors(109))
		aPromptGeneralInfo(PROMPT_B_ANYERROR) = True
	Else
		Call WritePromptError(sPin, aPromptInfo, sErrDescription)
		aPromptGeneralInfo(PROMPT_B_ANYERROR) = True
		'sURL = GetGeneralParasinURL(oRequest)
		'Response.Redirect "JobError.asp?ErrNum=" & lErrNumber & "&ErrDesc=" & Server.URLEncode(sErrDescription) '& "&" & sURL
		'Response.Redirect "JobError.asp?ErrNum=" & CStr(lErrNumber) & "&ErrDesc=" & Server.UrlEncode(sErrDescription) & "&ReportID=" & aReportInfo(S_REPORT_ID_REPORT) & "&Page=" & aPageInfo(N_ALIAS_PAGE)
	End If

	DisplaySinglePrompt	= Err.number
	Err.Clear
End Function

Function CreateDisplayXMLForSinglePrompt(aConnectionInfo, oSinglePrompt, lOrder, oSession, aPromptInfo, sPin, lPType, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oInputs, oDisplayXML, sErrDescription)
'*******************************************************************************************************************************
'Purpose:	display a single prompt.
'Inputs:    aConnectionInfo, oSession, aPromptInfo, sPin, lPType, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oInputs
'Outputs:   Err.Number
'*******************************************************************************************************************************
    On Error Resume Next
    Dim lBlockBegin
    Dim lBlockCount
    Dim lErrNumber
    Dim sErrCode
    Dim oError
    Dim oRootXML
    Dim oInfo

    Set oRootXML = oSinglePromptQuestionXML.selectSingleNode("/")

    Select Case lPType
	Case DssXmlPromptLong, DssXmlPromptString, DssXmlPromptDouble, DssXmlPromptDate
		lErrNumber = CreateDisplayXMLForConstantPrompt(aConnectionInfo, oSinglePrompt, sPin, aPromptInfo, oSinglePromptQuestionXML, oSinglePromptTempXML, oDisplayXML, sErrDescription)

    Case DssXmlPromptObjects
        lErrNumber = CreateDisplayXMLForObjectPrompt(aConnectionInfo, oSinglePrompt, aConnectionInfo(S_TOKEN_CONNECTION), oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)

    Case DssXmlPromptElements
        lErrNumber = CreateDisplayXMLForElementPrompt(aConnectionInfo, oSinglePrompt, aConnectionInfo(S_TOKEN_CONNECTION), oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)

    Case DssXmlPromptExpression
		if oSinglePrompt.ExpressionType = DssXmlFilterAllAttributeQual then
			lErrNumber = CreateDisplayXMLForHierachicalPrompt(aConnectionInfo, oSinglePrompt, aConnectionInfo(S_TOKEN_CONNECTION), oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)
		elseif oSinglePrompt.ExpressionType = DssXmlExpressionMDXSAPVariable Then
			if isSupportedMDXPrompt(oSinglePromptQuestionXML) Then
				lErrNumber = CreateDisplayXMLForHierachicalPrompt(aConnectionInfo, oSinglePrompt, aConnectionInfo(S_TOKEN_CONNECTION), oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)
			else
				lErrNumber = ERR_UNSUPPORTED_PROMPTS
			end if
		else
			lErrNumber = CreateDisplayXMLForExpressionPrompt(aConnectionInfo, oSinglePrompt, aConnectionInfo(S_TOKEN_CONNECTION), oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)
		end if

    Case DssXmlPromptDimty
		lErrNumber = CreateDisplayXMLForLevelPrompt(aConnectionInfo, oSinglePrompt, aConnectionInfo(S_TOKEN_CONNECTION), oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)
    Case Else
        Call LogErrorXML(aConnectionInfo, Err.Number, Err.Description, Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForSinglePrompt", "", "Unknown Prompt Type", LogLevelError)
        lErrNumber = ERR_CUSTOM_UNKNOWN_PROMPT_TYPE
    End Select

	'create <info> tag
    If lErrNumber = 0 Then
        Set oInfo = oRootXML.createElement("info")
        Call oDisplayXML.appendChild(oInfo)
        lErrNumber = BuildInfoforPrompt(aConnectionInfo, lOrder, CLng(sPin), aPromptInfo, oSinglePromptTempXML, oInfo)
        If lErrNumber <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForElementPrompt", "", "Error in call to BuildAvailableforElementPrompt", LogLevelTrace)
        End If
    Else
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForObjectPrompt", "", "Error creating search tag", LogLevelError)
    End If

    If lErrNumber <> 0 Then
        Call LogErrorXML(aConnectionInfo, lErrNumber, Err.Description, Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForSinglePrompt", "", "Error in call to CreateDisplayXMLForPrompt, lPType=" & lPType, LogLevelTrace)
    End If

	Set oInfo = Nothing
	Set oRootXML = Nothing
    Set oError = Nothing

    CreateDisplayXMLForSinglePrompt = lErrNumber
    Err.Clear
End Function

Function CreateDisplayXMLForConstantPrompt(aConnectionInfo, oSinglePrompt, sPin, aPromptInfo, oSinglePromptQuestionXML, oSinglePromptTempXML, oDisplayXML, sErrDescription)
'********************************************************************************************************
'Purpose:	create display XML for constant prompt.
'Input:     aConnectionInfo, sPin, aPromptInfo, oSinglePromptQuestionXML, oSinglePromptTempXML
'Output:    oDisplayXML
'********************************************************************************************************
    On Error Resume Next
    Dim oRootXML
    Dim oTempDisplayXML
    Dim sCurAnswer
    Dim oNewAnswerText
    Dim oInfo
    Dim lErrNumber
    Dim oPIF
    Dim sDisplayXML
    Dim oAvailable
    Dim aAvailable

    Call GetXMLDOM(aConnectionInfo, oTempDisplayXML, sErrDescription)

    Set oRootXML = oTempDisplayXML.selectSingleNode("/")
    'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)
    aAvailable = SplitRequest(oRequest("available_" & sPin))
    If UBound(aAvailable) > -1 Then
		For Each oAvailable In aAvailable
			oSinglePrompt.Value = CStr(oAvailable)
		Next
	End If

    sDisplayXML = oSinglePrompt.DisplayXML

    'In some cases, some prompts won't be fully
    'initialized and hence DisplayXML might be blank
    If Err.number <> NO_ERR Then
		'Check if DisplayXML is blank and restore
		'prompt to original values
		If Len(CStr(sDisplayXML)) = 0 Then
			Err.Clear
			Call oSinglePrompt.Reset()
			'Retrieving original DisplayXML
			sDisplayXML = CStr(oSinglePrompt.DisplayXML)
		End If

		'Keep error value either. Mostly, either
		'there was another error and DisplayXML isn't blank; or
		'2nd call to DisplayXML fails again; or Err.number is cleared.
		lErrNumber = Err.number
		sErrDescription = Err.Description
    End If

    If lErrNumber = NO_ERR Then
		Call oTempDisplayXML.loadXML(sDisplayXML)
		Set oDisplayXML = oTempDisplayXML.selectSingleNode("mi/pif")
	Else
		Call GetXMLDOM(aConnectionInfo, oDisplayXML, sErrDescription)
  		Call oDisplayXML.loadXML("<mi><pif pin=""" & sPin & """ errnumber="""" errdescription="""" /></mi>")
  		Set oDisplayXML = oDisplayXML.selectSingleNode("mi/pif")
	End If

    Set oRootXML = Nothing
    Set oInfo = Nothing
    Set oNewCurAnswer = Nothing
    Set  oNewAnswerText = Nothing

    CreateDisplayXMLForConstantPrompt = lErrNumber
    Err.Clear
End Function

Function CreateDisplayXMLForObjectPrompt(aConnectionInfo, oSinglePrompt, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)
'******************************************************************************************************************************
'Purpose:	create display XML for object prompt.
'Input:     aConnectionInfo, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest
'Output:    oDisplayXML
'******************************************************************************************************************************

    On Error Resume Next
    Dim oRootXML
    Dim oAvailable
    Dim oIncreFetch
    Dim oPaIDL
    Dim oPaIL
    Dim oInfo
    Dim lTotalCount
    Dim sNP
    Dim oSearch
    Dim sSearchPattern
    Dim sXSL
    Dim lBlockBegin
    Dim lBlockCount
    Dim oSearchObject
    Dim oTempDisplayXML
    Dim bDisplaySearch
    Dim sHighlight
    Dim sFDDid
    Dim sDisplayXML

    Call CO_GetBlockBegin(oSinglePromptTempXML, lBlockBegin)
    Call CO_GetBlockCount(BLOCKCOUNT_OBJPROMPT, lBlockCount)
    oSinglePrompt.DisplayBlockBegin = lBlockBegin
    oSinglePrompt.DisplayBlockCount = lBlockCount
    ''oSinglePrompt.Locale = GetLocale()

    Call GetXMLDOM(aConnectionInfo, oTempDisplayXML, sErrDescription)
    Set oRootXML = oTempDisplayXML.selectSingleNode("/")

    Call CO_GetSearchField(oSinglePromptTempXML, sSearchPattern)
	sSearchPattern = Trim(sSearchPattern)

	Set oSearchObject = oSinglePrompt.SearchObject
	If oSearchObject Is Nothing Then
		bDisplaySearch = False
	Else
		bDisplaySearch = True
		If Len(oSearchObject.NamePattern) > 0 And Len(sSearchPattern)=0 Then
			bDisplaySearch = False
		    If StrComp(aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE),"PromptObject_cart_HIbrowsing.xsl")=0 Then
				oSearchObject.Flags = DssXmlSearchUsesRecursive
			End If
		End If
	End If

	If bDisplaySearch then
		Set oSearch = oRootXML.createElement("search")
		Call oSearch.setAttribute("text", CStr(sSearchPattern))

		If len(sSearchPattern) > 0 then
			If InStr(1, sSearchPattern, "*", vbBinaryCompare) = 0 Then
			    sSearchPattern = "*" & sSearchPattern & "*"
			End If
		End if
		oSearchObject.NamePattern = sSearchPattern
	End if

	'HI-browse
	Call CO_GetHiLinkforObjectPrompt(oSinglePromptTempXML, sFDDid)
	if Len(sFDDid) > 0 then
		Call GetHighlightString(sFDDid, "", "", sHighlight)
		oSinglePrompt.HighlightedObjs = sHighlight
	end if

    sDisplayXML = oSinglePrompt.DisplayXML

    'In some cases, some prompts won't be fully
    'initialized and hence DisplayXML might be blank
    If Err.number <> NO_ERR Then
		'Check if DisplayXML is blank and restore
		'prompt to original values
		If Len(CStr(sDisplayXML)) = 0 Then
			Err.Clear
			Call oSinglePrompt.Reset()
			'Retrieving original DisplayXML
			sDisplayXML = CStr(oSinglePrompt.DisplayXML)
		End If

		'Keep error value either. Mostly, either
		'there was another error and DisplayXML isn't blank; or
		'2nd call to DisplayXML fails again; or Err.number is cleared.
		lErrNumber = Err.number
		sErrDescription = Err.Description
    End If

    If lErrNumber = NO_ERR Then
		Call oTempDisplayXML.loadXML(sDisplayXML)
		Set oDisplayXML = oTempDisplayXML.selectSingleNode("mi/pif")
	Else
		Call GetXMLDOM(aConnectionInfo, oDisplayXML, sErrDescription)
  		Call oDisplayXML.loadXML("<mi><pif pin=""" & sPin & """ errnumber="""" errdescription="""" /></mi>")
  		Set oDisplayXML = oDisplayXML.selectSingleNode("mi/pif")
	End If

	If Not oSearch Is Nothing then
		Call oDisplayXML.appendChild(oSearch)
	End if

    'create <increfetch> tag for cart style and derived list case only
    If lErrNumber = NO_ERR Then
        if aPromptInfo(Clng(sPin), PROMPTINFO_B_ISCART) then
            Set oPaIDL = oDisplayXML.selectSingleNode("./pa[@idl='1']")
	        If Not (oPaIDL Is Nothing) Then
                Set oIncreFetch = oRootXML.createElement("increfetch")
                Call oDisplayXML.appendChild(oIncreFetch)
                Call oIncreFetch.setAttribute("pin", sPin)

                If lErrNumber = NO_ERR Then
					sXSL = aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE)

			        If StrComp(sXSL, "PromptObject_cart_HIbrowsing.xsl", vbTextCompare) = 0  Then
						lTotalCount = oPaIDL.selectSingleNode("./mi/oi/fct").getAttribute("cc")
					Else
						lTotalCount = oPaIDL.selectSingleNode("./mi").getAttribute("cc")
					End if

                    lErrNumber = BuildIncreFetchforBrowsing(aConnectionInfo, aPromptInfo, sPin, oSinglePrompt, lTotalCount, lBlockBegin, lBlockCount, oIncreFetch)
                    If lErrNumber <> 0 Then
                        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForObjectPrompt", "", "Error in call to BuildIncreFetchforObjectBrowsing", LogLevelTrace)
                    End If
                Else
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForObjectPrompt", "", "Error setting IncreFetch", LogLevelError)
	            End If
            End If
        End If
    Else
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForObjectPrompt", "", "Error creating info tag", LogLevelError)
    End If

    Set oRootXML = Nothing
    Set oAvailable = Nothing
    Set oIncreFetch = Nothing
    Set oPaIDL = Nothing
    Set oPaIL = Nothing
    Set oInfo = Nothing

    CreateDisplayXMLForObjectPrompt = lErrNumber
    Err.Clear
End Function


Function CreateDisplayXMLForLevelPrompt(aConnectionInfo, oSinglePrompt, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayPIF, sErrDescription)
'******************************************************************************************************************************
'Purpose:	create display XML for level prompt.
'Input:     aConnectionInfo, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest
'Output:    oDisplayXML
'******************************************************************************************************************************
    On Error Resume Next
    Dim oRootXML
    Dim oSelected
    Dim oAvailable
    Dim oIncreFetch
    Dim oPaIDL
    Dim oPaIL
    Dim oInfo
    Dim lTotalCount
    Dim lErrNumber
    Dim oSO
    Dim sSearchResult
    Dim sXSL
    Dim lBlockBegin
    Dim lBlockCount
	Dim sDisplayXML
    Dim oDefaultAnswer

    Set oDisplayPIF = oSinglePromptQuestionXML.cloneNode(false)
    'Set oRootXML = oDisplayPIF.selectSingleNode("/")
    Set oRootXML = oDisplayPIF.ownerDocument

    Call CO_GetBlockBegin(oSinglePromptTempXML, lBlockBegin)
    Call CO_GetBlockCount(BLOCKCOUNT_OBJPROMPT, lBlockCount)
    oSinglePrompt.DisplayBlockBegin = lBlockBegin
    oSinglePrompt.DisplayBlockCount = lBlockCount
    'oSinglePrompt.Locale = GetLocale()
    lErrNumber = Err.number

    Set oDefaultAnswer = oRootXML.createElement("defaultAnswer")
    Call oDisplayPIF.appendChild(oDefaultAnswer)

    If oSinglePrompt.HasDefaultAnswer Then
		oDefaultAnswer.text = "1"
    Else
		oDefaultAnswer.text = "0"
    End If

	'create <available> tag
    If lErrNumber = 0 Then
        Set oAvailable = oRootXML.createElement("available")
        Call oDisplayPIF.appendChild(oAvailable)
        lErrNumber = BuildAvailableforLevelPrompt(aConnectionInfo, oSinglePromptQuestionXML, lBlockBegin, lBlockCount, oAvailable)
        If lErrNumber <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error in call to BuildAvailableforLevelPromptFromSearch", LogLevelTrace)
            'Err.Raise lErrNumber
        End If
	End if

	If lErrNumber=0 then
		lErrNumber = ChangeLevelDisplayXML(aConnectionInfo, aPromptInfo, sPin, oAvailable, oSelected)
	    If lErrNumber <> 0 Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForElementPrompt", "", "Error in call to ChangeElementDisplayXML", LogLevelTrace)
	    End If
	End if

	If lErrNumber = 0 Then
        Set oInfo = oRootXML.createElement("info")
        Call oDisplayPIF.appendChild(oInfo)
        lErrNumber = BuildInfoforPrompt(aConnectionInfo, CLng(sPin), aPromptInfo, oSinglePromptTempXML, oInfo)
        If lErrNumber <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error in call to BuildAvailableforLevelPrompt", LogLevelTrace)
        End If
    Else
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error creating search tag", LogLevelError)
    End If

	'create <increfetch> tag for cart style only
    If lErrNumber = 0 Then
		if aPromptInfo(Clng(sPin), PROMPTINFO_B_ISCART) then
            Set oPaIL = oSinglePromptQuestionXML.selectSingleNode("./pa[@il='1']")
            If (oPaIL Is Nothing) Then   'derived list
                Set oIncreFetch = oRootXML.createElement("increfetch")
                Call oDisplayPIF.appendChild(oIncreFetch)
                Call oIncreFetch.setAttribute("pin", sPin)

                If lErrNumber = 0 Then
                    lTotalCount = oAvailable.getAttribute("cc")
                    lErrNumber = BuildIncreFetchforBrowsing(aConnectionInfo, aPromptInfo, sPin, oSinglePrompt, lTotalCount, lBlockBegin, lBlockCount, oIncreFetch)
                    If lErrNumber <> 0 Then
                        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForObjectPrompt", "", "Error in call to BuildIncreFetchforObjectBrowsing", LogLevelTrace)
                    End If
                Else
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForObjectPrompt", "", "Error setting IncreFetch", LogLevelError)
	            End If
            End If
        End If
    Else
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForObjectPrompt", "", "Error creating info tag", LogLevelError)
    End If

	Exit function

	Set oSO = oSinglePromptQuestionXML.selectSingleNode("./res/so")
	If Not (oSO Is Nothing) Then    'don't care about predefined list	'all att/hier ???
		lErrNumber = ExecuteSearchforLevelPrompt(aConnectionInfo, oSinglePromptQuestionXML, oSinglePromptTempXML, oSO, sSearchPattern, lBlockBegin, lBlockCount, sSearchResult)
        If lErrNumber <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "ExecuteSearchforLevelPrompt", "", "Error in call to ExecuteSearchforLevelPrompt", LogLevelTrace)
        End If
	end if

    'create <available> tag
    If lErrNumber = 0 Then
        Set oAvailable = oRootXML.createElement("available")
        Call oDisplayPIF.appendChild(oAvailable)
        If Len(sSearchResult) > 0 Then
            lErrNumber = BuildAvailableforLevelPromptFromSearch(aConnectionInfo, oSinglePromptQuestionXML, lBlockBegin, lBlockCount, sSearchResult, oAvailable)
            If lErrNumber <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error in call to BuildAvailableforLevelPromptFromSearch", LogLevelTrace)
            End If
        Else
            lErrNumber = BuildAvailableforLevelPromptFromQuestion(aConnectionInfo, oSinglePromptQuestionXML, sToken, lBlockBegin, lBlockCount, oAvailable)
            If lErrNumber <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error in call to BuildAvailableforLevelPromptFromQuestion", LogLevelTrace)
            End If
        End If
    Else
    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error getting block", LogLevelTrace)
    End If

    'create XSL logic
    if lErrNumber=0 then
		lErrNumber = ChangeLevelDisplayXML(aConnectionInfo, oAvailable, oSelected)
        If lErrNumber <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error in call to ChangeLevelDisplayXML", LogLevelTrace)
        End If
	end if

    'create <increfetch> tag for cart style and not-predefined list
    If lErrNumber = 0 and aPromptInfo(Clng(sPin), PROMPTINFO_B_ISCART) then
        Set oPaIDL = oSinglePromptQuestionXML.selectSingleNode("./pa[@idl='1']")
        Set oPaIL = oSinglePromptQuestionXML.selectSingleNode("./pa[@il='1']")

		If Not (oSO Is Nothing) or (oSO is nothing and oPAIL is nothing ) Then    'don't care about predefined list
	    'If Not (oPaIDL Is Nothing) And (oPaIL Is Nothing) And Err.Number = 0 Then   'derived list
            Set oIncreFetch = oRootXML.createElement("increfetch")
            Call oDisplayPIF.appendChild(oIncreFetch)
            Call oIncreFetch.setAttribute("pin", sPin)

            If lErrNumber = 0 Then
                lTotalCount = oAvailable.selectSingleNode("./fct").getAttribute("cc")
                lErrNumber = BuildIncreFetchforBrowsing(aConnectionInfo, aPromptInfo, sPin, oSinglePrompt, lTotalCount, lBlockBegin, lBlockCount, oIncreFetch)
                If lErrNumber <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error in call to BuildIncreFetchforObjectBrowsing", LogLevelTrace)
                    'Err.Raise lErrNumber
                End If
            Else
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error setting IncreFetch", LogLevelError)
	        End If
        End If
    Else
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForLevelPrompt", "", "Error creating info tag", LogLevelError)
    End If

	Set oRootXML = nothing
    Set oDisplayPIF = nothing
    Set oSelected = nothing
    Set oAvailable = nothing
    Set oIncreFetch = nothing
    Set oPaIDL = nothing
    Set oPaIL = nothing
    Set oInfo = nothing
    Set oSO = nothing

	CreateDisplayXMLForLevelPrompt = lErrNumber
	Err.Clear
End Function

Function CreateDisplayXMLForElementPrompt(aConnectionInfo, oSinglePrompt, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)
'****************************************************************************************************************************************
'Purpose:	create display XML for element prompt.
'Input:     aConnectionInfo, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest
'Output:    oDisplayXML
'****************************************************************************************************************************************
    On Error Resume Next
    Dim oRootXML
    Dim oSelected
    Dim oSelectedNodes
    Dim oAvailable
    Dim oIncreFetch
    Dim lErrNumber
    Dim lTotalCount
    Dim sStyle
    Dim oLastSearch
    Dim oSearch
    Dim sLastSearch
    Dim oBrowseFormUsage
    Dim sFormUage_rfd
    Dim oFormDatatype
    Dim sDatatype
    Dim bShowSearch
    Dim oInfo
    Dim oPaIDL
    Dim oPaIL
    Dim sMatchCase
    Dim oPIF
    Dim lBlockBegin
    Dim lBlockCount
    Dim oElementSource
    Dim oTempDisplayXML
    Dim sDisplayXML
    Dim oAvailableNode
    Dim oAttrNode

    Call CO_GetBlockBegin(oSinglePromptTempXML, lBlockBegin)
    Call CO_GetBlockCount(BLOCKCOUNT_ELEPROMPT, lBlockCount)
    oSinglePrompt.DisplayBlockBegin = lBlockBegin
    oSinglePrompt.DisplayBlockCount = lBlockCount

    Call GetXMLDOM(aConnectionInfo, oTempDisplayXML, sErrDescription)

    Set oRootXML = oTempDisplayXML.selectSingleNode("/")

    'browse forms only
    If lErrNumber = 0 Then
        bShowSearch = True
        For Each oBrowseFormUsage In oSinglePromptQuestionXML.selectNodes("./or/at/mi/fi/bfs/fu")
            sFormUage_rfd = oBrowseFormUsage.getAttribute("rfd")
            Set oFormDatatype = oSinglePromptQuestionXML.selectSingleNode("./or/at/mi/in/oi[@id = '"&sFormUage_rfd&"']/fdt" )
            sDatatype = oFormDatatype.getAttribute("dt")
            Select Case Clng(sDatatype)
            Case DssXmlBaseFormDateTime, DssXmlBaseFormDate, DssXmlBaseFormTime
                bShowSearch = False
				Exit For
            End Select
        Next
    Else
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForElementPrompt", "", "Error setting forms", LogLevelError)
    End If

    'create <search> tag
    If lErrNumber = 0 And bShowSearch Then
        Set oSearch = oRootXML.createElement("search")
        Call CO_GetSearchField(oSinglePromptTempXML, sLastSearch)
        Call oSearch.setAttribute("text", CStr(sLastSearch))
        Call CO_GetMatchCase(oSinglePromptTempXML, sMatchCase)

		If (Len(sMatchCase) = 0) Then
			sMatchCase = ReadUserOption(DEFAULT_PROMPT_MATCH_CASE_OPTION)
			If Strcomp(sMatchCase,"checked",1) = 0 Then
			'Final value "1" or "0"
				sMatchCase = "1"
			Else
				sMatchCase = "0"
			End If
		End IF

		If (StrComp(sMatchCase, "1") = 0) Then
			Call oSearch.setAttribute("case", "1")
		Else
			Call oSearch.setAttribute("case", "0")
		end if
		Call SetFilterForElementPrompt(aConnectionInfo, oSinglePrompt, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, lBlockBegin, lBlockCount)
        lErrNumber = Err.number
    End If

    sDisplayXML = oSinglePrompt.DisplayXML

    'In some cases, some prompts won't be fully
    'initialized and hence DisplayXML might be blank
    If Err.number <> NO_ERR Then
		'Check if DisplayXML is blank and restore
		'prompt to original values
		If Len(sDisplayXML) = 0 Then
			Err.Clear
			Call oSinglePrompt.Reset()
			'Retrieving original DisplayXML
			sDisplayXML = oSinglePrompt.DisplayXML
		End If

		'Keep error value either. Mostly, either
		'there was another error and DisplayXML isn't blank; or
		'2nd call to DisplayXML fails again; or Err.number is cleared.
		lErrNumber = Err.number
		sErrDescription = Err.Description
    End If

    If lErrNumber = 0 Then
		Call oTempDisplayXML.loadXML(sDisplayXML)

		Set oDisplayXML = oTempDisplayXML.selectSingleNode("mi/pif")
		Call oDisplayXML.appendChild(oSearch)

		'create <increfetch> tag for cart style only
		If lErrNumber = 0 Then
			If aPromptInfo(Clng(sPin), PROMPTINFO_B_ISCART) then
				Set oPaIDL = oDisplayXML.selectSingleNode("./pa[@idl='1']")
				Set oPaIL = oDisplayXML.selectSingleNode("./pa[@il='1']")
				If Not (oPaIDL Is Nothing) And (oPaIL Is Nothing) And lErrNumber = 0 Then   'derived list
					Set oIncreFetch = oRootXML.createElement("increfetch")
					Call oDisplayXML.appendChild(oIncreFetch)
					Call oIncreFetch.setAttribute("pin", sPin)

					If lErrNumber = 0 Then
						Set oAvailable = oDisplayXML.selectSingleNode("./pa[@il='1' $or$ @idl='1']/mi/oi/es")
		                lTotalCount = oAvailable.getAttribute("cc")
			            lErrNumber = BuildIncreFetchforBrowsing(aConnectionInfo, aPromptInfo, sPin, oSinglePrompt, lTotalCount, lBlockBegin, lBlockCount, oIncreFetch)
				        If lErrNumber <> 0 Then
					        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForElementPrompt", "", "Error in call to BuildIncreFetchforBrowsing", LogLevelTrace)
						End If
					End If
				End If
			End If
  		End If

  		Set oSelectedNodes = oDisplayXML.selectNodes("./pa[@ia='1']/oi/es/e")

  		For Each oSelected In oSelectedNodes
  			Set oAvailableNode = oDisplayXML.selectSingleNode("./pa[@il='1' $or$ @idl='1']/mi/oi/es/e[@ei='" & oSelected.selectSingleNode("./@ei").text & "']")
  			If Not oAvailableNode Is Nothing Then
  				Set oAttrNode = oDisplayXML.ownerDocument.createAttribute("selected")
  				oAttrNode.value = 1
  				Call oAvailableNode.attributes.setNamedItem(oAttrNode)
  			End If
  		Next
  	Else
  		Call GetXMLDOM(aConnectionInfo, oDisplayXML, sErrDescription)
  		Call oDisplayXML.loadXML("<mi><pif pin=""" & sPin & """ errnumber="""" errdescription="""" /></mi>")
  		Set oDisplayXML = oDisplayXML.selectSingleNode("mi/pif")
  	End If

    Set oAvailableNode = Nothing
    Set oRootXML = Nothing
    Set oSelected = Nothing
    Set oAvailable = Nothing
    Set oIncreFetch = Nothing
    Set oLastSearch = Nothing
    Set oSearch = Nothing
    Set oBrowseFormUsage = Nothing
    Set oFormDatatype = Nothing
    Set oInfo = Nothing
    Set oAttrNode = Nothing

    CreateDisplayXMLForElementPrompt = lErrNumber
    Err.Clear
End Function

Function CreateDisplayXMLForExpressionPrompt(aConnectionInfo, oSinglePrompt, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)
'***********************************************************************************************************************************
'Purpose:	create display XML for expression prompt.
'Input:     aConnectionInfo, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest
'Output:    oDisplayXML
'***********************************************************************************************************************************
    On Error Resume Next
    Dim oRootXML
    Dim oAvailable
    Dim oSelected
    Dim oRes
    Dim oIncreFetch
    Dim oCurrent
    Dim oInfo
    Dim sRes
    Dim lErrNumber
    Dim lBlockBegin
    Dim lBlockCount
    Dim lTotalCount
	Dim oTempDisplayXML
	Dim oSearch
	Dim sDisplayUnknownDef
	Dim oEXP
	Dim oUnknownDef
	Dim sDisplayXML

    Call CO_GetBlockBegin(oSinglePromptTempXML, lBlockBegin)
    Call CO_GetBlockCount(BLOCKCOUNT_OBJPROMPT, lBlockCount)
    oSinglePrompt.DisplayBlockBegin = lBlockBegin
    'oSinglePrompt.DisplayBlockCount = lBlockCount
    oSinglePrompt.DisplayBlockCount = 0
    'oSinglePrompt.Locale = GetLocale()
    Call GetXMLDOM(aConnectionInfo, oTempDisplayXML, sErrDescription)

    Set oRootXML = oTempDisplayXML.selectSingleNode("/")

    sDisplayXML = oSinglePrompt.DisplayXML

    'In some cases, some prompts won't be fully
    'initialized and hence DisplayXML might be blank
    If Err.number <> NO_ERR Then
		'Check if DisplayXML is blank and restore
		'prompt to original values
		If Len(sDisplayXML) = 0 Then
			Err.Clear
			Call oSinglePrompt.Reset()
			'Retrieving original DisplayXML
			sDisplayXML = oSinglePrompt.DisplayXML
		End If

		'Keep error value either. Mostly, either
		'there was another error and DisplayXML isn't blank; or
		'2nd call to DisplayXML fails again; or Err.number is cleared.
		lErrNumber = Err.number
		sErrDescription = Err.Description
    End If

    If lErrNumber = NO_ERR Then
		Call oTempDisplayXML.loadXML(sDisplayXML)
		Set oDisplayXML = oTempDisplayXML.selectSingleNode("mi/pif")
	Else
		Call GetXMLDOM(aConnectionInfo, oDisplayXML, sErrDescription)
  		Call oDisplayXML.loadXML("<mi><pif pin=""" & sPin & """ errnumber="""" errdescription="""" /></mi>")
  		Set oDisplayXML = oDisplayXML.selectSingleNode("mi/pif")
	End If

    'set unknown default flag
    Call CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)
	if StrComp(sDisplayUnknownDef, "1", vbBinaryCompare) = 0 Then
		Set oEXP = oDisplayXML.selectSingleNode("./pa[@ia='1']/exp")
    	Set oUnknownDef = oRootXML.createElement("unknowndef")
		Call oEXP.appendChild(oUnknownDef)
		call oUnknownDef.setAttribute("text", asDescriptors(267)) 'Descriptor: (default)
	End If

    'create <current> tag
    If lErrNumber = 0 Then
		Set oCurrent = oRootXML.createElement("current")
		Call oDisplayXML.appendChild(oCurrent)
		lErrNumber = BuildDefaultSelectionsforExpressionPrompt(aConnectionInfo, CStr(oSinglePrompt.ExpressionType), sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oDisplayXML, oCurrent)
        If lErrNumber <> 0 Then
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforExpressionPrompt", "", "Error in call to BuildDefaultSelectionsforExpressionPrompt", LogLevelTrace)
        End If
	Else
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforExpressionPrompt", "", "Unknown Prompt Type", LogLevelError)
    End If

    'create <increfetch> tag for cart style only
    If lErrNumber = 0 Then
        if aPromptInfo(Clng(sPin), PROMPTINFO_B_ISCART) And (oSinglePrompt.ExpressionType = DssXmlFilterSingleMetricQual) then
        	Set oIncreFetch = oRootXML.createElement("increfetch")
			Call oDisplayXML.appendChild(oIncreFetch)
			Call oIncreFetch.setAttribute("pin", sPin)
            lErrNumber = Err.number

			If lErrNumber = 0 Then
				Set oAvailable = oDisplayXML.selectSingleNode("./pa[@il='1' $or$ @idl='1']/mi")
			    lTotalCount = oAvailable.getAttribute("cc")
			    lErrNumber = BuildIncreFetchforBrowsing(aConnectionInfo, aPromptInfo, sPin, oSinglePrompt, lTotalCount, lBlockBegin, lBlockCount, oIncreFetch)
			    If lErrNumber <> 0 Then
			        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForExpressionPrompt", "", "Error in call to BuildIncreFetchforBrowsing", LogLevelTrace)
			    End If
			Else
			    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForExpressionPrompt", "", "Error setting IncreFetch", LogLevelError)
			End If
        End If
    Else
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForExpressionPrompt", "", "Error creating available tag", LogLevelError)
    End If

    Set oRootXML = Nothing
    Set oAvailable = Nothing
    Set oSelected = Nothing
    Set oRes = Nothing
	Set oIncreFetch = Nothing
	Set oCurrent = Nothing
	Set oInfo = Nothing

    CreateDisplayXMLForExpressionPrompt = lErrNumber
    Err.Clear
End Function

Function CreateDisplayXMLForHierachicalPrompt(aConnectionInfo, oSinglePrompt, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oDisplayXML, sErrDescription)
'***********************************************************************************************************************************
'Purpose:	create display XML for hierachical prompt.
'Input:     aConnectionInfo, sToken, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest
'Output:    oDisplayXML
'***********************************************************************************************************************************
    On Error Resume Next
    Dim oNewAnswer
    Dim oAvailable
    Dim oSelected
    Dim oRes
    Dim sRes
    Dim oNewRES
    Dim oFilterOP
    Dim oROOTXML
    Dim sFilterOP
    Dim oFilterOPText
    Dim oOP
    Dim sFNT
    Dim oIncreFetch
    Dim sStyle
    Dim oInfo
    Dim oPickHier
    Dim oCurrent
    Dim lBlockBegin
    Dim lBlockCount
    Dim lTotalCount
	Dim oTempDisplayXML
	Dim oSearch
	Dim sHIFlag
	Dim oDrill
	Dim oDown
	Dim oUp
	Dim oAttribute
	Dim sATName
	Dim sATDID
	Dim sHighlight
	Dim bShowSearch
	Dim sDDT
	Dim sDateTimeDDT
	Dim oFMOI
	Dim sSearch
	Dim oFilterXML_Drill
	Dim sStyleXSL
	Dim sHIDID
	Dim sDisplayUnknownDef
	Dim oUnknownDef
	Dim oEXP
	Dim sFolderDID
	Dim oAttributeAD
	Dim bLock
	Dim lLockLimit
	Dim oFiltered
	Dim oHierachy
	Dim oFolder
	Dim oRootNode
	Dim sFilterExp_Drill
	Dim oFilterExp_Search
	Dim bFilter_Drill
	Dim bFilter_Search
	Dim oFilterExp_Drill
	Dim sDisplayXML
	Dim oSelectedNodes
  	Dim	oAvailableNode
  	Dim	oAttrNode

	Call CO_GetBlockBegin(oSinglePromptTempXML, lBlockBegin)
    Call CO_GetBlockCount(BLOCKCOUNT_ELEPROMPT, lBlockCount)
    oSinglePrompt.DisplayBlockBegin = lBlockBegin
    'oSinglePrompt.DisplayBlockCount = lBlockCount
    oSinglePrompt.DisplayBlockCount = 0  'Elements in first instance aren't needed
    'oSinglePrompt.Locale = GetLocale()

    'If aPromptInfo(Clng(sPin), PROMPTINFO_B_ISALLDIMENSION) Then
	'	If IsEmpty(oRequest("Attribute_" & sPin)) And IsEmpty(oRequest("HIGo_" & sPin)) Then
	'		oSinglePrompt.DisplayBlockCount = 0
	'	End If
    'End If

    'HighLight String
	Call CO_GetAttributeforHIPrompt(oSinglePromptTempXML, sATName, sATDID)
	Call CO_GetHierachyDIDForHIPrompt(oSinglePromptTempXML, sHIDID)
	Call CO_GetSubFolderForPromptAllDimensions(oSinglePromptTempXML, sFolderDID)
	If Len(sATDID) > 0 or Len(sHIDID) > 0 or Len(sFolderDID) > 0 Then
		Call GetHighlightString(sFolderDID, sHIDID, sATDID, sHighlight)
		oSinglePrompt.HighlightedObjs = sHighlight
    End If

    'Get DisplayXML first time
    Call GetXMLDOM(aConnectionInfo, oTempDisplayXML, sErrDescription)
	Set oRootXML = oTempDisplayXML.selectSingleNode("/")

    sDisplayXML = oSinglePrompt.DisplayXML

    'In some cases, some prompts won't be fully
    'initialized and hence DisplayXML might be blank
    If Err.number <> NO_ERR Then
		'Check if DisplayXML is blank and restore
		'prompt to original values
		If Len(CStr(sDisplayXML)) = 0 Then
			Err.Clear
			Call oSinglePrompt.Reset()
			'Retrieving original DisplayXML
			sDisplayXML = CStr(oSinglePrompt.DisplayXML)
		End If

		'Keep error value either. Mostly, either
		'there was another error and DisplayXML isn't blank; or
		'2nd call to DisplayXML fails again; or Err.number is cleared.
		lErrNumber = Err.number
		sErrDescription = Err.Description
    End If

    Call oTempDisplayXML.loadXML(sDisplayXML)

    lErrNumber = Err.number
    If lErrNumber <> NO_ERR Then
		Call GetXMLDOM(aConnectionInfo, oDisplayXML, sErrDescription)
  		Call oDisplayXML.loadXML("<mi><pif pin=""" & sPin & """ errnumber="""" errdescription="""" /></mi>")
  		Set oDisplayXML = oDisplayXML.selectSingleNode("mi/pif")

	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErrDescription, Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForHierachicalPrompt", "", "Error getting XML for Prompts", LogLevelError)
    Else
		Set oDisplayXML = oTempDisplayXML.selectSingleNode("mi/pif")

		Set oAvailable = oDisplayXML.selectSingleNode("./pa[@idl='1' $or$ @il='1']/mi")

 		'set current folder
		If Len(sFolderDID) = 0 then
			Set oFolder = oAvailable.selectSingleNode(".//oi[@tp='8' $and$ @highlight='1' ]")
			If Not (oFolder Is Nothing) Then
				sFolderDID = oFolder.getAttribute("did")
				Call CO_SetSubFolderForPromptAllDimensions(oSinglePromptTempXML, sFolderDID)
			Else
				Err.Clear
			End If
		End If

		'set current Hierachy
		If Len(sHIDID) = 0 Then
			Set oHierachy = oAvailable.selectSingleNode(".//oi[@tp='14' $and$ @highlight='1' ]")
			If Not (oHierachy Is Nothing) Then
				sHIDID = oHierachy.getAttribute("did")
				Call CO_SetHierachyDIDForHIPrompt(oSinglePromptTempXML, sHIDID)
			Else
				Err.Clear
			End If
		End If

		'set current attribute
		If Len(sATName) = 0 Then
			Set oAttribute = oAvailable.selectSingleNode(".//oi[@tp='12' $and$ ./ad[@iep='1'] ]")
			If Not (oAttribute Is Nothing) Then
				sATName = oAttribute.getAttribute("disp_n")
				sATDID = oAttribute.getAttribute("did")
				Call CO_SetAttributeforHIPrompt(oSinglePromptTempXML, sATName, sATDID)
			Else
				Err.Clear
			End If
		Else
			Set oAttribute = oAvailable.selectSingleNode(".//oi[@tp='12' $and$ @highlight='1']")
		End If

		'HighLight String
		Call CO_GetAttributeforHIPrompt(oSinglePromptTempXML, sATName, sATDID)
		Call CO_GetHierachyDIDForHIPrompt(oSinglePromptTempXML, sHIDID)
		Call CO_GetSubFolderForPromptAllDimensions(oSinglePromptTempXML, sFolderDID)
		If (Len(sATDID) > 0 or Len(sHIDID) > 0 or Len(sFolderDID) > 0) Then
			Call GetHighlightString(sFolderDID, sHIDID, sATDID, sHighlight)
			oSinglePrompt.HighlightedObjs = sHighlight
		End If

		Set oRootNode = oSinglePrompt.ElementSourceObject.ExpressionObject.RootNode
		bFilter_Drill = False
		bFilter_Search = False

		'Drill filter
		Call CO_GetFilterXMLForDrillInHIPrompt(aConnectionInfo, oSinglePromptTempXML, sFilterExp_Drill)
		If Len(sFilterExp_Drill) > 0 Then
			Set oFilterExp_Drill = Server.CreateObject("WebAPIHelper.DSSXMLExpression")
			Call oFilterExp_Drill.LoadFromXML(sFilterExp_Drill)
			Call oRootNode.AppendChild(oFilterExp_Drill.RootNode)
			bFilter_Drill = True
		End If

		'Search Filter
		Call CO_GetSearchField(oSinglePromptTempXML, sSearch)
		If Len(sSearch) > 0 Then
		    lErrNumber = CO_BuildFilterXMLForSearchField(aConnectionInfo, oSession, oSinglePrompt, sSearch, oSinglePromptQuestionXML, oSinglePromptTempXML, SEARCHFIELD_ELEPROMPT, oFilterExp_Search)
			If lErrNumber = NO_ERR And Not(oFilterExp_Search Is Nothing) Then
				Call oRootNode.AppendChild(oFilterExp_Search.RootNode)
				bFilter_Search = True
			End If
		End If

		'Lock
		bLock  = False
		Set oAttributeAD = Nothing
		Set oAttributeAD = oAttribute.selectSingleNode("./ad")
		If Not oAttributeAD Is Nothing Then
			If Not IsNull(oAttributeAD.getAttribute("lt")) Then
				If Clng(oAttributeAD.getAttribute("lt")) = DssXmlLockCustom Then
					bLock = True
				End If
			End If
		Else
			Err.Clear
		End If

		'We need to set to block count for 2nd DisplayXML call
		If Len(oSinglePrompt.HighlightedObjs)>0 Or Len(sFilterExp_Drill)>0 Or Len(sSearch)>0 Then
			If aPromptInfo(Clng(sPin), PROMPTINFO_B_ISALLDIMENSION) And Len(oRequest("HIGo_" & sPin))=0 And _
			    Len(oRequest("DrillGO_" & sPin))=0 And Len(oRequest("AttributeGo_" & sPin))=0 And _
			    Len(oRequest("prev_" & sPin & ".x"))=0 And Len(oRequest("next_" & sPin & ".x"))=0 And _
			    Not bFilter_Search Then
				oSinglePrompt.DisplayBlockCount = 0
			Else
				oSinglePrompt.DisplayBlockCount = lBlockCount
			End If
		End If

		'Block Count
		If lErrNumber = 0 Then
			sStyleXSL = aPromptInfo(Clng(sPin), PROMPTINFO_S_XSLFILE)	'oSinglePrompt.StyleXSL
			If (StrComp(sStyleXSL, "PromptExpression_HierCart_REQsearch_drill_NOqual.xsl", vbTextCompare) = 0) Or (StrComp(sStyleXSL, "PromptExpression_HierCart_REQsearch_drill_qual.xsl", vbTextCompare) = 0) Then
				If not bFilter_Search And Not bFilter_Drill Then
					oSinglePrompt.DisplayBlockCount = 0
					bLock = True
				End If
			End If
			If Not bLock Then
				lLockLimit = CLng(oAttributeAD.getAttribute("ll"))
				If (Err.number = 0) And (lLockLimit <> 0) Then
					oSinglePrompt.DisplayBlockCount = lLockLimit
				Else
					Err.Clear
				End If
			End If
		End If

		'Get DisplayXML second time
		sDisplayXML = oSinglePrompt.DisplayXML

		'In some cases, some prompts won't be fully
		'initialized and hence DisplayXML might be blank
		If Err.number <> NO_ERR Then
			'Check if DisplayXML is blank and restore
			'prompt to original values
			If Len(CStr(sDisplayXML)) = 0 Then
				Err.Clear
				Call oSinglePrompt.Reset()
				'Retrieving original DisplayXML
				sDisplayXML = CStr(oSinglePrompt.DisplayXML)
			End If

			'Keep error value either. Mostly, either
			'there was another error and DisplayXML isn't blank; or
			'2nd call to DisplayXML fails again; or Err.number is cleared.
			lErrNumber = Err.number
			sErrDescription = Err.Description
		End If

		Call oTempDisplayXML.loadXML(sDisplayXML)

		Set oDisplayXML = oTempDisplayXML.selectSingleNode("mi/pif")
		If aPromptGeneralInfo(PROMPT_B_XML) Then
		    Response.Write "<!-- DisplayXML From Helper: " & oTempDisplayXML.xml & " -->"
		End If

		If Not oDisplayXML.selectSingleNode("./pa[@ia='1']/exp/nd/nd/nd[@et='1' $and$ @nt='2']/oi[@tp='12' $and$ @did='" & sATDID & "']") Is Nothing Then
			Set oSelectedNodes = oDisplayXML.selectNodes("./pa[@ia='1']/exp/nd/nd/nd[@et='1' $and$ @nt='2']/oi[@tp='12' $and$ @did='" & sATDID & "']/es/e")
		Else
			Set oSelectedNodes = oDisplayXML.selectNodes("./pa[@ia='1']/oi/es/e")
		End If

  		For Each oSelected In oSelectedNodes
  			Set oAvailableNode = oDisplayXML.selectSingleNode("./pa[@il='1' $or$ @idl='1']/mi/oi[@highlight='1']/mi/oi[@highlight='1']/es/e[@ei='" & oSelected.selectSingleNode("./@ei").text & "']")
  			If Not oAvailableNode Is Nothing Then
  				Set oAttrNode = oDisplayXML.ownerDocument.createAttribute("selected")
  				oAttrNode.value = 1
  				Call oAvailableNode.attributes.setNamedItem(oAttrNode)
  			End If
  		Next

		'set current attribute
		Set oAvailable = oDisplayXML.selectSingleNode("./pa[@idl='1' $or$ @il='1']/mi")
		Set oAttribute = oAvailable.selectSingleNode(".//oi[@tp='12' $and$ @highlight='1']")

		'Set "filtered" flag for current attribute
		If lErrNumber = 0 Then
		    If bFilter_Drill Then
		        Set oFiltered = oAttribute.cloneNode(True)
		        Call oAttribute.setAttribute("filtered", "1")
		        Call oFiltered.removeAttribute("highlight")
		        Call oAttribute.parentNode.insertBefore(oFiltered, oAttribute)
		    End If
   		Else
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforHIPrompt", "", "Error working with XML", LogLevelError)
		End If

		'Show Search Field or not
		If lErrNumber = 0 Then
		    bShowSearch = True
		    If Not oAttribute Is Nothing Then
				For Each oFMOI In oAttribute.selectNodes("./oi[@tp='21']")
				    sDDT = oFMOI.getAttribute("ddt")
				    Select Case Clng(sDDT)
					'Case DssXmlBaseFormDateTime, DssXmlBaseFormDate, DssXmlBaseFormTime
					Case DssXmlDataTypeTimeStamp, DssXmlDataTypeDate, DssXmlDataTypeTime
					    bShowSearch = False
						Exit For
					End Select
				Next
			End If
   		Else
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforHIPrompt", "", "Error working with XML", LogLevelError)
		End If

		'<search>
		If lErrNumber = 0 Then
		    If bShowSearch Then
		        Set oSearch = oRootXML.createElement("search")
		        Call oSearch.setAttribute("text", CStr(sSearch))
				Call oAvailable.appendChild(oSearch)
		    End If
   		Else
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforHIPrompt", "", "Error working with XML", LogLevelError)
		End If

		'set unknown default flag
		Call CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)
		If StrComp(sDisplayUnknownDef, "1", vbBinaryCompare) = 0 Then
			Set oEXP = oDisplayXML.selectSingleNode("./pa[@ia='1']/exp")
			Set oUnknownDef = oRootXML.createElement("unknowndef")
			Call oEXP.appendChild(oUnknownDef)
			call oUnknownDef.setAttribute("text", asDescriptors(267)) 'Descriptor: (default)
		End If

		'<filterHier>
		If Len(sFilterExp_Drill) > 0 Then
			Dim oFilterHier
			Set oFilterHier = oRootXML.createElement("filterHier")
			oFilterHier.text = sFilterExp_Drill
			Call oDisplayXML.appendChild(oFilterHier)
		End If

		If aPromptInfo(Clng(sPin), PROMPTINFO_B_ISALLDIMENSION) Then
		    Set oPickHier = oRootXML.createElement("pickhier")
		    Call oDisplayXML.appendChild(oPickHier)
		    call oPickHier.setAttribute("pin", sPin)
		    lErrNumber = CreateDisplayXMlforAllDimensions(aConnectionInfo, oSinglePrompt, sToken, oSession, sPin, sRes, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oPickHier)
		    If lErrNumber <> 0 Then
		        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforHierachicalPrompt", "", "Error in call to CreateDisplayXMlforAllDimensions", LogLevelTrace)
		    End If
		Else
			Call oAvailable.setAttribute("onedim", "yes")
		End if

		'Set flag
		If lErrNumber = 0 Then
			Call CO_GetHiFlagForHIPrompt(oSinglePromptTempXML, aPromptInfo(cLng(sPin), PROMPTINFO_B_ISALLDIMENSION), sHIFlag)
			Call oAvailable.setAttribute("flag", sHIFlag)
		End If

		'<increfetch>
		If (lErrNumber = 0) Then
			If (StrComp(oAvailable.getAttribute("flag"), "ELEM", vbTextCompare) = 0) Then
				Set oIncreFetch = oRootXML.createElement("increfetch")
				Call oDisplayXML.appendChild(oIncreFetch)
				Call oIncreFetch.setAttribute("pin", sPin)
				lTotalCount = oAvailable.selectSingleNode(".//oi[@tp='14']/mi/oi[@tp='12' $and$ @highlight='1']/es").getAttribute("cc")
				lErrNumber = BuildIncreFetchforBrowsing(aConnectionInfo, aPromptInfo, sPin, oSinglePrompt, lTotalCount, lBlockBegin, oSinglePrompt.DisplayBlockCount, oIncreFetch)
				If lErrNumber <> 0 Then
				    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforHIPrompt", "", "Error in call to BuildIncreFetchforBrowing", LogLevelTrace)
				End If
			End If
   		Else
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforHIPrompt", "", "Error working with XML", LogLevelError)
		End If

		'Drill
		Set oDrill = oAttribute.selectSingleNode("./ad/ar")
		If Not(oDrill Is Nothing) Then
			Set oDown = oDrill.selectSingleNode("./rc/oi[0]")
			If oDown is nothing Then
				Set oUp = oDrill.selectSingleNode("./rp/oi[0]")
				If Not(oUp is nothing) then
					Call oUp.setAttribute("highlight", "1")
				End If
			Else
				Call oDown.setAttribute("highlight", "1")
			End If
		Else
			Err.Clear
		End If

  		If aPromptGeneralInfo(PROMPT_B_XML) Then
		    Response.Write "<!-- Final DisplayXML: " & oTempDisplayXML.xml & " -->"
		End If
	End If

    Set oAvailable = Nothing
    Set oSelected = Nothing
    Set oRes = Nothing
    Set oRootXML = Nothing
	Set oIncreFetch = Nothing
	Set oPickHier = nothing
	Set oInfo = nothing

    CreateDisplayXMLForHierachicalPrompt = lErrNumber
    Err.Clear
End Function

Function BuildInfoforPrompt(aConnectionInfo, lOrder, lPin, aPromptInfo, oSinglePromptTempXML, oInfo)
'*****************************************************************************************************
'Purpose:	create <info> part of display XML for prompt
'Input:     aConnectionInfo, lPin, aPromptInfo, oSinglePromptTempXML
'Output:    oInfo
'*****************************************************************************************************
    On Error Resume Next
	Dim sErrCode

	Call oInfo.setAttribute("msg", aPromptInfo(lPin, PROMPTINFO_S_MSG))
	Call oInfo.setAttribute("step", replace(aPromptInfo(lPin, PROMPTINFO_S_STEP), "##", CStr(lOrder)))

	Call CO_GetPromptError(oSinglePromptTempXML, sErrCode)
	If CLng(sErrCode) > 0 Then 'And CStr(sErrCode) <> "0" Then 'And aPromptGeneralInfo(PROMPT_B_NEEDPROCESS) Then
		Call oInfo.setAttribute("error", asDescriptors(CLng(sErrCode)))
	End if

    'set lOrder digit(s) for displaying images
    Call oInfo.setAttribute("order", Cstr(lOrder))

    If aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT)>1 Then
		If lOrder>=10 Then
			Call oInfo.setAttribute("digit1", Left(Cstr(lOrder), 1))
			Call oInfo.setAttribute("digit2", Right(Cstr(lOrder), 1))
		Else
			If aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT)>=10 And aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) Then
				Call oInfo.setAttribute("digit1", "E")
				Call oInfo.setAttribute("digit2", Left(Cstr(lOrder), 1))
			Else
				Call oInfo.setAttribute("digit1", Left(Cstr(lOrder), 1))
			End if
		End if
	End if

	If lOrder>1 and aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) Then
		Call oInfo.setAttribute("totop", "1")
	End if

    BuildInfoforPrompt = Err.Number
    Err.Clear
End Function

Function SetFilterForElementPrompt(aConnectionInfo, oSinglePrompt, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, lBlockBegin, lBlockCount)
'*****************************************************************************************************
'Purpose:   create Derived Answer list of lBlockCount Elements beginning at lBlockBegin, from oElemServer
'Input:     aConnectionInfo, sToken, oSession, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, lBlockBegin, lBlockCount
'Output:    oIDLAnswerXML
'*****************************************************************************************************
    On Error Resume Next
    Dim sSearch

    Call CO_GetSearchField(oSinglePromptTempXML, sSearch)

    If Len(CStr(sSearch)) > 0 Then
        'lErrNumber = CO_BuildFilterXMLForSearchField(aConnectionInfo, oSession, oSinglePrompt, sSearch, oSinglePromptQuestionXML, oSinglePromptTempXML, SEARCHFIELD_ELEPROMPT, oFilterExp_Search)
        lErrNumber = CO_BuildFilterXMLForSearchField(aConnectionInfo, oSession, oSinglePrompt, sSearch, oSinglePromptQuestionXML, oSinglePromptTempXML, SEARCHFIELD_ELEPROMPT, oSinglePrompt.ElementSourceObject.ExpressionObject)
        If lErrNumber <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "SetFilterForElementPrompt", "", "Error in call to CO_BuildFilterXMLForSearchField", LogLevelTrace)
        End If
    End If
       ''''hydra''''''
	If Len(aPromptGeneralInfo(PROMPT_S_SECURITY_FILTERID)) > 0 Then
		lErrNumber = BuildSecurityFilter(aConnectionInfo, oSession, aPromptGeneralInfo, oSinglePrompt,oSinglePrompt.ElementSourceObject.ExpressionObject)
		If lErrNumber <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "SetFilterForElementPrompt", "", "Error in call to BuildSecurityFilter", LogLevelTrace)
        End If
	End if


    SetFilterForElementPrompt = lErrNumber
    Err.Clear
End Function


Function BuildDefaultSelectionsforExpressionPrompt(aConnectionInfo, sRes, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oDisplay, oCurrent)
'*****************************************************************************************************
'Purpose:	Set default selected items in display XML for expression prompt from oSinglePromptQuestionXML
'Input:     aConnectionInfo, sRes, oSinglePromptQuestionXML, oSinglePromptTempXML, oAvailable
'Output:    oAvailable
'*****************************************************************************************************
	On Error Resume Next
	Dim oExpItem
	Dim oRemove
	Dim oDefaultMT
	Dim oDefaultFM
	Dim oDefaultAT
	Dim oRootXML
	Dim oSO
	Dim sDefaultMT
    Dim sDefaultAT
    Dim sDefaultFM
    Dim sDefaultOP
    Dim lDefaultOP
    Dim sValue
	Dim lErrNumber
	Dim oNDNodes
	Dim oND
	Dim oAvailable
	Dim oOldDefaultFM
	Dim oOldDefaultMT
	Dim sOPTP
	Dim sDisplayUnknownDef
	sDefaultAT = ""
	sDefaultMT = ""
	sDefaultOP = ""
    Set oRootXML = oCurrent.selectSingleNode("/")
	Set oRemove = oSinglePromptTempXML.selectSingleNode("./temp/remove")
	Set oExpItem = oDisplay.selectSingleNode("(pa[@ia='1']/exp/nd/nd)[end()]")
	Set oAvailable = oDisplay.selectSingleNode("pa[@il='1' $or$ @idl='1']/mi")
	Set oDefaultAT = nothing
	Set oDefaultFM = nothing
	Set oDefaultMT = nothing


	Call CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)

	lErrNumber = Err.number
	Select Case Clng(sRes)
	Case DssXmlFilterSingleMetricQual
		If Not(oRemove Is Nothing) Then
			sDefaultMT = oRemove.getAttribute("mtdid")
			Set oOldDefaultMT = oAvailable.selectSingleNode("./oi[@highlight='1']")
			'Call oOldDefaultMT.setAttribute("highlight","0")
			Call oOldDefaultMT.removeAttribute("highlight")

			Set oDefaultMT = oAvailable.selectSingleNode("./oi[@did='" & sDefaultMT & "']")
			sDefaultOP = oRemove.getAttribute("op")
			Call oCurrent.setAttribute("op", sDefaultOP)
		ElseIf (oExpItem Is Nothing) Or (StrComp(sDisplayUnknownDef, "1", vbBinaryCompare) = 0) Then
			Call oCurrent.setAttribute("op", OperatorType_Metric & CStr(DssXmlFunctionGreater))
		Else
			sDefaultMT = oExpItem.selectSingleNode("./nd/oi[@tp='4']").getAttribute("did")
			sOPTP = oExpItem.getAttribute("optp") & ""
			sDefaultOP = oExpItem.selectSingleNode("./op").getAttribute("fnt")
			Set oOldDefaultMT = oAvailable.selectSingleNode("./oi[@highlight='1']")
			'Call oOldDefaultMT.setAttribute("highlight","0")
			Call oOldDefaultMT.removeAttribute("highlight")

			Set oDefaultMT = oAvailable.selectSingleNode("./oi[@did='" & sDefaultMT & "']")
			Select Case sOPTP
				Case "2" 'DssXmlOperatorRank
					sOPTP = OperatorType_Rank
				Case "3" 'DssXmlOperatorPercent
					sOPTP = OperatorType_Percent
				Case else
					sOPTP = OperatorType_Metric
			End Select
			Call oCurrent.setAttribute("op", sOPTP & sDefaultOP)
		End If

		If Not(oDefaultMT is nothing) Then
			Call oDefaultMT.setAttribute("highlight", "1")
		End if
		lErrNumber = Err.number

	 Case DssXmlFilterAttributeIDQual, DssXmlFilterAttributeDESCQual, DssXmlFilterAllAttributeQual

		If Not(oRemove Is Nothing) Then
			sDefaultAT = oRemove.getAttribute("atdid")
			sDefaultFM = oRemove.getAttribute("fmdid")
			If Clng(sRes) = DssXmlFilterAllAttributeQual Then
				'''' !!! XML could be different
				Set oDefaultFM = oAvailable.selectSingleNode("./oi[@did='" & sDefaultAT & "']/oi[@did='" & sDefaultFM & "']")
			Else
				Set oDefaultFM = oAvailable.selectSingleNode("./oi[@did='" & sDefaultAT & "']/oi[@did='" & sDefaultFM & "']")
			End If
			sDefaultOP = oRemove.getAttribute("op")
			Call oCurrent.setAttribute("op", sDefaultOP)
        ElseIf (oExpItem Is Nothing) Or (StrComp(sDisplayUnknownDef, "1", vbBinaryCompare) = 0) Then
		    If StrComp(aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE), "PromptExpression_textbox.xsl")=0 Then
				Call oCurrent.setAttribute("op", OperatorType_Metric & CStr(DssXmlFunctionIn))
			Else
				Call oCurrent.setAttribute("op", OperatorType_Metric & CStr(DssXmlFunctionEquals))
			End If
        Else
			sDefaultAT = oExpItem.selectSingleNode("./nd/oi[@tp='12']").getAttribute("did")
			sDefaultFM = oExpItem.selectSingleNode("./nd/oi[@tp='21']").getAttribute("did")
			Set oOldDefaultFM = oAvailable.selectSingleNode("./oi/oi[@highlight='1']")
			'Call oOldDefaultFM.setAttribute("highlight", "0")
			Call oOldDefaultFM.removeAttribute("highlight")

			If Clng(sRes) = DssXmlFilterAllAttributeQual Then
				'''' !!! XML could be different
				Set oDefaultFM = oAvailable.selectSingleNode("./oi[@did='" & sDefaultAT & "']/oi[@did='" & sDefaultFM & "']")
			Else
				Set oDefaultFM = oAvailable.selectSingleNode("./oi[@did='" & sDefaultAT & "']/oi[@did='" & sDefaultFM & "']")
			End If

			If StrComp(aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE), "promptexpression_textbox.xsl", vbTextCompare) = 0 Then
				Dim sAnswer

				sDefaultOP = OperatorType_Metric & CStr(DssXmlFunctionIn)
				Call oCurrent.setAttribute("op", sDefaultOP)

				If Not oExpItem.selectSingleNode("./op/@fnt") Is Nothing Then
					Dim oCst
					Dim oCstNodes
					Dim lCount

					Set oCstNodes = oExpItem.selectNodes("./nd/cst")

					lCount = 0
					For Each oCst In oCstNodes
						If lCount > 0 Then
							sAnswer = sAnswer & ";" & oCst.text
						Else
							sAnswer = oCst.text
						End If

						lCount = lCount + 1
					Next

					Call oCurrent.setAttribute("disp_n", sAnswer)
				Else
					sAnswer = oExpItem.selectSingleNode("./nd[1]").text
					Call oCurrent.setAttribute("disp_n", sAnswer)
				End If
			Else
				sDefaultOP = oExpItem.selectSingleNode("./op").getAttribute("fnt")
				Call oCurrent.setAttribute("op", OperatorType_Metric &  sDefaultOP)
			End If

			If Not(oDefaultFM is nothing) Then
				Call oDefaultFM.setAttribute("highlight", "1")
			End If
			lErrNumber = Err.number
		End If
	Case Else
        lErrNumber = ERR_CUSTOM_UNKNOWN_EXPPROMPT_TYPE
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildDefaultSelectionsforExpressionPrompt", "", "Unknown Prompt Type", LogLevelError)
    End Select


    lDefaultOP = CLng(oExpItem.selectSingleNode("./op/@fnt").text)

    If lDefaultOP=DssXmlFunctionIn Or lDefaultOP=DssXmlFunctionNotIn Then
		Set oNDNodes = oExpItem.selectNodes("./nd[@et='1' and @nt='3']")

		If oNDNodes.length > 0 Then
			lCount = 1
			For Each oND In oNDNodes
				If lCount > 1 Then
					sValue = sValue & ";" & oND.selectSingleNode("./cst").text
				Else
					sValue = oND.selectSingleNode("./cst").text
				End If

				lCount = lCount + 1
			Next
		Else
			sValue = ""
		End If

		Dim oCstValueAttribute

		Set oCstValueAttribute = oExpItem.ownerDocument.createAttribute("cstvalue")
		oCstValueAttribute.text = sValue

		Call oExpItem.selectSingleNode("./op").attributes.setNamedItem(oCstValueAttribute)
    End If

	Set oExpItem = Nothing
	Set oRemove = Nothing
	Set oDefaultMT = Nothing
	Set oDefaultAT = nothing
	Set oDefaultFM = nothing
	'Set oCurrent = Nothing
	BuildDefaultSelectionsforExpressionPrompt = lErrNumber
	Err.Clear
End Function

Function BuildIncreFetchforBrowsing(aConnectionInfo, aPromptInfo, sPin, oSinglePrompt, lTotalCount, lStart, lBlockCount, oIncreFetch)
'*****************************************************************************************************
'Purpose:	create <increfetch> part of display XML for prompt
'Input:     aConnectionInfo, aPromptInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, lTotalCount, oRequest, lLock_Limit, lBlockCountPromptType
'Output:    oIncreFetch
'*****************************************************************************************************
    On Error Resume Next
    Dim lPCount
    Dim lPLink
    Dim lCStart
    Dim lCEnd
    Dim lNCount
    Dim lNLink
    Dim lRemaining
    'Dim lBlockCount
    'Dim lStart
    Dim oROOTDisplay
    Dim oPrev
    Dim oCurr
    Dim oNext
    Dim sTitle
    Dim sTemp

    'Call CO_GetBlockBegin(oSinglePromptTempXML, lStart)
    'Call CO_GetBlockCount(lBlockCountPromptType, lBlockCount)
    'If lLock_Limit <> 0 Then
	'	lBlockCount = lLock_Limit
    'End If

    Select Case oSinglePrompt.PromptType
    Case DssXmlPromptObjects
		sTemp = asDescriptors(56) 'Descriptor: object(s)
    Case DssXmlPromptElements
		sTemp = asDescriptors(516) 'Descriptor: elements
	Case DssXmlPromptExpression
		Select Case oSinglePrompt.ExpressionType
		Case DssXmlFilterSingleMetricQual
			sTemp = asDescriptors(962) 'Descriptor: metrics
		Case DssXmlFilterAllAttributeQual
			sTemp = asDescriptors(516) 'Descriptor: elements
		Case DssXmlFilterAttributeIDQual, DssXmlFilterAttributeDESCQual
			sTemp = asDescriptors(56) 'Descriptor: object(s)
		End Select
	Case DssXmlPromptDimty
		sTemp = asDescriptors(56) 'Descriptor: object(s)
    End Select

    Set oROOTDisplay = oIncreFetch.selectSingleNode("/")

    If Err.Number = 0 Then
        Set oPrev = oROOTDisplay.createElement("prev")
        Call oIncreFetch.appendChild(oPrev)

        Set oCurr = oROOTDisplay.createElement("curr")
        Call oIncreFetch.appendChild(oCurr)

        Set oNext = oROOTDisplay.createElement("next")
        Call oIncreFetch.appendChild(oNext)
    Else
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildIncreFetchforBrowsing", "", "Error work with XML", LogLevelError)
    End If

    If CLng(lTotalCount) > 0 Then
        'Previous
        If CLng(lStart) > 1 Then
        'Are we in the first Set of objects? If so, we don't need to put the link
            If ((CLng(lStart) - CLng(lBlockCount)) <= 1) Then
                'the previous page is the first Set of objects
                lPCount = lBlockCount
                lPLink = 1
            Else
                lPCount = lBlockCount
                lPLink = CLng(lStart) - CLng(lBlockCount)
            End If
        End If

        'Current
        If ((CLng(lStart) + CLng(lBlockCount)) > CLng(lTotalCount)) Then
        'If we are in the last Set
            lCStart = lStart
            lCEnd = lTotalCount
        Else
            lCStart = lStart
            lCEnd = CLng(lStart) + CLng(lBlockCount) - 1
        End If

        'Next
        lRemaining = CLng(lTotalCount) - (CLng(lStart) + CLng(lBlockCount) - 1)

        If (CLng(lRemaining) > 0) Then
            If (CLng(lRemaining) < CLng(lBlockCount)) Then
            ' The next page is the last Set of objects
            lNCount = lRemaining
            lNLink = CLng(lStart) + CLng(lBlockCount)
            Else
            ' The next page is not the last Set of objects
            lNCount = lBlockCount
            lNLink = CLng(lStart) + CLng(lBlockCount)
            End If
        End If
    End If

    If Err.Number = 0 Then
        Call oPrev.setAttribute("count", CStr(lPCount))
        Call oPrev.setAttribute("link", CStr(lPLink))
        sTitle = Replace(asDescriptors(846), "##", CStr(lPCount)) & " " & sTemp 'Descriptor: Previous ##
		Call oPrev.setAttribute("title", sTitle)

        Call oCurr.setAttribute("start", CStr(lCStart))
        Call oCurr.setAttribute("end", CStr(lCEnd))
        Call oCurr.setAttribute("total", CStr(lTotalCount))
        If lTotalCount > 0 then
			sTitle = "(" & Replace(Replace(Replace(asDescriptors(117), "####", CStr(lTotalCount)), "###", CStr(lCEnd)), "##", CStr(lCStart)) & ")" 'Descriptor: ## - ### of ####
		Else
			sTitle = "(0 " & sTemp & ")" '(0 sth)
		End if
		Call oCurr.setAttribute("title", sTitle)

        Call oNext.setAttribute("count", CStr(lNCount))
        Call oNext.setAttribute("link", CStr(lNLink))
        sTitle = Replace(asDescriptors(847), "##", CStr(lNCount)) & " " & sTemp 'Descriptor: next
		Call oNext.setAttribute("title", sTitle)
    Else
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildIncreFetchforBrowsing", "", "Error computing parameter with increment fetch", LogLevelError)
    End If

    Set oROOTDisplay = Nothing
    Set oPrev = Nothing
    Set oCurr = Nothing
    Set oNext = Nothing

    BuildIncreFetchforBrowsing = Err.Number
    Err.Clear

End Function


Function GetExpressionTextforQual(aConnectionInfo, sATName, sFMName, sMTName, sOP, sCST, lExpType, sExpItemText)
'***************************************************************
'Purpose:   Get Exp Item Text for AQ / MQ
'Inputs:    aConnectionInfo, oExpItem, sFlag
'Outputs:   sExpItemText
'***************************************************************
    On Error Resume Next
    Dim sOPText
    Dim lErrNumber

		Select Case lExpType
			Case DssXmlFilterSingleMetricQual
				sExpItemText = CStr(sMTName)
			Case DssXmlFilterAllAttributeQual,DssXmlFilterAttributeIDQual,DssXmlFilterAttributeDESCQual
				sExpItemText = CStr(sATName) & "(" & CStr(sFMName) & ")"
		End Select

		lErrNumber = Err.number
		If lErrNumber = 0 and Left(sOP,1)= OperatorType_Metric Then
		    Select Case Clng(Right(sOP, Len(sOP)-1))
		        Case DssXmlFunctionBetween
		            sOPText = asDescriptors(696) 'Descriptor: Between
		        Case DssXmlFunctionNotBetween
		            sOPText = asDescriptors(746) 'Descriptor: Not between
		        Case DssXmlFunctionEquals
		            sOPText = "="
		        Case DssXmlFunctionNotEqual
		            sOPText = "<>"
		        Case DssXmlFunctionGreater
		            sOPText = ">"
		        Case DssXmlFunctionGreaterEqual
		            sOPText = ">="
		        Case DssXmlFunctionLess
		            sOPText = "<"
		        Case DssXmlFunctionLessEqual
		            sOPText = "<="
		        Case DssXmlFunctionLike
		            sOPText = asDescriptors(525) 'Descriptor: Like
		        Case DssXmlFunctionNotLike
		            sOPText = asDescriptors(526) 'Descriptor: Not Like
		        Case DssXmlFunctionIn
		            sOPText = asDescriptors(587) 'Descriptor: In
				Case DssXmlFunctionNotIn
		            sOPText = asDescriptors(2204) 'Descriptor: Not In
		        Case DssXmlFunctionLessEqualEnhanced
		            sOPText = ""
		        Case Else
		            Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "GetExpressionTextforQual", "", "Error working with the XML", LogLevelError)
		            GetExpressionTextforQual = Err.Number
		            Err.Clear
		            Exit Function
		    End Select
		    lErrNumber = Err.Number
		Else
		    Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "GetExpressionTextforQual", "", "Error working with XML", LogLevelError)
		End If

		'rank and percentage
		If (lErrNumber = 0) and ((Left(sOP,1)= OperatorType_Rank) or (Left(sOP,1)= OperatorType_Percent)) Then
			Select case Right(sOP, Len(sOP)-1)
			Case CStr(DssXmlMRPFunctionTop)
				sOPText = asDescriptors(529)	'Descriptor: Highest
			Case CStr(DssXmlMRPFunctionBottom)
				sOPText = asDescriptors(530)	'Descriptor: Lowest
		    End Select
		End If

		If StrComp(sOP, OperatorType_Metric & CStr(DssXmlFunctionIn)) = 0 Or _
		   StrComp(sOP, OperatorType_Metric & CStr(DssXmlFunctionNotIn)) = 0 Then
			sExpItemText = sExpItemText & " " & sOPText & " " & "(" & CStr(sCST) & ")"
		Elseif sOP = OperatorType_Metric & CStr(DssXmlFunctionBetween) or sOP= OperatorType_Metric & CStr(DssXmlFunctionNotBetween) Then
			sExpItemText = sExpItemText & " " & sOPText & " " & Replace(sCST,aPromptGeneralInfo(PROMPT_S_INSEPERATOR)," " & asDescriptors(701) & " ") 'Descriptor: and
		Else
			sExpItemText = sExpItemText & " " & sOPText & " " & CStr(sCST)
		End If
	'End If

    lErrNumber = Err.Number

    GetExpressionTextforQual = lErrNumber
    Err.Clear
End Function

Function CreateDisplayXMLForAllDimensions(aConnectionInfo, oSinglePrompt, sToken, oSession, sPin, sRes, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oPickHier)
'*****************************************************************************************************
'Purpose:	Create DisplayXML for browsing all dimensions
'Inputs:	aConnectionInfo, sToken, oSession, sPin, sRes, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest
'Outputs:	oPickHier
'*****************************************************************************************************

    On Error Resume Next
    Dim oRootXML
    Dim sDataExploreFolderID
    Dim sDataExploreFolderXML
    Dim sRootFolderDID
    Dim sRootFolder
    Dim sFolderXML
    Dim oFolder
    Dim sPathXML
    Dim oPath
    Dim oHIpath
    Dim sCurFolderDID
    Dim oCurFolderOI
    Dim sCurFolder
    Dim oHierachies
    Dim oSubfolders
    Dim oFCT
    Dim oFD
    Dim sRFD
    Dim oOI
    Dim sDID
    Dim oLink
    Dim oParentLink
    Dim bStart
    Dim oAncestor
    Dim sAncestor
    Dim sAncestorRFD
    Dim oAncestorOI
    Dim sAncestorDID
    Dim sURL
    Dim oDM
    Dim sDM
    Dim sDMRFD
    Dim oDMOI
    Dim sDMDID
    Dim oHI
    Dim sHierachyXML
    Dim temArray
    Dim sHIName
    Dim sHIDid
    Dim lErrNumber
    Dim oRootFolder
    Dim oRootFCT
    Dim sCurHIDID
    Dim bFirst
    Dim oRootSubfolders
	Dim oSO
	Dim bSearchHier
	Dim sResultXML

    Set oRootXML = oPickHier.selectSingleNode("/")

    Set oSO = oSinglePromptQuestionXML.selectSingleNode("./or/so")
	bSearchHier = Not (oSO Is Nothing)

	lErrNumber = Err.number
	If not bSearchHier then
		Call CO_GetSubFolderForPromptAllDimensions(oSinglePromptTempXML, sCurFolderDID)
		lErrNumber = CO_GetSpecialFolderXML(aConnectionInfo, DssXmlFolderNameSchemaDataExplorer, sToken, oSession, sDataExploreFolderID, sDataExploreFolderXML)
		If lErrNumber <> 0 Then
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMlforAllDimensions", "", "Error in call to CO_GetSpecialFolderXML", LogLevelTrace)
		    'Err.Raise lErrNumber
		End If

		If lErrNumber = 0 Then
		    Call GetXMLDOM(aConnectionInfo, oRootFolder, sErrDescription)

		    If IsObject(oRootFolder) Then
		        Call oRootFolder.loadXML(sDataExploreFolderXML)
			    Set oRootFCT = oRootFolder.selectSingleNode("/mi/fct")
		    Else
		        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMlforAllDimensions", "", "Couldn't create XMLDOM Object", LogLevelError)
		    End If
		    lErrNumber = Err.Number
		Else
		    Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMlforAllDimensions", "", "Error in call to CO_GetSpecialFolderXML", LogLevelTrace)
		End If

		'<subfolders>
		If lErrNumber = 0 Then
		    Set oRootSubfolders = oRootFCT.selectNodes("./fd")
		    if oRootSubfolders.length = 0 Then	'if no subfolders, no subfolder list
				sCurFolderDID = sDataExploreFolderID
		    Else
				Set oSubfolders = oRootXML.createElement("subfolders")
				Call oPickHier.appendChild(oSubfolders)
				bFirst = True

				'<link fd=".." did="..">
				For Each oFD In oRootSubfolders
				     sRFD = oFD.getAttribute("rfd")
				     Set oOI = oRootFolder.selectSingleNode("/mi/in/oi[@id = '" & sRFD & "']")
				     sDID = oOI.getAttribute("did")

				     Set oLink = oRootXML.createElement("link")
				     Call oSubfolders.appendChild(oLink)

				     Call oLink.setAttribute("fd", CStr(oFD.Text))
				     Call oLink.setAttribute("did", CStr(sDID))

				     If Len(sCurFolderDID) = 0 Then
						If bFirst Then
							sCurFolderDID = sDID		'Set to first subfolder
							bFirst = False
						End If
					ElseIf sDID = sCurFolderDID Then
						Call oLink.setAttribute("selected", "1")
				     End If
				Next
				If oRootFCT.selectNodes("./dm").length > 0 Then		'others subfolder
					Set oLink = oRootXML.createElement("link")
				    Call oSubfolders.appendChild(oLink)

				    Call oLink.setAttribute("fd", asDescriptors(586)) 'Descriptor: Others
				    Call oLink.setAttribute("did", sDataExploreFolderID)

					If StrComp(sDataExploreFolderID, sCurFolderDID, vbBinaryCompare) = 0 Then
						Call oLink.setAttribute("selected", "1")
				    End If
				End If
			End If
			lErrNumber = Err.number
		Else
		    Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMlforAllDimensions", "", "Error creating hierachies", LogLevelError)
		End If
	End if

	If bSearchHier Then		'search on hierachies
		Set oFolder = oSinglePromptQuestionXML.selectSingleNode("./or")
		Set oFCT = oFolder.selectSingleNode("./mi/fct")
		If oFCT Is Nothing Then
			If oSinglePrompt.SearchObject Is Nothing Then
		'		create Helper Search Object
		'		call oNewSearch.LoadFromXML("or/so/mi").xml
		'		set oSinglePrompt.SearchObject = oNewSearch
			End If
			sResultXML = oSinglePrompt.SearchObject.GetResults
			lErrNumber = GetXMLDOM(aConnectionInfo, oFolder, sErrDescription)
			If lErrNumber <> NO_ERR Then
				Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), "", Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMLForAllDimensions", "", "Error in call to GetXMLDOM()", LogLevelTrace)
			Else
				oFolder.LoadXML(sResultXML)
				set oFCT = oFolder.selectSingleNode("/mi/fct")
			End If
		End If
    Else	'prompt on all dimensions
		if sDataExploreFolderID = sCurFolderDID then
		    Set oFolder = oRootFolder
		    Set oFCT = oRootFCT
		else
			lErrNumber = CO_SearchHierachy(aConnectionInfo, sCurFolderDID, oFolder)
			If lErrNumber <> 0 Then
			    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMlforAllDimensions", "", "Error in call to CO_GetNormalFolderXML", LogLevelTrace)
			    'Err.Raise lErrNumber
			Else
				Set oFCT = oFolder.selectSingleNode("/mi/fct")
				lErrNumber = Err.number
			End If
		end if
	End if

    '<hierachies>
    If lErrNumber = 0 Then
        Set oHierachies = oRootXML.createElement("hierachies")
        Call oPickHier.appendChild(oHierachies)
        Call CO_GetHierachyDIDForHIPrompt(oSinglePromptTempXML, sCurHIDID)
		bFirst = True

        '<hi n=".." did="..">
        For Each oDM In oFCT.selectNodes("./dm")
            sDM = oDM.Text
            sDMRFD = oDM.getAttribute("rfd")
            'Set oDMOI = oFolder.selectSingleNode("/mi/in/oi[@id = '"&sDMRFD&"']")
            Set oDMOI = oFolder.selectSingleNode("mi/in/oi[@id = '"&sDMRFD&"']")
            sDMDID = oDMOI.getAttribute("did")

            Set oHI = oRootXML.createElement("hi")
            Call oHierachies.appendChild(oHI)

            Call oHI.setAttribute("n", CStr(sDM))
            Call oHI.setAttribute("did", CStr(sDMDID))

			if Len(sCurHIDid)=0 then
				if bFirst then
					Call oHI.setAttribute("selected", "1")
					bFirst = False
				end if
			elseif sDMDID = sCurHIDID then
				Call oHI.setAttribute("selected", "1")
			end if
        Next
	Else
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CreateDisplayXMlforAllDimensions", "", "Error creating current folder", LogLevelError)
    End If

    Set oFolder = Nothing
    Set oPath = Nothing
    Set oRootXML = Nothing
    Set oHIpath = Nothing
    Set oHierachies = Nothing
    Set oSubfolders = Nothing
    Set oFCT = Nothing
    Set oFD = Nothing
    Set oOI = Nothing
    Set oLink = Nothing
    Set oAncestor = Nothing
    Set oAncestorOI = Nothing
    Set oDM = Nothing
    Set oDMOI = Nothing
    Set oHI = Nothing
    Set oParentLink = Nothing
    Set oRootFolder = nothing
    Set oRootFCT = nothing
    Set oRootSubfolders = nothing
    Set oSO = nothing

    CreateDisplayXMlforAllDimensions = lErrNumber
    Err.Clear
End Function

Function BuildInputNode(aConnectionInfo, aPromptGeneralInfo, oInputs)
'*********************************************************************
'Purpose:   Build <inputs> for displayXML
'Input:     aConnectionInfo, aPromptGeneralInfo(PROMPT_O_QUESTIONSXML), aPromptGeneralInfo(PROMPT_B_DHTML)
'Output:    oInputs
'*********************************************************************

    On Error Resume Next
    Dim oRoot
    Dim oInputsElement
    Dim sErrDescription

    Set oRoot = aPromptGeneralInfo(PROMPT_O_QUESTIONSXML).selectSingleNode("/")
    Set oInputs = oRoot.createElement("inputs")

    Set oInputsElement = oRoot.createElement("Desc_10")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(10) 'Descriptor: Search

	Set oInputsElement = oRoot.createElement("Desc_45")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(45) 'Descriptor: advanced...

	Set oInputsElement = oRoot.createElement("Desc_56")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(56) 'Descriptor: object(s)

	Set oInputsElement = oRoot.createElement("Desc_57")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(57) 'Descriptor: You cannot search on attributes of this type.

	Set oInputsElement = oRoot.createElement("Desc_846")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(846) 'Descriptor: Previous ##

    Set oInputsElement = oRoot.createElement("Desc_58")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(58) 'Descriptor: Page ## of ###

	Set oInputsElement = oRoot.createElement("Desc_847")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(847) 'Descriptor: Next ##

    Set oInputsElement = oRoot.createElement("Desc_110")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(110) 'Descriptor: Go

	Set oInputsElement = oRoot.createElement("Desc_145")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(145) 'Descriptor: Drill

    Set oInputsElement = oRoot.createElement("Desc_152")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(152) 'Descriptor: Up

	Set oInputsElement = oRoot.createElement("Desc_153")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(153) 'Descriptor: Down

	Set oInputsElement = oRoot.createElement("Desc_183")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(183) 'Descriptor: Drill to

	Set oInputsElement = oRoot.createElement("Desc_190")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(190) 'Descriptor: Remove

    Set oInputsElement = oRoot.createElement("Desc_212")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(212) 'Descriptor: Continue

	Set oInputsElement = oRoot.createElement("Desc_221")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(221) 'Descriptor: Cancel

	Set oInputsElement = oRoot.createElement("Desc_330")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(330) 'Descriptor: Back to top

    Set oInputsElement = oRoot.createElement("Desc_509")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(509) 'Descriptor: no more than

	Set oInputsElement = oRoot.createElement("Desc_510")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(510) 'Descriptor: no less than

    Set oInputsElement = oRoot.createElement("Desc_511")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(511) 'Descriptor: selections

	Set oInputsElement = oRoot.createElement("Desc_512")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(512) 'Descriptor: none

    Set oInputsElement = oRoot.createElement("Desc_513")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(513) 'Descriptor: Available

    Set oInputsElement = oRoot.createElement("Desc_514")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(514) 'Descriptor: Selected

    Set oInputsElement = oRoot.createElement("Desc_515")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(515) 'Descriptor: Find

	Set oInputsElement = oRoot.createElement("Desc_516")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(516) 'Descriptor: elements

    Set oInputsElement = oRoot.createElement("Desc_517")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(517) 'Descriptor: Metric

    Set oInputsElement = oRoot.createElement("Desc_518")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(518) 'Descriptor: Attribute

    Set oInputsElement = oRoot.createElement("Desc_519")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(519) 'Descriptor: Between (enter value1 ; value2)

	Set oInputsElement = oRoot.createElement("Desc_520")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(520) 'Descriptor: Exactly

    Set oInputsElement = oRoot.createElement("Desc_521")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(521) 'Descriptor: Greater than

	Set oInputsElement = oRoot.createElement("Desc_522")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(522) 'Descriptor: Greater than or equal to

    Set oInputsElement = oRoot.createElement("Desc_523")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(523) 'Descriptor: Less than

	Set oInputsElement = oRoot.createElement("Desc_524")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(524) 'Descriptor: Less than or equal to

    Set oInputsElement = oRoot.createElement("Desc_525")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(525) 'Descriptor: Like

	Set oInputsElement = oRoot.createElement("Desc_526")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(526) 'Descriptor: Not Like

    Set oInputsElement = oRoot.createElement("Desc_527")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(527) 'Descriptor: Value

	Set oInputsElement = oRoot.createElement("Desc_528")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(528) 'Descriptor: Is

    Set oInputsElement = oRoot.createElement("Desc_529")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(529) 'Descriptor: Highest

	Set oInputsElement = oRoot.createElement("Desc_530")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(530) 'Descriptor: Lowest

    Set oInputsElement = oRoot.createElement("Desc_531")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(531) 'Descriptor: Add

	Set oInputsElement = oRoot.createElement("Desc_532")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(532) 'Descriptor: Your selections

    Set oInputsElement = oRoot.createElement("Desc_533")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(533) 'Descriptor: Match

	Set oInputsElement = oRoot.createElement("Desc_534")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(534) 'Descriptor: All selections

    Set oInputsElement = oRoot.createElement("Desc_535")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(535) 'Descriptor: Any selection

	Set oInputsElement = oRoot.createElement("Desc_536")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(536) 'Descriptor: My selections

    Set oInputsElement = oRoot.createElement("Desc_537")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(537) 'Descriptor: Add to selections

    Set oInputsElement = oRoot.createElement("Desc_538")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(538) 'Descriptor: Search for:

    Set oInputsElement = oRoot.createElement("Desc_541")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(541) 'Descriptor: Qualify on description of attribute

    Set oInputsElement = oRoot.createElement("Desc_542")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(542) 'Descriptor: Back to select hierarchy

    Set oInputsElement = oRoot.createElement("Desc_543")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(543) 'Descriptor: Pick a hierarchy

	Set oInputsElement = oRoot.createElement("Desc_544")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(544) 'Descriptor: Other hierarchies:

    Set oInputsElement = oRoot.createElement("Desc_545")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(545) 'Descriptor: Operator

    Set oInputsElement = oRoot.createElement("Desc_546")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(546) 'Descriptor: Qualify

    Set oInputsElement = oRoot.createElement("Desc_547")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(547) 'Descriptor: Select

	Set oInputsElement = oRoot.createElement("Desc_548")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(548) 'Descriptor: Elements

    Set oInputsElement = oRoot.createElement("Desc_549")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(549) 'Descriptor: Current Folder

    Set oInputsElement = oRoot.createElement("Desc_550")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(550) 'Descriptor: Subfolders

    Set oInputsElement = oRoot.createElement("Desc_551")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(551) 'Descriptor: Folder contents

	Set oInputsElement = oRoot.createElement("Desc_552")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(552) 'Descriptor: Please enter search criteria.

    Set oInputsElement = oRoot.createElement("Desc_553")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(553) 'Descriptor: List may be long.

    Set oInputsElement = oRoot.createElement("Desc_554")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(554) 'Descriptor: There are no entry points in this hierarchy.

    Set oInputsElement = oRoot.createElement("Desc_579")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(579) 'Descriptor: After you select a hierarchy on the left, you will be able to view the attributes of that hierarchy in this section.

    Set oInputsElement = oRoot.createElement("Desc_584")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(584) 'Descriptor: Go to step 2

    Set oInputsElement = oRoot.createElement("Desc_587")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(587) 'Descriptor: In

    Set oInputsElement = oRoot.createElement("Desc_611")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(611) 'Descriptor: (Only text type allowed)

    Set oInputsElement = oRoot.createElement("Desc_612")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(612) 'Descriptor: Not exactly

    Set oInputsElement = oRoot.createElement("Desc_614")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(614) 'Descriptor: Not between (enter value1 ; value2)

    Set oInputsElement = oRoot.createElement("Desc_865")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(865) 'Descriptor: Back to parent folder

	Set oInputsElement = oRoot.createElement("Desc_875")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(875) 'Descriptor: Remove from selections

    Set oInputsElement = oRoot.createElement("Desc_898")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(898) 'Descriptor: In (enter value1; value2; ...; valueN)

	Set oInputsElement = oRoot.createElement("Desc_2394")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(2394) 'Descriptor: Not In (enter value1; value2; ...; valueN)

    Set oInputsElement = oRoot.createElement("Desc_916")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(916) 'Descriptor: This prompt requires between ## and ### selections.

    Set oInputsElement = oRoot.createElement("Desc_917")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(917) 'Descriptor: This prompt requires at least ## selections.

	Set oInputsElement = oRoot.createElement("Desc_918")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(918) 'Descriptor: This prompt cannot accept more than ## selections.

    Set oInputsElement = oRoot.createElement("Desc_936")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(936) 'Descriptor: This prompt requires exactly ## selections.

    Set oInputsElement = oRoot.createElement("Desc_960")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(960) 'Descriptor: The attribute is locked

    Set oInputsElement = oRoot.createElement("Desc_961")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(961) 'Descriptor: Element list cannot be displayed

	Set oInputsElement = oRoot.createElement("Desc_962")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(962) 'Descriptor: metrics

    Set oInputsElement = oRoot.createElement("Desc_968")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(968) 'Descriptor: This prompt requires only 1 selection.

    Set oInputsElement = oRoot.createElement("Desc_981")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(981) 'Descriptor: Warning!

    Set oInputsElement = oRoot.createElement("Desc_987")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(987) 'Descriptor: This prompt requires at least 1 selection.

	Set oInputsElement = oRoot.createElement("Desc_988")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(988) 'Descriptor: This prompt cannot accept more than 1 selection.

	Set oInputsElement = oRoot.createElement("Desc_1049")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1049) 'Descriptor: Match case

	Set oInputsElement = oRoot.createElement("Desc_1348")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1348) 'Descriptor: Load File

	Set oInputsElement = oRoot.createElement("Desc_1381")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1381) 'Descriptor: Import filter from a file:

	Set oInputsElement = oRoot.createElement("Desc_1415")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1415) 'Descriptor: Your report designer chose to have a shopping cart and text file style represent this prompt.

	Set oInputsElement = oRoot.createElement("Desc_1416")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1416) 'Descriptor: Because you have DHTML turned off, you cannot use the import text file feature.

	Set oInputsElement = oRoot.createElement("Desc_1417")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1417) 'Descriptor: You may use the shopping cart to answer this prompt.

	Set oInputsElement = oRoot.createElement("Desc_1956")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1956) 'Descriptor: January

	Set oInputsElement = oRoot.createElement("Desc_1957")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1957) 'Descriptor: February

		Set oInputsElement = oRoot.createElement("Desc_1958")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1958) 'Descriptor: March

		Set oInputsElement = oRoot.createElement("Desc_1959")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1959) 'Descriptor: April

		Set oInputsElement = oRoot.createElement("Desc_1960")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1960) 'Descriptor: May

		Set oInputsElement = oRoot.createElement("Desc_1961")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1961) 'Descriptor: June

		Set oInputsElement = oRoot.createElement("Desc_1962")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1962) 'Descriptor: July

		Set oInputsElement = oRoot.createElement("Desc_1963")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1963) 'Descriptor: August

		Set oInputsElement = oRoot.createElement("Desc_1964")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1964) 'Descriptor: September

		Set oInputsElement = oRoot.createElement("Desc_1965")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1965) 'Descriptor: October

		Set oInputsElement = oRoot.createElement("Desc_1966")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1966) 'Descriptor: November

		Set oInputsElement = oRoot.createElement("Desc_1967")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1967) 'Descriptor: December

		Set oInputsElement = oRoot.createElement("Desc_1968")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = Mid(asDescriptors(1968),1,1) 'Descriptor: Sunday

		Set oInputsElement = oRoot.createElement("Desc_1969")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = Mid(asDescriptors(1969),1,1) 'Descriptor: Monday

		Set oInputsElement = oRoot.createElement("Desc_1970")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = Mid(asDescriptors(1970),1,1) 'Descriptor: Tuesday

		Set oInputsElement = oRoot.createElement("Desc_1971")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = Mid(asDescriptors(1971),1,1) 'Descriptor: Wednesday

		Set oInputsElement = oRoot.createElement("Desc_1972")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = Mid(asDescriptors(1972),1,1) 'Descriptor: Thursday

		Set oInputsElement = oRoot.createElement("Desc_1973")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = Mid(asDescriptors(1973),1,1) 'Descriptor: Friday

		Set oInputsElement = oRoot.createElement("Desc_1974")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = Mid(asDescriptors(1974),1,1) 'Descriptor: Saturday

	Set oInputsElement = oRoot.createElement("Desc_2254")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(2254) 'Descriptor: This Prompy question cannot be displayed in MicroStrategy Web. The default answer will be used for report execution.

	Set oInputsElement = oRoot.createElement("Desc_2255")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(2255) 'Descriptor: This Prompy question cannot be displayed in MicroStrategy Web. As there is no default answer associated with it, this report may not be executed.

	Set oInputsElement = oRoot.createElement("Desc_2396")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(2396) 'Descriptor: Nothing selected

	Set oInputsElement = oRoot.createElement("Desc_2413")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(2413) 'Descriptor: All Prompt answers must be numbers.

	Set oInputsElement = oRoot.createElement("Desc_2414")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(2414) 'Descriptor: All Prompt answers must be dates in the correct date format.

	Set oInputsElement = oRoot.createElement("Desc_2462")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(2462) 'Descriptor: There are no slections available in this Prompt.

	Set oInputsElement = oRoot.createElement("Desc_1825")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = asDescriptors(1825) & "..." 'Descriptor: Browse

	Set oInputsElement = oRoot.createElement("FontFamily")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = aFontInfo(S_FAMILY_FONT)

    Set oInputsElement = oRoot.createElement("ProjectURL")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = aConnectionInfo(S_PROJECT_URL_CONNECTION)

    Set oInputsElement = oRoot.createElement("smallFont")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = CStr(aFontInfo(N_SMALL_FONT))

    Set oInputsElement = oRoot.createElement("mediumFont")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = CStr(aFontInfo(N_MEDIUM_FONT))

	Set oInputsElement = oRoot.createElement("DHTML")
	oInputs.appendChild oInputsElement
	if aPromptGeneralInfo(PROMPT_B_DHTML) then
		oInputsElement.Text = "1"
	else
		oInputsElement.Text = ""
	end if

	If aPromptGeneralInfo(PROMPT_B_DHTML) then
		Set oInputsElement = oRoot.createElement("DATE_FORMAT")
		oInputs.appendChild oInputsElement
		oInputsElement.Text = GetDateFormatForLocale(aConnectionInfo, sErrDescription)
	End if

	Set oInputsElement = oRoot.createElement("msgid")
	oInputs.appendChild oInputsElement
	oInputsElement.Text = aPromptGeneralInfo(PROMPT_S_MSGID)

	Set oInputsElement = oRoot.createElement("doc")
	oInputs.appendChild oInputsElement
	if aPromptGeneralInfo(PROMPT_B_ISDOC) then
		oInputsElement.Text = "1"
	else
		oInputsElement.Text = "0"
	end if

	If not len(ReadUserOption(ACCESSIBILITY_OPTION))>0 then
	    Set oInputsElement = oRoot.createElement("accessibilityModeOff")
		oInputs.appendChild oInputsElement
		oInputsElement.Text = CStr(ReadUserOption(ACCESSIBILITY_OPTION))
	End If

	Set oInputsElement = Nothing
    Set oRoot = Nothing

    BuildInputNode = Err.Number
    Err.Clear
End Function

Function BuildHiBrowsingforObjectPrompt(aConnectionInfo, sToken, oSession, sPin, oSinglePromptTempXML, oSinglePromptQuestionXML, sSearchResult, sSearchResultFolder, oRequest, oNewAnswer)
'*****************************************************************************************************
'Purpose:	Build hi browsing / search field part for object prompt displayXML
'Inputs:	aConnectionInfo, sToken, oSession, sPin, oSinglePromptTempXML, oSinglePromptQuestionXML, sSearchResult, sSearchResultFolder, oRequest
'Outputs:	oNewAnswer
'*****************************************************************************************************

    On Error Resume Next
    Dim oRootXML
    Dim oSO
    Dim sSearchPattern
    Dim oHIpath
    Dim lErrNumber
    Dim oHILinks
    Dim oSearchRes
    Dim oSearchResultFolder

    Set oSO = oSinglePromptQuestionXML.selectSingleNode("./res/so")
    Set oRootXML = oNewAnswer.selectSingleNode("/")

    '<hipath>
    Set oHIpath = oRootXML.createElement("hipath")
    Call oNewAnswer.appendChild(oHIpath)

    lErrNumber = Err.number
    If lErrNumber = 0 Then
        lErrNumber = BuildHIPathforObjectPrompt(aConnectionInfo, sToken, oSession, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oSO, oHIpath)
        If lErrNumber <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHiBrowsingforObjectPrompt", "", "Error in call to BuildHiPathforObjectPrompt", LogLevelTrace)
            'Err.Raise lErrNumber
        End If
    Else
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHiBrowsingforObjectPrompt", "", "Error working with XML", LogLevelError)
    End If

    Call CO_GetSearchField(oSinglePromptTempXML, sSearchPattern)

	If lErrNumber = 0 Then
		lErrNumber = Err.number
	End If

    'if search field is not empty, hi browsing is hidden
    If lErrNumber = 0 And (Len(sSearchPattern) = 0) Then
        Set oHILinks = oRootXML.createElement("hilinks")
        Call oNewAnswer.appendChild(oHILinks)

 		Call GetXMLDOM(aConnectionInfo, oSearchRes, sErrDescription)

        oSearchRes.loadXML (sSearchResultFolder)
        Set oSearchRes = oSearchRes.selectSingleNode("/mi")

		lErrNumber = Err.number
        If lErrNumber = 0 Then
            lErrNumber = BuildHILinksforObjectPrompt(aConnectionInfo, sPin, oSinglePromptQuestionXML, oSearchRes, oHILinks)
            If lErrNumber <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHiBrowsingforObjectPrompt", "", "Error in call to BuildHiPathforObjectPrompt", LogLevelTrace)
                'Err.Raise lErrNumber
            End If
		Else
			Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHiBrowsingforObjectPrompt", "", "Error working with XML", LogLevelError)
		End If
        'save to answerXML

        Call GetXMLDOM(aConnectionInfo, oSearchResultFolder, sErrDescription)
        oSearchResultFolder.loadXML (sSearchResultFolder)
        Set oSearchResultFolder = oSearchResultFolder.selectSingleNode("/mi")
        Call CO_SetFolderXMLForHIBrowse(oSinglePromptTempXML, oSearchResultFolder)

    Elseif lErrNumber <> 0 Then
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHiBrowsingforObjectPrompt", "", "Error in call to CO_GetSearchField", LogLevelTrace)
    End If

    Set oSO = Nothing
    Set oRootXML = Nothing
    Set oHIpath = Nothing
    Set oHILinks = Nothing
    Set oSearchRes = Nothing
    Set oSearchResultFolder = nothing
    BuildHiBrowsingforObjectPrompt = lErrNumber
    Err.Clear
End Function

Function BuildHIPathforObjectPrompt(aConnectionInfo, sToken, oSession, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oSO, oHIpath)
'*****************************************************************************************************
'Purpose:   build <hipath> And <link>s under it for object prompt
'Inputs:	aConnectionInfo, sToken, oSession, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oSO
'Outputs:	oHIPath
'*****************************************************************************************************

    On Error Resume Next
    Dim oRootXML
    Dim oRootFolder
    Dim sRootFolder
    Dim sRootFolderRFD
    Dim oRootFolderOI
    Dim sRootFolderDID
    Dim oLink
    Dim oParentLink
    Dim sURL
    Dim sCurFolderDID
    Dim lErrNumber
    Dim oPath
    Dim sPathXML
    Dim oAncestor
    Dim sAncestor
    Dim sAncestorRFD
    Dim sAncestorDID
    Dim oAncestorOI
    Dim oAncestors
    Dim i
    Dim oCurFolderOI
    Dim sCurFolder
    Dim oPathMI
    Dim bStart
    Dim sSearchPattern

    Set oRootXML = oHIpath.selectSingleNode("/")

    'root folder
    Set oRootFolder = oSO.selectSingleNode("./mi/sct/fd")
    if oRootFolder is nothing then
		sRootFolder = aConnectionInfo(S_PROJECT_CONNECTION)
		sRootFolderDID = ""
	else
		sRootFolder = oRootFolder.Text
		sRootFolderRFD = oRootFolder.getAttribute("rfd")
		Set oRootFolderOI =  oSO.selectSingleNode("./mi/in/oi[@id = '"&sRootFolderRFD&"']")
		sRootFolderDID = oRootFolderOI.getAttribute("did")
	end if

    '<link> for root
    If Err.Number = 0 Then
        Set oLink = oRootXML.createElement("link")
        Call oHIpath.appendChild(oLink)

        Call oLink.setAttribute("fd", CStr(sRootFolder))
        Call oLink.setAttribute("did", CStr(sRootFolderDID))

        Call CO_GetHiLinkforObjectPrompt(oSinglePromptTempXML, sCurFolderDID)
        If Len(CStr(sCurFolderDID)) = 0 Then
            sCurFolderDID = sRootFolderDID
        End If
	Else
		Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHIPathforObjectPrompt", "", "Error working with XML", LogLevelError)
    End If

    If Err.Number = 0 And sCurFolderDID = sRootFolderDID Then
        Call CO_GetSearchField(oSinglePromptTempXML, sSearchPattern)
        If Len(sSearchPattern) = 0 Then		'make it a link if search field is not empty
            Call oLink.setAttribute("cur", "1")		'no needed any more
        End If
    ElseIf Err.Number = 0 Then
        Call CO_GetPathForFolder(aConnectionInfo, sCurFolderDID, sToken, oSession, sPathXML)
        If aPromptGeneralInfo(PROMPT_B_XML) Then
            Response.Write "<!-- pathXML: " & sPathXML & " -->"         'test only
        End If

        Call GetXMLDOM(aConnectionInfo, oPath, sErrDescription)
        oPath.loadXML (sPathXML)
        If Err.Number=0 Then
            Set oPathMI = oPath.selectSingleNode("./mi")
            'save to answerXML
            Call CO_SetPathXMLForHIBrowse(oSinglePromptTempXML, oPathMI)
        Else
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHIPathforObjectPrompt", "", "Error loading pathXML", LogLevelError)
        End If

        If Err.Number = 0 Then
            bStart = False
            Set oAncestors = oPath.selectNodes("//fd")
            For i = 0 To CLng(oAncestors.length - 1)
                Set oAncestor = oAncestors.Item(i)
                sAncestor = oAncestor.Text
                sAncestorRFD = oAncestor.getAttribute("rfd")
                Set oAncestorOI =  oPath.selectSingleNode("//oi[@id = '"&sAncestorRFD&"']")
                sAncestorDID = oAncestorOI.getAttribute("did")

                If Err.Number = 0 And bStart Then
                    '<link fd=".." did=".." URL=".." />
                    Set oLink = oRootXML.createElement("link")
                    Call oHIpath.appendChild(oLink)

                    Call oLink.setAttribute("fd", CStr(sAncestor))
                    Call oLink.setAttribute("did", CStr(sAncestorDID))
                End If

                'folders higher than rootFolder shouldn't be displayed
                If sAncestorDID = sRootFolderDID Then
                    bStart = True
                End If
            Next
   		Else
			Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHIPathforObjectPrompt", "", "Error in call to CO_SetPathXMLForHIBrowse", LogLevelTrace)
		End If

        'current folder
        If Err.Number = 0 Then
            Set oLink = oRootXML.createElement("link")
            Call oHIpath.appendChild(oLink)

            Set oCurFolderOI =  oPath.selectSingleNode("//oi[@did = '"&sCurFolderDID&"']")
            sCurFolder = oCurFolderOI.getAttribute("n")

            Call oLink.setAttribute("fd", CStr(sCurFolder))
            Call oLink.setAttribute("did", CStr(sCurFolderDID))

            Call CO_GetSearchField(oSinglePromptTempXML, sSearchPattern)
            If Len(sSearchPattern) = 0 Then 'make it a link if search field is not empty
                Call oLink.setAttribute("cur", "1")
            End If
       Else
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHIPathforObjectPrompt", "", "Error working with XML", LogLevelError)
       End If

    	'Set flag for parent link
        Set oParentLink = oLink.previousSibling
        If NOT(oParentLink Is Nothing) Then
			Call oParentLink.setAttribute("parent", "1")
		End if

    End If

    Set oRootXML = Nothing
    Set oRootFolder = Nothing
    Set oRootFolderOI = Nothing
    Set oLink = Nothing
    Set oParentLink = Nothing
    Set oPath = Nothing
    Set oAncestor = Nothing
    Set oAncestorOI = Nothing
    Set oAncestors = Nothing
    Set oCurFolderOI = Nothing
    Set oPathMI = Nothing

    BuildHIPathforObjectPrompt = Err.Number
    Err.Clear
End Function

Function BuildHILinksforObjectPrompt(aConnectionInfo, sPin, oSinglePromptQuestionXML, oSearchRes, oHILinks)
'*****************************************************************************************************
'Purpose:  build HI links for object prompt
'Inputs:   aConnectionInfo, sPin, oSinglePromptQuestionXML, oSearchRes
'Outputs:  oHILinks
'*****************************************************************************************************

    On Error Resume Next
    Dim oFCT
    Dim oFD
    Dim sRFD
    Dim sDID
    Dim oLink
    Dim sURL
    Dim oRootXML
	Dim bFirst

    Set oRootXML = oHILinks.selectSingleNode("/")
    Set oFCT = oSearchRes.selectSingleNode("/mi/fct")

    '<link fd=".." did=".." URL=".." />
    If Err.Number = 0 Then
		bFirst = True
        For Each oFD In oFCT.selectNodes("./fd")
            sRFD = oFD.getAttribute("rfd")
            sDID = oSearchRes.selectSingleNode("/mi/in/oi[@id = '"&sRFD&"']").getAttribute("did")

            Set oLink = oRootXML.createElement("link")
            Call oHILinks.appendChild(oLink)

            If bFirst then
                Call oLink.setAttribute("first", "1")
                bFirst = False
            End if
            Call oLink.setAttribute("fd", CStr(oFD.Text))
            Call oLink.setAttribute("did", CStr(sDID))
        Next
    Else
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildHILinksforObjectPrompt", "", "Error working with XML", LogLevelError)
    End If

    Set oFCT = Nothing
    Set oFD = Nothing
    Set oLink = Nothing
    Set oRootXML = Nothing

    BuildHILinksforObjectPrompt = Err.Number
    Err.Clear
End Function

Function MapPromptXSL(aConnectionInfo, sXSL, oSinglePromptXSL)
'**********************************************************
'Purpose:   map sXSL into XSL object defined in global.asa
'Inputs:    aConnectionInfo, sXSL
'Outputs:   oSinglePromptXSL
'**********************************************************

    On Error Resume Next
	Dim bXSLCaching
	Dim bXSLLoaded

	bXSLCaching = False
	if bXSLCaching then

		Select Case sXSL
		Case "PromptConstant_textbox.xsl"
			Set oSinglePromptXSL = PromptExpression_textbox.xsl

		Case "PromptObject_cart.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTOBJECT_CART
		Case "PromptObject_cart_HIbrowsing.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTOBJECT_CART_HI
		Case "PromptObject_checkbox.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTOBJECT_CHECKBOX
		Case "PromptObject_radio.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTOBJECT_RADIO
		Case "PromptObject_SingleSelect_listbox.xsl"
		    Set oSinglePromptXSL = oXSL_PROMPTOBJECT_SLISTBOX
		Case "PromptObject_MultiSelect_listbox.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTOBJECT_MLISTBOX
		Case "PromptObject_pulldown.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTOBJECT_PULLDOWN

		Case "PromptElement_cart.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTELEMENT_CART
		Case "PromptElement_checkbox.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTELEMENT_CHECKBOX
		Case "PromptElement_radio.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTELEMENT_RADIO
		Case "PromptElement_SingleSelect_listbox.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTELEMENT_SLISTBOX
		Case "PromptElement_MultiSelect_listbox.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTELEMENT_MLISTBOX
		Case "PromptElement_pulldown.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTELEMENT_PULLDOWN

		Case "PromptExpression_cart.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTEXPRESSION_CART
		Case "PromptExpression_pulldown.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTEXPRESSION_PULLDOWN
		Case "PromptExpression_radio.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTEXPRESSION_RADIO
		Case "PromptExpression_SingleSelect_listbox.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTEXPRESSION_SLISTBOX
		Case "PromptExpressionTextbox.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTEXPRESSION_TEXTBOX

		Case "PromptExpression_HierCart_REQsearch_drill_NOqual.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTHI_CART_REQ
		Case "PromptExpression_HierCart_OPTsearch_drill_NOqual.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTHI_CART_OPT
		Case "PromptExpression_HierCart_REQsearch_drill_qual.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTHI_CART_REQ_QUAL
		Case "PromptExpression_HierCart_OPTsearch_drill_qual.xsl"
			Set oSinglePromptXSL = oXSL_PROMPTHI_CART_OPT_QUAL

		Case else		'customized XSL

		    Call GetXMLDOM(aConnectionInfo, oSinglePromptXSL, sErrDescription)
		    Call oSinglePromptXSL.Load(Server.MapPath(sXSL))
		    If Err.Number <> 0 Then
		        Call LogErrorXML(aConnectionInfo, Err.Number, Err.Description, Err.source, "PromptDisplayCuLib.asp", "MapPromptXSL", "", "Couldn't create XMLDOM Object", LogLevelError)
		    End If

		End Select
	else		'develop use only

	    Call GetXMLDOM(aConnectionInfo, oSinglePromptXSL, sErrDescription)

		bXSLLoaded = False
		If (Strcomp(ReadUserOption(KEEP_WHITESPACE_IN_PROMPTS_OPTION),"checked",vbTextCompare) = 0)  Then
			If strcomp(sXSL,"PromptExpression_TextFile.xsl",vbTextCompare) = 0 Then
				oSinglePromptXSL.Load(Server.MapPath("PromptExpression_TextFile_withspace.xsl"))
				bXSLLoaded = True
			End If
		End If

		If Not bXSLLoaded Then
			Call oSinglePromptXSL.Load(Server.MapPath(sXSL))
		End If

		If Err.Number <> 0 Then
		    Call LogErrorXML(aConnectionInfo, Err.Number, Err.Description, Err.source, "PromptDisplayCuLib.asp", "MapPromptXSL", "", "Couldn't create XMLDOM Object", LogLevelError)
		End If
	end if

	MapPromptXSL = Err.number
	Err.Clear
End Function


Function BuildSelectedforLevelPrompt(aConnectionInfo, oSinglePromptTempXML, oSinglePromptQuestionXML, oSelected)
'*****************************************************************************************************
'purpose:   build <selected> tag for displayXML for Level prompt
'input:     oaConnectionInfo, oSinglePromptTempXML, oSinglePromptQuestionXML
'output:    oSelected
'*****************************************************************************************************
    On Error Resume Next
	Dim oOI

	for each oOI in oSinglePromptTempXML.selectNodes("./oi")
		call oSelected.appendChild(oOI.cloneNode(true))
	next

	set oOI = nothing

    BuildSelectedforLevelPrompt = Err.Number
    Err.Clear
End Function

Function BuildAvailableforLevelPrompt(aConnectionInfo, oSinglePromptQuestionXML, lBlockBegin, lBlockCount, oAvailable)
'*****************************************************************************************************
'purpose:   build <available> tag for displayXML for Level prompt directly from questionXML
'input:     aConnectionInfo, oSinglePromptQuestionXML, sToken, lBlockBegin, lBlockCount
'output:    oAvailable
'*****************************************************************************************************
    On Error Resume Next
	Dim oPA
	Dim lErrNumber
	Dim oAnswerXML

    'check predefined list
    Set oPA = oSinglePromptQuestionXML.selectSingleNode("./pa[@il='1']")
    If Not (oPA Is Nothing) Then
		lErrNumber = GetPredefinedListForLevelPrompt(aConnectionInfo, oSinglePromptQuestionXML, oAvailable)
        If lErrNumber <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforObjectPrompt", "", "Error in call to GetPredefinedListForObjectPrompt", LogLevelTrace)
        End If
    Else
		lErrNumber = GetDerivedListForLevelPrompt(aConnectionInfo, oSinglePromptQuestionXML, lBlockBegin, lBlockCount, oAvailable)
		If lErrNumber <> 0 Then
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "BuildAvailableforObjectPrompt", "", "Error in call to GetPredefinedListForObjectPrompt", LogLevelTrace)
		End If
	End if

	Set oPA = nothing
    Set oAnswerXML = nothing
    BuildAvailableforLevelPrompt = Err.number
    Err.Clear
End Function

Function GetPredefinedListForLevelPrompt(aConnectionInfo, oSinglePromptQuestionXML, oAvailable)
'*****************************************************************************************************
'Purpose:	create Predefined Answer list XML for Level prompt, from lBlockBegin
'Input:     aConnectionInfo, oSinglePromptQuestionXML
'Output:    oILAnswerXML
'*****************************************************************************************************

    On Error Resume Next
    Dim oFCT
	Dim oIN
	Dim oRootXML
	Dim oOBJ
	Dim sRFD
	Dim oNewOI
	Dim oOI

    Set oFCT = oSinglePromptQuestionXML.selectSingleNode("./pa[@il='1']/mi/fct")
    Set oIN = oFCT.parentNode.selectSingleNode("./in")
    set oRootXML = oAvailable.selectSingleNode("/")

    for each oObj in oFCT.selectNodes("./*")
		Set oNewOI = oRootXML.ownerDocument.createElement("oi")
		Call oAvailable.appendChild(oNewOI)
		sRFD = oObj.getAttribute("rfd")
		set oOI = oIN.selectSingleNode("./oi[@id='" & sRFD &"']")
		call oNewOI.setAttribute("did", oOI.getAttribute("did"))
		call oNewOI.setAttribute("tp", oOI.getAttribute("tp"))
		call oNewOI.setAttribute("n", oOI.getAttribute("n"))
    next

    Set oFCT = Nothing
	Set oIN = nothing
	set oRootXML = nothing
	set oOBJ = nothing
	set oNewOI = nothing
	set oOI = nothing

    GetPredefinedListForLevelPrompt = Err.Number
    Err.Clear
End Function

Function GetDerivedListForLevelPrompt(aConnectionInfo, oSinglePromptQuestionXML, lBlockBegin, lBlockCount, oAvailable)
'*****************************************************************************************************
'Purpose:	create Predefined Answer list XML for Level prompt, from lBlockBegin
'Input:     aConnectionInfo, oSinglePromptQuestionXML
'Output:    oILAnswerXML
'*****************************************************************************************************
    On Error Resume Next
	Dim oObjSearch
	Dim oRES
	Dim oSearchResultsXML
	Dim sSearchResultsXML

    Dim oFCT
	Dim oIN
	Dim oRootXML
	Dim oOBJ
	Dim sRFD
	Dim oNewOI
	Dim oOI

    lErrNumber = GetSearchHelperObject(aConnectionInfo, oObjSearch, sErrDescription)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "SearchCuLib.asp", "ReceiveSearchRequest", "", "Error after calling GetSearchHelperObject", LogLevelTrace)
	End If

	set oRES = oSinglePromptQuestionXML.selectSingleNode("./res")
	If (oRES is nothing) then			'all attr/hier case
		oObjSearch.AppendType(DssXmlSubTypeAttribute)			'DssXmlTypeAttribute
		oObjSearch.AppendType(DssXmlSubTypeDimensionSystem)	'DssXmlTypeDimension
		oObjSearch.AppendType(DssXmlSubTypeDimensionUser)
	Else								'search case
		oObjSearch.LoadFromXML(oRES.selectSingleNode("./so").xml)
	End if

	oObjSearch.MaxObjects = -1
	oObjSearch.Flags = oObjSearch.Flags Or DssXmlSearchVisibleOnly
	Call oObjSearch.Submit()

	oObjSearch.BlockBegin = lBlockBegin
	oObjSearch.BlockCount = lBlockCount
	oObjSearch.ObjectFlags = oObjSearch.ObjectFlags Or DssXmlObjectFindHidden

	sSearchResultsXML = oObjSearch.GetResults(CBool(False), Application.Value("lExecCycleSleepTime"), (CLng(Server.ScriptTimeout) * 1000))
    Call GetXMLDOM(aConnectionInfo, oSearchResultsXML, sErrDescription)
    oSearchResultsXML.loadXML (sSearchResultsXML)

	Set oFCT = oSearchResultsXML.selectSingleNode("/mi/fct")
    Set oIN = oFCT.parentNode.selectSingleNode("./in")
    set oRootXML = oAvailable.selectSingleNode("/")

    for each oObj in oFCT.selectNodes("./*")
		Set oNewOI = oRootXML.ownerDocument.createElement("oi")
		Call oAvailable.appendChild(oNewOI)
		sRFD = oObj.getAttribute("rfd")
		set oOI = oIN.selectSingleNode("./oi[@id='" & sRFD &"']")
		call oNewOI.setAttribute("did", oOI.getAttribute("did"))
		call oNewOI.setAttribute("tp", oOI.getAttribute("tp"))
		call oNewOI.setAttribute("n", oOI.getAttribute("n"))
    next

	Call oAvailable.setAttribute("cc", oFCT.getAttribute("cc"))
	Call oAvailable.setAttribute("pcc", oFCT.getAttribute("pcc"))

    set oObjSearch = nothing
	set oRES = nothing
	set oSearchResultsXML = nothing
    set oFCT = nothing
	set oIN = nothing
	set oRootXML = nothing
	set oOBJ = nothing
	set oNewOI = nothing
	set oOI = nothing

	GetDerivedListForLevelPrompt = Err.number
	Err.Clear
End Function

Function GetOperatorName(sOP, sOperatorName)
'*******************************************************
'Purpose:   get operator name from operator ID
'Inputs:    sOP
'Outputs:   sOperatorName
'*******************************************************
    On Error Resume Next

	Select Case CLng(sOP)
		case DssXmlFunctionBetween
			sOperatorName = "M17"
		case DssXmlFunctionNotBetween
			sOperatorName = "M44"
		case DssXmlFunctionEquals
			sOperatorName = "M6"
		case DssXmlFunctionNotEqual
			sOperatorName = "M7"
		case DssXmlFunctionGreater
			sOperatorName = "M8"
		case DssXmlFunctionGreaterEqual
			sOperatorName = "M10"
		case DssXmlFunctionLess
			sOperatorName = "M9"
		case DssXmlFunctionLessEqual
			sOperatorName = "M11"
		case DssXmlFunctionLike
			sOperatorName = "M18"
		case DssXmlFunctionNotLike
			sOperatorName = "M43"
		case DssXmlFunctionIn
			sOperatorName = "M22"
		case else
			sOperatorName = "M6"
	End Select

	GetOperatorName = Err.number
	Err.Clear
End Function

Function GetOperatorID(sOperatorName, sOP)
'*******************************************************
'Purpose:   get operator ID from operator name
'Inputs:    sOperatorName
'Outputs:   sOP
'*******************************************************
    On Error Resume Next

	Select Case CStr(sOperatorName)
		case "M17"
			sOP = DssXmlFunctionBetween
		case "M44"
			sOP = DssXmlFunctionNotBetween
		case "M6"
			sOP = DssXmlFunctionEquals
		case "M7"
			sOP = DssXmlFunctionNotEqual
		case "M8"
			sOP = DssXmlFunctionGreater
		case "M10"
			sOP = DssXmlFunctionGreaterEqual
		case "M9"
			sOP = DssXmlFunctionLess
		case "M11"
			sOP = DssXmlFunctionLessEqual
		case "M18"
			sOP = DssXmlFunctionLike
		case "M43"
			sOP = DssXmlFunctionNotLike
		case "M22"
			sOP = DssXmlFunctionIn
		case else
			sOP = DssXmlFunctionEquals
	End Select

	GetOperatorID = Err.number
	Err.Clear
End Function


Function WritePromptError(sPin, aPromptInfo, sErrorMesg)
'*******************************************************
'Purpose:   Display general error for whole prompt page
'Inputs:    sErrorMesg
'Outputs:	Err.Number
'*******************************************************
    On Error Resume Next

	'Response.Write "<TABLE BGCOLOR=""#FFCC66"" WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR>"
	'Response.Write	"<TD BGCOLOR=""#FFCC66""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""8"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
	'Response.Write	"<TD WIDTH=""100%"" ALIGN=""LEFT""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""black""><B>" & aPromptInfo(CLng(sPin), PROMPTINFO_S_STEP)& ":</B></FONT>"
	'Response.Write	"<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""red""><B>" & sErrorMesg & "</B><BR /></FONT></TD></TR></TABLE>"	'Descriptor: Your request has timed out. Please try again later.

	If aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) > 1 Then
		Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
		Response.Write "<TD VALIGN=""TOP"" ROWSPAN=""3"">"
		Response.Write "<IMG WIDTH=""14"" HEIGHT=""22"" ALT="""" BORDER=""0"" SRC=""Images/" & sPin & "_olive.gif"" /></TD>"
		Response.Write "<TD VALIGN=""TOP"" ROWSPAN=""3""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""4"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
		Response.Write "<TD COLSPAN=""2"" BGCOLOR=""#DDDDBB"" WIDTH=""100%"" ALIGN=""LEFT"">"
		Response.Write "<A NAME=""1"" />"
		Response.Write "<FONT FACE=""Verdana,Arial,MS Sans Serif"" SIZE=""2"">" & aPromptInfo(CLng(sPin), PROMPTINFO_O_QUESTION).selectSingleNode("@ttl").text
		Response.Write "<FONT COLOR=""#CC0000""></FONT></FONT></TD>"
		Response.Write "<TD BGCOLOR=""#DDDDBB"" ALIGN=""RIGHT"" VALIGN=""TOP""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
		Response.Write "</TR></TABLE>"
	End If

	Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR>"
	Response.Write "<TD WIDTH=""11""><IMG SRC=""images/1ptrans.gif"" WIDTH=""11"" HEIGHT=""1"" BORDER=""0"" ALT="""" /></TD>"
	Response.Write "<TD VALIGN=""TOP"" ALIGN=""LEFT"" WIDTH=""23""><IMG SRC=""images/promptError_white.gif"" WIDTH=""23"" HEIGHT=""23"" ALT=""Warning!"" BORDER=""0"" /></TD>"
	Response.Write "<TD WIDTH=""4""><IMG SRC=""images/1ptrans.gif"" WIDTH=""4"" HEIGHT=""1"" BORDER=""0"" ALT="""" /></TD>"
	Response.Write "<TD VALIGN=""TOP"" ALIGN=""LEFT"">"
	Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""CC0000""><B>" & sPin & ":" & sErrorMesg & "</B><BR /></FONT>"
	Response.Write "</TD></TR></TABLE>"


	WritePromptError = Err.number
	Err.Clear
End Function


Function GetPinbyOrder(aPromptGeneralInfo, aPromptInfo, lOrder, lPin)
'************************************************************************************************
'Purpose:   put all form values in an attribute of oE
'Inputs:    aConnectionInfo, oE
'Outputs:   oE
'************************************************************************************************

    On Error Resume Next
    Dim temArray()
	Dim lTemIndex
	Dim i
	Dim oSinglePrompt

	Redim temArray(aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT))
	lTemIndex = 0
	If aPromptGeneralInfo(PROMPT_B_REQUIREDFIRST) Then
		For i = 1 to aPromptGeneralInfo(PROMPT_L_MAXPIN)
			Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(i)
			If oSinglePrompt.Used And not oSinglePrompt.Closed And oSinglePrompt.Required Then
				lTemIndex =	lTemIndex + 1
				temArray(lTemIndex) = i
			End If
		Next
		For i = 1 to aPromptGeneralInfo(PROMPT_L_MAXPIN)
			Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(i)
			If oSinglePrompt.Used And not oSinglePrompt.Closed And Not oSinglePrompt.Required Then
				lTemIndex =	lTemIndex + 1
				temArray(lTemIndex) = i
			End If
		Next
	Else
		For i = 1 to aPromptGeneralInfo(PROMPT_L_MAXPIN)
			Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(i)
			If oSinglePrompt.Used And not oSinglePrompt.Closed Then
				lTemIndex =	lTemIndex + 1
				temArray(lTemIndex) = i
			End If
		Next
	End If
	lPin = temArray(lOrder)

	GetPinbyOrder = Err.number
 	Err.Clear
End Function

Function WritePromptGeneralError(sErrorMesg)
'*******************************************************
'Purpose:   Display general error for whole prompt page
'Inputs:    sErrorMesg
'Outputs:	Err.Number
'*******************************************************
    On Error Resume Next

	Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
	Response.Write "<TD WIDTH=""11""><IMG SRC=""images/1ptrans.gif"" WIDTH=""11"" HEIGHT=""1"" BORDER=""0"" ALT="""" /></TD>"
	Response.Write "<TD VALIGN=""TOP"" ALIGN=""LEFT"" WIDTH=""23""><IMG SRC=""images/promptError_white.gif"" WIDTH=""23"" HEIGHT=""23"" ALT=""Warning!"" BORDER=""0"" /></TD>"
	Response.Write "<TD WIDTH=""4""><IMG SRC=""images/1ptrans.gif"" WIDTH=""4"" HEIGHT=""1"" BORDER=""0"" ALT="""" /></TD>"
	Response.Write "<TD VALIGN=""TOP"" ALIGN=""LEFT"">"
	Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""CC0000""><B>" & sErrorMesg & "</B><BR /></FONT>"
	Response.Write "</TD></TR></TABLE>"

	WritePromptGeneralError = Err.number
	Err.Clear
End Function

Function DisplayPromptIndex(aPromptGeneralInfo, aPromptInfo)
'*******************************************************
'Purpose:   Display prompt index part for whole prompt page
'Inputs:    aPromptGeneralInfo, aPromptInfo
'Outputs:   Err.Number
'*******************************************************
    On Error Resume Next
	Dim lOrder
	Dim lPin
	Dim lStart
	Dim lEnd
	Dim lPCount
	Dim lPLink
	Dim lCStart
	Dim lCEnd
	Dim lNCount
	Dim lNLink
	Dim sFirst
	Dim sPrev
	Dim sNext
	Dim sLast
	Dim sCurr

	Response.Write	"<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"" WIDTH=""100%"">"

	If aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) Then	'all index
		For lOrder = 1 to aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT)
			Call DisplaySinglePromptIndex(aPromptGeneralInfo, lOrder)
		Next

		If aPromptGeneralInfo(PROMPT_B_SUMMARY) Then
			Response.Write "<TR><TD BGCOLOR=""#AAAA77"" COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
			Response.Write "<TR><TD BGCOLOR=""#AAAA77"" WIDTH=""30"" ALIGN=""CENTER"" VALIGN=""TOP""><INPUT TYPE=""IMAGE"" SRC=""Images/PromptSummary.gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & asDescriptors(1066) & """ BORDER=""0"" ID=""PromptSummary"" Name=""PromptSummary"" /></A></TD><TD BGCOLOR=""#AAAA77"" VALIGN=""TOP""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF""><B>" & asDescriptors(1066) & "</B></FONT></TD></TR>" ' Descriptor: Summary of your selections
			Response.Write "<TR><TD BGCOLOR=""#AAAA77"" COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
		Else
			Response.Write "<TR><TD COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
			Response.Write "<TR><TD WIDTH=""30"" ALIGN=""CENTER"" VALIGN=""TOP""><INPUT TYPE=""IMAGE"" SRC=""Images/PromptSummary.gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & asDescriptors(1066) & """ BORDER=""0"" ID=""PromptSummary"" Name=""PromptSummary"" /></A></TD><TD VALIGN=""TOP""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(1066) & "</FONT></TD></TR>" ' Descriptor: Summary of your selections
			Response.Write "<TR><TD COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
		End If
		Response.Write "<TR><TD BGCOLOR=""#FFFFFF"" COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD></TR>"
	Else	'these 5 index
		lStart = Int(CLng(aPromptGeneralInfo(PROMPT_S_CURORDER)-1) / 5) * 5 + 1
		If aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) >= lStart+5 Then
			lEnd = lStart+4
		Else
			lEnd = aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT)
		End If
		For lOrder = lStart To lEnd
			Call DisplaySinglePromptIndex(aPromptGeneralInfo, lOrder)
		Next

		If aPromptGeneralInfo(PROMPT_B_SUMMARY) Then
			Response.Write "<TR><TD BGCOLOR=""#AAAA77"" COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
			Response.Write "<TR><TD BGCOLOR=""#AAAA77"" WIDTH=""30"" ALIGN=""CENTER"" VALIGN=""TOP""><INPUT TYPE=""IMAGE"" SRC=""Images/PromptSummary.gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & asDescriptors(1066) & """ BORDER=""0"" ID=""PromptSummary"" Name=""PromptSummary"" /></A></TD><TD BGCOLOR=""#AAAA77"" VALIGN=""TOP""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF""><B>" & asDescriptors(1066) & "</B></FONT></TD></TR>" ' Descriptor: Summary of your selections
			Response.Write "<TR><TD BGCOLOR=""#AAAA77"" COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
		Else
			Response.Write "<TR><TD COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
			Response.Write "<TR><TD WIDTH=""30"" ALIGN=""CENTER"" VALIGN=""TOP""><INPUT TYPE=""IMAGE"" SRC=""Images/PromptSummary.gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & asDescriptors(1066) & """ BORDER=""0"" ID=""PromptSummary"" Name=""PromptSummary"" /></A></TD><TD VALIGN=""TOP""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(1066) & "</FONT></TD></TR>" ' Descriptor: Summary of your selections
			Response.Write "<TR><TD COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
		End If
		'Step up and down
		Call GetIndexIncrefetch(aPromptGeneralInfo(PROMPT_L_MAXPIN), lStart, CONST_PROMTINDEX_BLOCKCOUNT, lPCount, lPLink, lCStart, lCEnd, lNCount, lNLink)
		sFirst = Replace(asDescriptors(1062), "##", CStr(lPCount))	 ' Descriptor: Previous ## steps
		sPrev = Replace(asDescriptors(846), "##", CStr(lPCount))	 ' Descriptor: Previous ##
		sNext = Replace(asDescriptors(847), "##", CStr(lNCount))	 ' Descriptor: Next ##
		sLast = Replace(asDescriptors(1063), "##", CStr(lNCount))	 ' Descriptor: Last ## steps
		sCurr = Replace(asDescriptors(1060), "####", CStr(aPromptGeneralInfo(PROMPT_L_MAXPIN)))
		sCurr = Replace(sCurr, "###", CStr(lCEnd))
		sCurr = Replace(sCurr, "##", CStr(lCStart))	 'Descriptor: Steps ## - ### of ####

		Response.Write	"<TR><TD COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""10"" ALT="""" BORDER=""0"" /></TD></TR>"
		Response.Write	"<TR><TD COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
		If lPCount > 0 Or lNCount > 0 Then
			Response.Write "<TR><TD ALIGN=""CENTER"">"
			If lPCount > 0 Then
				'first set
				Response.Write "<INPUT TYPE=""IMAGE"" NAME=""IndexFirst"" SRC=""Images/arrow_first_inc_fetch_v.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""" & sFirst & """ BORDER=""0"" /><BR />"
				'previous set
				Response.Write "<INPUT TYPE=""IMAGE"" NAME=""IndexPrev"" SRC=""Images/arrow_left_inc_fetch_v.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""" & sPrev & """ BORDER=""0"" /><BR />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IndexPrevValue"" VALUE=""" & CStr(lPLink) & """ />"
			End If
			If lNCount > 0 Then
				'next set
				Response.Write	"<INPUT TYPE=""IMAGE"" NAME=""IndexNext"" SRC=""Images/arrow_right_inc_fetch_v.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""" & sNext & """ BORDER=""0"" /><BR />"
				'last set
				Response.Write	"<INPUT TYPE=""IMAGE"" NAME=""IndexLast"" SRC=""Images/arrow_end_inc_fetch_v.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""" & sLast & """ BORDER=""0"" />"
				Response.Write  "<INPUT TYPE=""HIDDEN"" NAME=""IndexNextValue"" VALUE=""" & CStr(lNLink) & """ />"
			End If
			Response.Write	"</TD>"
			Response.Write	"<TD ALIGN=""CENTER""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & sCurr & "</FONT></TD></TR>"
		end if
	End if

	Response.Write	"<TR><TD COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""15"" ALT="""" BORDER=""0"" /></TD></TR>"
	Response.Write	"</TABLE>"

	set oSinglePrompt = nothing
  	DisplayPromptIndex = Err.number
	Err.Clear
End Function

Function DisplaySinglePromptIndex(aPromptGeneralInfo, lOrder)
'*******************************************************
'Purpose:   Display prompt index part for whole prompt page
'Inputs:    aPromptGeneralInfo, lPin
'Outputs:   Err.Number
'*******************************************************
    On Error Resume Next
	Dim bRequired
	Dim oSinglePrompt
	Dim sBackGround
	Dim sGifName
	Dim bCurrent
	Dim lPin
	Dim sOrder

	sOrder = Cstr(lOrder)
	Call GetPinbyOrder(aPromptGeneralInfo, aPromptInfo, lOrder, lPin)
	set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)

	bCurrent = not aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) And lOrder = Clng(aPromptGeneralInfo(PROMPT_S_CURORDER)) And (Not aPromptGeneralInfo(PROMPT_B_SUMMARY))
	bRequired = oSinglePrompt.Required
	sGifName = sOrder

	If bCurrent Then
		sBackGround = " BGCOLOR=""#AAAA77"" "
	Else
		sBackGround = ""
	End if

	Response.Write	"<TR><TD" & sBackGround & " COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
	Response.Write	"<TR><TD" & sBackGround & " WIDTH=""30"" ALIGN=""CENTER"" VALIGN=""TOP"">"

	If aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) And Not aPromptGeneralInfo(PROMPT_B_SUMMARY) then 'anchor for image
		Response.Write	"<A HREF=""#" & sOrder & """>"
		If Len(sOrder)=1 Then		'1 - 10
			Response.Write	"<IMG SRC=""Images/" & sOrder & ".gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & oSinglePrompt.Title & """ BORDER=""0"" />"
		Else
			Response.Write	"<IMG SRC=""Images/" & Left(sOrder,1) & ".gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & oSinglePrompt.Title & """ BORDER=""0"" />"
			Response.Write	"<IMG SRC=""Images/" & Right(sOrder,1) & ".gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & oSinglePrompt.Title & """ BORDER=""0"" />"
		End if
		Response.Write	"</A>"
	Else	'submit image
		If Len(sOrder)=1 Then		'1 - 10
			Response.Write	"<INPUT TYPE=""IMAGE"" NAME=""PromptCurr_" & sOrder & """ SRC=""Images/" & sOrder & ".gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & oSinglePrompt.Title & """ BORDER=""0"" />"
		Else
			Response.Write	"<INPUT TYPE=""IMAGE"" NAME=""PromptCurr_" & sOrder & """ SRC=""Images/" & Left(sOrder,1) & ".gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & oSinglePrompt.Title & """ BORDER=""0"" />"
			Response.Write	"<INPUT TYPE=""IMAGE"" NAME=""PromptCurr_" & sOrder & """ SRC=""Images/" & Right(sOrder,1) & ".gif"" WIDTH=""14"" HEIGHT=""22"" ALT=""" & oSinglePrompt.Title & """ BORDER=""0"" />"
		End If
	End if

	Response.Write	"</TD>"
	Response.Write	"<TD VALIGN=""TOP"" " & sBackGround & " >"

	If aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) And Not aPromptGeneralInfo(PROMPT_B_SUMMARY) then 'anchor for title
		Response.Write	"<A HREF=""#" & sOrder & """>"
	Else
		If aPromptGeneralInfo(PROMPT_B_DHTML) Then
			Response.Write	"<A HREF=""javascript:SubmitPromptIndex('PromptCurr_" & sOrder & "');"">"
		End If
	End If

	If bCurrent Then
		Response.Write	"<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF""><B>"
	Else
		Response.Write	"<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#000000"">"
	End if

	If Len(oSinglePrompt.Title) > 16 Then
		Response.Write Left(oSinglePrompt.Title, 16) & "..."
	Else
		Response.Write oSinglePrompt.Title
	End If

	If bCurrent Then
		Response.Write	"</B>"
	End If

	Response.Write	"</FONT>"

	'If aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) then 'anchor for title
		Response.Write	"</A>"
	'End if

	Response.Write	"<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#CC0000""><BR />"

	If bRequired Then
		Response.Write asDescriptors(661) 'Descriptor: Required
	End If
	Response.Write	"</FONT></TD></TR>"
	Response.Write	"<TR><TD COLSPAN=""2"" " & sBackGround & " ><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" ALT="""" BORDER=""0"" /></TD></TR>"
	Response.Write	"<TR><TD BGCOLOR=""#FFFFFF"" COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD></TR>"

	DisplaySinglePromptIndex = Err.number
	Err.Clear
End Function

Function GetIndexIncrefetch(lTotalCount, lStart, lBlockCount, lPCount, lPLink, lCStart, lCEnd, lNCount, lNLink)
'*******************************************************
'Purpose:   Get Index IncreFetch info
'Inputs:    lTotalCount, lStart, lBlockCount,
'Outputs:   lPCount, lPLink, lCStart, lCEnd, lNCount, lNLink
'*******************************************************
    On Error Resume Next
	Dim lRemaining

	If CLng(lTotalCount) > 0 Then
        'Previous
        If CLng(lStart) > 1 Then
        'Are we in the first Set of objects? If so, we don't need to put the link
            If ((CLng(lStart) - CLng(lBlockCount)) <= 1) Then
                'the previous page is the first Set of objects
                lPCount = lBlockCount
                lPLink = 1
            Else
                lPCount = lBlockCount
                lPLink = CLng(lStart) - CLng(lBlockCount)
            End If
        End If

        'Current
        If ((CLng(lStart) + CLng(lBlockCount)) > CLng(lTotalCount)) Then
        'If we are in the last Set
            lCStart = lStart
            lCEnd = lTotalCount
        Else
            lCStart = lStart
            lCEnd = CLng(lStart) + CLng(lBlockCount) - 1
        End If

        'Next
        lRemaining = CLng(lTotalCount) - (CLng(lStart) + CLng(lBlockCount) - 1)

        If (CLng(lRemaining) > 0) Then
            If (CLng(lRemaining) < CLng(lBlockCount)) Then
            ' The next page is the last Set of objects
            lNCount = lRemaining
            lNLink = CLng(lStart) + CLng(lBlockCount)
            Else
            ' The next page is not the last Set of objects
            lNCount = lBlockCount
            lNLink = CLng(lStart) + CLng(lBlockCount)
            End If
        End If
    End If

	GetIndexIncrefetch = Err.number
	Err.Clear
End Function

Function DisplayPromptSummary(aPromptGeneralInfo, aPromptInfo)
'*******************************************************
'Purpose:   Display prompt summary
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
    On Error Resume Next
    Dim lOrder
	Dim lPin
	Dim sSinglePromptSummary
	Dim sDisplayUnknownDef
	Dim sDefault
	Dim lErrNumber
	Dim oSinglePromptTempXML
	Dim oSinglePrompt
	Dim sTitle

	Response.Write	"<TABLE WIDTH=""98%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR>"
	Response.Write  "<TD VALIGN=""TOP"" ROWSPAN=""3""><IMG SRC=""Images/PromptSummary.gif"" WIDTH=""14"" HEIGHT=""22"" ALT="""" BORDER=""0"" /></TD>"
	Response.Write  "<TD BGCOLOR=""#DDDDBB"" WIDTH=""100%"" ALIGN=""LEFT""><A NAME=""1"" /><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """>"
	Response.Write	asDescriptors(1066) & "</FONT></TD><TD BGCOLOR=""#DDDDBB"" ALIGN=""RIGHT"" VALIGN=""TOP""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD></TR>"
	Response.Write  "<TR><TD BGCOLOR=""#FFFFFF"" ALIGN=""LEFT"" COLSPAN=""2""><IMG SRC=""Images/1ptrans.gif"" WIDTH=""5"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD></TR>"
	Response.Write  "<TR><TD BGCOLOR=""#000000"" ALIGN=""LEFT"" COLSPAN=""2"">"
	Response.Write	"	<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR><TD>&nbsp;</TD><TD>"
	Response.Write	"		<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"

	For lOrder = 1 to aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT)
		Call GetPinbyOrder(aPromptGeneralInfo, aPromptInfo, lOrder, lPin)
		Set oSinglePromptTempXML = aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML).selectSingleNode("/mi/pif[@pin='" & CStr(lPin) &"']")
		Call CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)
		If StrComp(sDisplayUnknownDef, "1", vbBinaryCompare) = 0 Then
			Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)
			sTitle = oSinglePrompt.Title
			If oSinglePrompt.Required Then
				Response.Write "<B> Prompt " & CStr(lOrder) & ": " & sTitle & " </B><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#CC0000"">(" & asDescriptors(661) & ")</FONT><BR />"	'Descriptor: Required
			Else
				Response.Write "<B> Prompt " & CStr(lOrder) & ": " & sTitle & " </B><BR /> "
			End If
			lErrNumber = GetDefaultMeaningForSinglePrompt(aPromptGeneralInfo, aPromptInfo, lPin, sDefault)
			Response.Write sDefault & " <BR /><BR />"
		Else
			lErrNumber = GetPromptSummaryForSinglePrompt(lOrder, aPromptGeneralInfo, lPin, sSinglePromptSummary)
			Response.Write sSinglePromptSummary
			sSinglePromptSummary = ""
		End If
	Next

	Response.Write	"</FONT>"
	Response.Write	"	</TD><TD>&nbsp;</TD></TR></TABLE>"
	Response.Write	"</TD></TR></TABLE>"
End Function

Function DisplaySinglePromptDefaultMeaning(aPromptGeneralInfo, aPromptInfo, lPin)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
    On Error Resume Next
    Dim oSinglePrompt
    Dim sDefault
    Dim lErrNumber
    Dim oSinglePromptTempXML
    Dim sDisplayUnknownDef

    Set oSinglePromptTempXML = aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER)

    Call CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)
    If StrComp(sDisplayUnknownDef, "1", vbBinaryCompare) = 0 Then
		Response.Write	"<TABLE WIDTH=""98%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR><TD><IMG SRC=""Images/1ptrans.gif"" WIDTH=""5"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD><TD BGCOLOR=""#000000"">"
		Response.Write	"	<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR><TD>&nbsp;</TD><TD>"
		Response.Write	"		<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"

		If aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).item(lPin).hasAnswer Or _
		   aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).item(lPin).hasPreviousAnswer Then
			Response.Write	asDescriptors(1069) & " <BR />" 'Your selection:
		ElseIf aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).item(lPin).hasDefaultAnswer Then
			Response.Write	asDescriptors(1071) & " <BR />" 'The default selection is:
		End If

		lErrNumber = GetDefaultMeaningForSinglePrompt(aPromptGeneralInfo, aPromptInfo, lPin, sDefault)
		Response.Write sDefault

		Response.Write	"</FONT>"
		Response.Write	"	</TD><TD>&nbsp;</TD></TR></TABLE>"
		Response.Write	"</TD></TR></TABLE><BR />"
	End If

	set oSinglePrompt = nothing

	DisplaySinglePromptDefaultMeaning = Err.number
	Err.Clear
End Function

Function GetDefaultMeaningForSinglePrompt(aPromptGeneralInfo, aPromptInfo, lPin, sDefault)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
	On Error Resume Next
	Dim oSinglePrompt
	Dim sText
	Dim lErrNumber
	Dim oDefault
	Dim oSinglePromptDisplayXML
	Dim sDisplayXML
	Dim oElementSourceObject

	set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)

	Call GetXMLDOM(aConnectionInfo, oSinglePromptDisplayXML, sErrDescription)

	Set oElementSourceObject = oSinglePrompt.ElementSourceObject
	oElementSourceObject.BlockCount = 0
	oSinglePrompt.DisplayBlockCount = 0
	oSinglePrompt.ExecutionFlags = DssXmlExecutionUseWebCacheOnly Or DssXmlExecutionCheckWebCache

	'oSinglePrompt.setPreviousAsAnswer
	sDisplayXML = oSinglePrompt.DisplayXML

	'In some cases, some prompts won't be fully
    'initialized and hence DisplayXML might be blank
    If Err.number <> NO_ERR Then
		'Check if DisplayXML is blank and restore
		'prompt to original values
		If Len(CStr(sDisplayXML)) = 0 Then
			Err.Clear
			Call oSinglePrompt.Reset()
			'Retrieving original DisplayXML
			sDisplayXML = CStr(oSinglePrompt.DisplayXML)
		End If

		'Keep error value either. Mostly, either
		'there was another error and DisplayXML isn't blank; or
		'2nd call to DisplayXML fails again; or Err.number is cleared.
		lErrNumber = Err.number
    End If

    Call oSinglePromptDisplayXML.loadXML(sDisplayXML)

    Call CreateMeaningForSinglePromptXML(oSinglePromptDisplayXML.selectSingleNode("mi/pif/pa[@ia='1']/exp/nd"), 0, sText)

	If IsEmpty(sText) Then
		sText = "(" & "..." & ")"
	Else
		sText = "(" & sText & ")"
	End If

	aFilterProperties(TABS) = PROMPT_INDENTTAB
	aFilterProperties(CARRIAGE_RETURN) = "<BR />"
	aFilterProperties(CROP_FILTER) = False
	aFilterProperties(MAXIMUM_SIZE) = -1
	aFilterProperties(FOR_EXPORT) = False
	aFilterProperties(FILTER_WIDTH) = 65
	aFilterProperties(FILTER_CONTENTS) = ""
	Call ParseFilterDetails(aConnectionInfo, sText, aFilterProperties, sErrDescription)
	aFilterProperties(FILTER_CONTENTS) = sText
	sDefault = aFilterProperties(FILTER_CONTENTS)

	GetDefaultMeaningForSinglePrompt = Err.number
	Err.Clear
End Function


Function CreateMeaningForSinglePromptXML(oNodes, lLevel, sText)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oNodeList
	Dim lCount
	Dim sOperator
	Dim lExpType

	sOperator = oNodes.selectSingleNode("@disp_n").text
	Set oNodeList = oNodes.selectNodes("./nd")

	For lCount = 0 To oNodeList.length-1
		If (lCount > 0 And lCount < oNodeList.length) Or (lLevel>0 And Len(sText)>0) Then
			sText = sText & " " & sOperator & " "
		End If

		lExpType = CLng(oNodeList.item(lCount).selectSingleNode("@et").text)

		If lExpType = DssXmlFilterBranchQual Then
			Call CreateMeaningForSinglePromptXML(oNodeList.item(lCount), lLevel + 1, sText)
		ElseIf lExpType = DssXmlFilterListQual Then
			Dim oE
			Dim oElemNodes
			Dim sSinglePromptSummary
			Dim lCountAux

			Set oElemNodes = oNodeList.item(lCount).selectNodes("./nd[@et='1' and @nt='2']/oi/es/e")
			lCountAux = 1

			For Each oE In oElemNodes
				If oElemNodes.length = 1 Then
					sSinglePromptSummary = oE.getAttribute("disp_n")
				ElseIf lCountAux = oElemNodes.length Then
					sSinglePromptSummary = sSinglePromptSummary & " " & asDescriptors(1064) & " " &  oE.getAttribute("disp_n") 'Descriptor: or
				Else
					If lCountAux = 1 Then
						sSinglePromptSummary = oE.getAttribute("disp_n")
					Else
						sSinglePromptSummary = sSinglePromptSummary & ", " & oE.getAttribute("disp_n")
					End If
				End If

				lCountAux = lCountAux + 1
			Next

			If Len (sSinglePromptSummary) > 0 Then
				sText = sText & "({" & oNodeList.item(lCount).selectSingleNode("./nd[@et='1' and @nt='5']/@disp_n").text & "} = " & sSinglePromptSummary & ")"
			Else
				sText = oNodeList.item(lCount).selectSingleNode("@disp_n").text & "})"
			End If
		ElseIf lExpType = DssXmlFilterSingleBaseFormExpression Then

			'sText = sText & "({" & oNodeList.item(lCount).selectSingleNode("../..").text & "})"
			sText = sText & "({" & oNodeList.item(lCount).selectSingleNode("@disp_n").text & "})"

			Exit For
		Else
			sText = sText & "({" & oNodeList.item(lCount).selectSingleNode("@disp_n").text & "})"
		End If
	Next

	CreateMeaningForSinglePromptXML = lErrNumber
	Err.Clear
End Function


Function GetPromptSummaryForSinglePrompt(lOrder, aPromptGeneralInfo, lPin, sSinglePromptSummary)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
	On Error Resume Next
	Dim oSingleAnswerPrompt
	Dim oSinglePrompt
	Dim sTitle
	Dim bAnswered
	Dim oDisplayXML
	Dim oElementSourceObject
	Dim sDisplayXML

	set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)
	Call GetXMLDOM(aConnectionInfo, oDisplayXML, sErrDescription)

	'To avoid getting all elements for attributes
	Set oElementSourceObject = oSinglePrompt.ElementSourceObject
	oElementSourceObject.BlockCount = 0
	oSinglePrompt.DisplayBlockCount = 0
	oSinglePrompt.ExecutionFlags = DssXmlExecutionUseWebCacheOnly Or DssXmlExecutionCheckWebCache

	sDisplayXML = oSinglePrompt.DisplayXML
	'sDisplayXML = oSinglePrompt.AnswerXML

    'In some cases, some prompts won't be fully
    'initialized and hence DisplayXML might be blank
    If Err.number <> NO_ERR Then
		'Check if DisplayXML is blank and restore
		'prompt to original values
		If Len(CStr(sDisplayXML)) = 0 Then
			Err.Clear
			Call oSinglePrompt.Reset()
			'Retrieving original DisplayXML
			sDisplayXML = CStr(oSinglePrompt.DisplayXML)
			'sDisplayXML = CStr(oSinglePrompt.AnswerXML)
		End If

		'Keep error value either. Mostly, either
		'there was another error and DisplayXML isn't blank; or
		'2nd call to DisplayXML fails again; or Err.number is cleared.
		lErrNumber = Err.number
    End If

	Call oDisplayXML.loadXML(sDisplayXML)

	'Set oSingleAnswerPrompt = oDisplayXML.selectSingleNode("./mi/in/oi/mi/pif[@pin='"&CStr(lPin)& "']/pa[@ia='1']")
	Set oSingleAnswerPrompt = oDisplayXML.selectSingleNode("./mi/pif[@pin='"&CStr(lPin)& "']/pa[@ia='1']")

	sTitle = oSinglePrompt.Title
	sTitle = "<B> " & replace(asDescriptors(1070), "##", CStr(lOrder)) & ": " & sTitle & " </B> "

	Select Case oSinglePrompt.PromptType
    Case DssXmlPromptLong, DssXmlPromptString, DssXmlPromptDouble, DssXmlPromptDate
	    lErrNumber = GetPromptSummaryForConstantPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)

	Case DssXmlPromptObjects
	    lErrNumber = GetPromptSummaryForObjectPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)

	Case DssXmlPromptElements
	    lErrNumber = GetPromptSummaryForElementPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)

	Case DssXmlPromptExpression
		if oSinglePrompt.ExpressionType = DssXmlFilterAllAttributeQual or oSinglePrompt.ExpressionType = DssXmlExpressionMDXSAPVariable then
		    lErrNumber = GetPromptSummaryForHierachicalPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)
    	else
		    lErrNumber = GetPromptSummaryForExpressionPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)
    	end if

	Case DssXmlPromptDimty
	    lErrNumber = GetPromptSummaryForLevelPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)

	Case Else
	    Call LogErrorXML(aConnectionInfo, Err.Number, Err.Description, Err.source, "PromptDisplayCuLib.asp", "DisplaySinglePromptForJob", "", "Unknown Prompt Type", LogLevelError)
	    lErrNumber = ERR_CUSTOM_UNKNOWN_PROMPT_TYPE
	End Select

	If oSinglePrompt.Required Then
		sTitle = sTitle & "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#CC0000"">(" & asDescriptors(661) & ")</FONT>"
	End If

	If bAnswered Then
		sSinglePromptSummary = sTitle & "<BR />" & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & _
							   "<tr valign=""top""><td width=""15%"">" & "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#000000"">" & _
							   asDescriptors(1069) & "</FONT></td><td>" & "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#000000"">" & _
							   sSinglePromptSummary & "</FONT></td></tr></table>"	'Descriptor: Your selection:
	Else
		sSinglePromptSummary = sTitle & "<BR />" & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & _
							   "<tr valign=""top""><td>" & "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#000000"">" & _
							   sSinglePromptSummary & "</FONT></td></tr></table>"
	End If

	sSinglePromptSummary = sSinglePromptSummary & " <BR />"
	GetPromptSummaryForSinglePrompt = Err.number
	Err.Clear
End Function

Function GetPromptSummaryForConstantPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim sAnswer

	sAnswer = Trim(CStr(oSingleAnswerPrompt.text))
	If Len(sAnswer) = 0 Then
		sSinglePromptSummary = sSinglePromptSummary & asDescriptors(1068) 'Descriptor: (Prompt not answered)
		bAnswered = false
	Else
		sSinglePromptSummary = sAnswer
		bAnswered = true
	End If
	sSinglePromptSummary = sSinglePromptSummary & " <BR />"

	GetPromptSummaryForConstantPrompt = Err.number
	Err.Clear
End Function

Function GetPromptSummaryForObjectPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oO
	Dim sTitle

	Set oO = oSingleAnswerPrompt.selectSingleNode("./oi")
	If oO Is Nothing Then
		sSinglePromptSummary = sSinglePromptSummary & asDescriptors(1068) 'Descriptor: Prompt not answered
		bAnswered = false
	Else
		For each oO in oSingleAnswerPrompt.selectNodes("./oi")
			sSinglePromptSummary = sSinglePromptSummary & oO.getAttribute("disp_n") & ", "
		Next
		sSinglePromptSummary = Left(sSinglePromptSummary, Len(sSinglePromptSummary)-2)
		bAnswered = true
	End If
	sSinglePromptSummary = sSinglePromptSummary & " <BR />"

	GetPromptSummaryForObjectPrompt = Err.number
	Err.Clear
End Function

Function GetPromptSummaryForElementPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
	On Error Resume Next
	Dim sAttribute
	Dim lErrNumber
	Dim oElementObject
	Dim oE
	Dim sTitle

	Set oElementObject = oSinglePrompt.ElementsObject
	sAttribute = oElementObject.AttributeInfo.Name

	Set oE = oSingleAnswerPrompt.selectSingleNode("./oi/es/e")
	If oE Is Nothing Then
		sSinglePromptSummary = sSinglePromptSummary & asDescriptors(1068) 'Descriptor: (Prompt not answered)
		bAnswered = False
	Else
		sSinglePromptSummary = sSinglePromptSummary & sAttribute & " = "
		For each oE in oSingleAnswerPrompt.selectNodes("./oi/es/e")
			sSinglePromptSummary = sSinglePromptSummary & oE.getAttribute("disp_n") & " " & asDescriptors(1064) & " " 'Descriptor: or
		Next
		sSinglePromptSummary = Left(sSinglePromptSummary, Len(sSinglePromptSummary)-4)
		bAnswered = True
	End If
	sSinglePromptSummary = sSinglePromptSummary & " <BR />"

	GetPromptSummaryForElementPrompt = Err.number
	Err.Clear
End Function

Function GetPromptSummaryForExpressionPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'Author				Date			Description
'Gregorio Parra		02/09/2001		Changed the way the filter operator is obtained. Instead of calling
'									function CO_GetFilterOperator, we get it from oSinglePrompt object
'*******************************************************
	On Error Resume Next
	Dim oExpItem
	Dim sFilterOP
	Dim sExpItemText
	Dim bFirst
	Dim cNdNodes

	'Getting the filter operator from oSinglePrompt instead of CO_GetFilterOperator
	If oSinglePrompt.ExpressionObject.RootNode.Operator = DssXmlFunctionOr Then
		sFilterOP = asDescriptors(1064) 'Descriptor: or
	Else
		sFilterOP = asDescriptors(308) 'Descriptor: and
	End If

	'Set flag for first element to be displayed.
	'Different HTML format is needed
	bFirst = true

	'Getting list of nodes to be formatted
	Set oExpItem = oSingleAnswerPrompt.selectSingleNode("./exp/nd/nd")
	'Append expression if there aren't nodes needed to be processed
	If oExpItem Is Nothing Then
		'Creating output string
		sSinglePromptSummary = sSinglePromptSummary & asDescriptors(1068) & " <BR /> " 'Descriptor: (Prompt not answered)
		bAnswered = False
	Else
		'Getting list of answer nodes
		Set cNdNodes = oSingleAnswerPrompt.selectNodes("./exp/nd/nd")
		'Create expression for each answer
		For each oExpItem in cNdNodes
			'Obtaining already formatted display for answer
			sExpItemText = oExpItem.getAttribute("disp_n")
			'Create first string output
			If bFirst Then
				sSinglePromptSummary = sSinglePromptSummary & sExpItemText & " <BR /> "
				bFirst = false
			Else
				'Append more row for remaining answers
				sSinglePromptSummary = sSinglePromptSummary & sFilterOP & " <BR /> "
				sSinglePromptSummary = sSinglePromptSummary & sExpItemText & " <BR /> "
			End If
		Next
		bAnswered = True
	End If

	'Cleaning up objects
	Set oExpItem = Nothing
	Set cNdNodes = Nothing
	GetPromptSummaryForExpressionPrompt = Err.number
	Err.Clear
End Function


Function GetPromptSummaryForHierachicalPrompt(aConnectionInfo, oSingleAnswerPrompt, oSinglePrompt, sSinglePromptSummary, bAnswered)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'History
'Author				Date			Description
'Gregorio Parra		02/09/2001		Modified code to process new XML format and hence,
'									construct new HTML output needed in prompt summary page.
'*******************************************************
	On Error Resume Next
	Dim oExpItem
	Dim sFilterOP
	Dim sExpItemText
	Dim bFirst
	Dim cNdNodes
	Dim sOpFnt
	Dim oElementsList
	Dim oElem
	Dim sValues
	Dim bRemoveLastSemicolon
	Dim oNdElem
	Dim oNdElements

	'Getting the filter operator from oSinglePrompt instead of CO_GetFilterOperator
	If oSinglePrompt.ExpressionObject.RootNode.Operator = DssXmlFunctionOr Then
		sFilterOP = asDescriptors(1064) 'Descriptor: or
	Else
		sFilterOP = asDescriptors(308) 'Descriptor: and
	End If

	'Set flag for first element to be displayed.
	'Different HTML format is needed
	bFirst = true

	'Getting list of nodes to be formatted
	Set oExpItem = oSingleAnswerPrompt.selectSingleNode("./exp/nd/nd")
	'Append expression if there aren't nodes needed to be processed
	If oExpItem Is Nothing Then
		'Creating output string
		sSinglePromptSummary = sSinglePromptSummary & asDescriptors(1068) & " <BR /> " 'Descriptor: (Prompt not answered)
		bAnswered = False
	Else
		'Getting list of answer nodes
		Set cNdNodes =  oSingleAnswerPrompt.selectNodes("./exp/nd/nd")
		'Create expression for each answer
		For each oExpItem in cNdNodes
			'Obtaining already formatted display for answer
			sExpItemText = oExpItem.getAttribute("disp_n")
			'Getting type of operator for single answer
			sOpFnt = oExpItem.selectSingleNode("./op").getAttribute("fnt")

			'If this format is blank, then construct one formmated output from scratch
			'If operator is IN then construct its HTML format
			If Len(sExpItemText) = 0 Or	StrComp(sOpFnt, "22") = 0 Then
				'Obtaining attribute
				sExpItemText = oExpItem.selectSingleNode("./nd").getAttribute("disp_n")
				'Obtaining list of values
				Set oElementsList = oExpItem.selectNodes("./nd[1]/oi/es/e")
				'Set oElementsList = oExpItem.selectNodes("./nd[@disp_n='']/mi/es/e")
				'Setting format flags
				bRemoveLastSemicolon = False
				sValues = ""
				'Create semicolon-delimited list of values
				For Each oElem in oElementsList
					If Not oElem Is Nothing Then
						sValues =  sValues & oElem.getAttribute("disp_n") & ";"
						bRemoveLastSemicolon = True
					End If
				Next
				'Remove last semicolon if needed
				If bRemoveLastSemicolon Then
					sValues = Left(sValues, Len(sValues)-1)
				End If
				'Put attribute, operator and values all together in a string
				sExpItemText = sExpItemText & " " & asDescriptors(587) & " (" & sValues & ")" 'Descriptor: In
			End If

			'Create first string output
			If bFirst Then
				sSinglePromptSummary = sSinglePromptSummary & sExpItemText & " <BR /> "
				bFirst = false
			Else
				'Append more row for remaining answers
				sSinglePromptSummary = sSinglePromptSummary & sFilterOP & " <BR /> "
				sSinglePromptSummary = sSinglePromptSummary & sExpItemText & " <BR /> "
			End If
		Next
		bAnswered = True
	End If

	'Cleaning up objects
	Set oExpItem = Nothing
	Set oElem = Nothing
	Set cNdNodes = Nothing
	Set oElementsList = Nothing

	GetPromptSummaryForHierachicalPrompt = Err.number
	Err.Clear
End Function

Function GetHighlightString(sFolderDID, sHIDID, sATDID, sHighlight)
'*******************************************************
'Purpose:   Display prompt (deafult) meaning
'Inputs:    aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
	On Error Resume Next

	sHighlight = ""
	If Len(sATDID) > 0 Then
		sHighlight = "<oi did='" & sATDID & "' tp='" & CStr(DssXmlTypeAttribute) & "' />"
	End If

	If Len(sHIDID) > 0 Then
		sHighlight = "<oi did='" & sHIDID & "' tp='" & CStr(DssXmlTypeDimension) & "'>" & sHighlight & "</oi>"
	End If

	If Len(sFolderDID) > 0 Then
		sHighlight = "<oi did='" & sFolderDID & "' tp='" & CStr(DssXmlTypeFolder) & "'>" & sHighlight & "</oi>"
	End If

	sHighlight = "<mi>" & sHighlight & "</mi>"

	GetHighlightString = Err.number
	Err.Clear
End Function

Function DisplayTriggersInPromptPage(aConnectionInfo, aPromptGeneralInfo)
'*******************************************************
'Purpose:   Display Triggers in prompt page
'Inputs:    aConnectionInfo, aPromptGeneralInfo
'Outputs:   Err.Number
'*******************************************************
Dim sImageName
Dim aNCSubscriptionInfo

	On Error Resume Next

	'~Draw the header:

	'We'll show the icon based on the object type:
	If aPromptGeneralInfo(PROMPT_B_ISDOC) Then
		sImageName = "Images/document_big.gif"
	Else
		If StrComp(aPromptGeneralInfo(PROMPT_S_VIEWMODE), "Graph", vbTextCompare) = 0 Then
			sImageName =  "Images/graph_big.gif"
		Else
			sImageName = "Images/report_big.gif"
		End If
	End If

	Response.Write "<TABLE BORDER=""0"" WIDTH=""100%"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
	Response.Write "<TD VALIGN=""TOP"" WIDTH=""65""><IMG SRC=""" & sImageName & """ WIDTH=""60"" HEIGHT=""76"" ALT="""" BORDER=""0"" /></TD>"
	Response.Write "<TD WIDTH=""100%"">"


	'Narrowcast Integration
	If aPromptGeneralInfo(PROMPT_B_USE_NC) Then
		'Receive values from request
		Call NCReceiveSubscriptionRequest(oRequest, aNCSubscriptionInfo)

		'Prompt specific values:
		aNCSubscriptionInfo(NC_SUBS_DISPLAY_FORM) = False

		If aNCSubscriptionInfo(NC_SUBS_ACTION) = NC_ACTION_SENDNOW Then
			Call NCDisplaySendNow(aConnectionInfo, aNCSubscriptionInfo)
			aPromptGeneralInfo(PROMPT_N_TRIGGERS_COUNT) = 1 'There is only one Send Now schedule
		Else
			Call NCDisplayTriggers(aConnectionInfo, aNCSubscriptionInfo)
			aPromptGeneralInfo(PROMPT_N_TRIGGERS_COUNT) = aNCSubscriptionInfo(NC_SUBS_SCHEDULES_COUNT)
		End IF


	Else
		'~Set up the schedule info elements:
		ReDim aScheduleInfo(MAX_ELEMENTS_SCHEDULE)

		'Object Id, type:
		If aPromptGeneralInfo(PROMPT_B_ISDOC) Then
			aScheduleInfo(S_OBJECT_ID_SCHEDULE) = aPromptGeneralInfo(PROMPT_S_DOCUMENTID)
			aScheduleInfo(L_OBJECT_TYPE_SCHEDULE) = DssXmlTypeDocumentDefinition
		Else
			aScheduleInfo(S_OBJECT_ID_SCHEDULE) = aPromptGeneralInfo(PROMPT_S_REPORTID)
			aScheduleInfo(L_OBJECT_TYPE_SCHEDULE) = DssXmlTypeReportDefinition
			aScheduleInfo(N_STATE_ID_SCHEDULE) = 0
		End If

		'Ids:
		aScheduleInfo(S_MESSAGE_ID_SCHEDULE) =  aPromptGeneralInfo(PROMPT_S_MSGID)
		aScheduleInfo(S_SELECTED_TRIGGER_ID_SCHEDULE) = aPromptGeneralInfo(S_TRIGGER_ID_PROMPT)

		'Prompt specific values:
		aScheduleInfo(B_SHOW_SUBMIT_SCHEDULE) = False
		aScheduleInfo(B_DISPLAY_FORM_SCHEDULE) = False
		aScheduleInfo(B_HIDE_INFO_SCHEDULE) = True

		'Display schedule list:
		Call DisplayTriggers(aConnectionInfo, aScheduleInfo)

		aPromptGeneralInfo(PROMPT_N_TRIGGERS_COUNT) = aScheduleInfo(L_NUMBER_OF_TRIGGERS_SCHEDULE)

	End If

	'If Not aPromptGeneralInfo(B_TRIGGERS_ONLY_PROMPT) Then
	'	Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>&nbsp;" & asDescriptors(1082) & "</FONT>" 'Descriptor: With the following personalization:
	'End If

	Response.Write "</TD>"
	Response.Write "</TR></TABLE><BR />"


	DisplayTriggersInPromptPage = Err.number
	Err.Clear
End Function

'====================From DisplayCuLib
Function AddInputsForPrompts(aConnectionInfo, sPin, aPromptInfo, oInputs, oDisplayXML)
'*****************************************************************************************
'Purpose:   insert <inputs> to displayXML
'Input:     aConnectionInfo, sPin, aPromptInfo, oInputs
'Output:    oDisplayXML
'*****************************************************************************************

    On Error Resume Next
	Dim oRoot
	Dim oNoncartName
	Dim oROOTMI
	Dim oOldInputs
	Dim oDomAux

	If not aPromptInfo(Clng(sPin), PROMPTINFO_B_ISCART) Then
		set oNoncartName = oInputs.selectSingleNode("noncartname")
		if not oNoncartName is nothing then
			call oInputs.removeChild(oNoncartName)
			Set oNoncartName = Nothing
		end if
		Set oRoot = oInputs.selectSingleNode("/")
		Set oNoncartName = oRoot.ownerDocument.createElement("noncartname")

		If IsEmpty(oNoncartName) Or (oNoncartName Is Nothing) Then
			Err.Clear
			Set oNoncartName = oRoot.createElement("noncartname")
		End If

		oInputs.appendChild oNoncartName
		oNoncartName.Text = "Available_" & sPin
	End If

	Set oROOTMI = oDisplayXML.selectSingleNode("/mi")
    Set oOldInputs = oROOTMI.selectSingleNode("inputs")
	If not oOldInputs is nothing then
		call oROOTMI.removeChild(oOldInputs)
	end if

	'Call oROOTMI.appendChild(oInputs)
	Set oDomAux = server.createobject("MSXML.DOMDocument")
	Call oDomAux.loadXML(oInputs.xml)
	Call oROOTMI.appendChild(oDomAux.documentElement)

    set oRoot = nothing
    set oOldInputs = nothing
    set oNoncartName = nothing
    set oROOTMI = nothing
    set oDomAux = nothing

    AddInputsForPrompts = Err.Number
    Err.Clear
End Function

Function CO_BuildEmbeddedFilterforHIPrompt(aConnectionInfo, oF, oFilterOI, oFilterXML)
'*****************************************************************
'Purpose:   build a filterXML from embedded filterOI in HI Prompt
'Input:     aConnectionInfo, oF, oFilterOI
'Output:    oFilterXML
'*****************************************************************

    On Error Resume Next
    Dim oRootXML
    Dim oIN
    Dim oMI
    Dim oND
    Dim oEXP

    Set oRootXML = oFilterOI.selectSingleNode("/")
    Set oFilterXML = oRootXML.createElement("f")
    Set oMI = oRootXML.createElement("mi")
    Call oFilterXML.appendChild(oMI)
    Set oIN = oRootXML.createElement("in")
    Call oMI.appendChild(oIN)
    Call oIN.appendChild(oFilterOI.cloneNode(True))

    If Err.Number = 0 Then
        Set oEXP = oRootXML.createElement("exp")
        Call oMI.appendChild(oEXP)
        Call oEXP.setAttribute("nc", "1")
        Set oND = oRootXML.createElement("nd")
        Call oEXP.appendChild(oND)
    Else
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CO_BuildEmbeddedFilterforHIPrompt", "", "Error working with XML", LogLevelError)
    End If

    If Err.Number = 0 Then
        With oND
        Call .setAttribute("et", CStr(DssXmlFilterEmbedQual))
        Call .setAttribute("nt", CStr(DssXmlNodeShortcut))
        Call .setAttribute("dmt", CStr(DssXmlNodeDimtyNone))
        Call .appendChild(oF.cloneNode(True))
        End With
    Else
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CO_BuildEmbeddedFilterforHIPrompt", "", "Error working with XML", LogLevelError)
    End If

    Set oRootXML = Nothing
    Set oMI = Nothing
    Set oIN = Nothing
    Set oND = Nothing
    Set oEXP = Nothing

    CO_BuildEmbeddedFilterforHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_GetSpecialFolderXML(aConnectionInfo, sSpecialFolderName, sToken, oSession, sSpecialFolderID, sSpecialFolderXML)
'**************************************************************************
'Purpose:   get a special folderXML given a special folder name
'Inputs:    aConnectionInfo, sSpecialFolderName, sToken, oSession
'Outputs:	sSpecialFolderID, sSpecialFolderXML
'**************************************************************************

    On Error Resume Next
    Dim oObjServer

    Set oObjServer = Nothing
    sSpecialFolderID = ""
    sSpecialFolderXML = ""

    Set oObjServer = oSession.ObjectServer
    If Not IsObject(oObjServer) Or Err.Number <> 0 Then
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CO_GetSpecialFolderXML", "oSession.ObjectServer", "Error accessing DSS Server's object server", LogLevelError)
    Else
        sSpecialFolderID = oObjServer.GetFolderId(sToken, sSpecialFolderName)
        If Len(sSpecialFolderID) = 0 Or Err.Number <> 0 Then
			sErrDescription = Err.Description
            Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CO_GetSpecialFolderXML", "oSession.ObjectServer", "Error obtaining ID For Special Folder", LogLevelError)
        Else
            sSpecialFolderXML = oObjServer.FindObject(sToken, sSpecialFolderID, DssXmlTypeFolder, OBJECT_BROWSING_FLAG, 0, 1, -1)
            If Err.Number <> 0 Then
				sErrDescription = Err.Description
                Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CO_GetSpecialFolderXML", "oObjServer.FindObject", "Error accessing DSS Server's object server", LogLevelError)
            End If
        End If
    End If

    Set oObjServer = Nothing

    CO_GetSpecialFolderXML = Err.Number
    Err.Clear
End Function

Function CO_SearchHierachy(aConnectionInfo, sFolderDID, oFolder)
'**************************************************************************
'Purpose:   get a normal folderXML given a special folderID
'Inputs:    aConnectionInfo, sFolderDID, sToken, oSession
'Outputs:   oFolder
'**************************************************************************
    On Error Resume Next
	Dim oSearch
	Dim lErrNumber
	Dim sSearchResultsXML
	Dim oSearchResultXML

	lErrNumber = GetSearchHelperObject(aConnectionInfo, oSearch, sErrDescription)

	oSearch.SearchRoot  = sFolderDID
	oSearch.Flags = oSearch.Flags Or DssXmlSearchVisibleOnly' + DssXmlSearchRootRecursive
	Call oSearch.AppendType(DssXmlTypeDimension)
	Call oSearch.Submit()
	oSearch.BlockBegin = 1	'lBlockBegin
	oSearch.BlockCount = 0	'lBlockCount
	oSearch.ObjectFlags = oSearch.ObjectFlags Or DssXmlObjectFindHidden
	sSearchResultsXML = oSearch.GetResults(CBool(False), Application.Value("lExecCycleSleepTime"), (CLng(Server.ScriptTimeout) * 1000))
    Call GetXMLDOM(aConnectionInfo, oSearchResultXML, sErrDescription)
    oSearchResultXML.loadXML (sSearchResultsXML)
    Set oSearchResultXML = oSearchResultXML.selectSingleNode("/mi")
	set oFolder = oSearchResultXML.selectSingleNode("/")

	CO_SearchHierachy = Err.number
	Err.Clear
End Function

Function CO_GetPathForFolder(aConnectionInfo, sFolderID, sToken, oSession, sPathXML)
'**************************************************************************
'Purpose:   get ancestor path xml For a folder
'Inputs:    aConnectionInfo, sFolderID, sToken, oSession
'Outputs:   sPathXML
'**************************************************************************
    On Error Resume Next
    Dim oObjServer

    Set oObjServer = Nothing

    Set oObjServer = oSession.ObjectServer
    If Not IsObject(oObjServer) Or Err.Number <> 0 Then
        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CO_GetPathForFolder", "oSession.ObjectServer", "Error Accesing ObjectServer", LogLevelError)
    Else
        sPathXML = oObjServer.FindObject(sToken, sFolderID, DssXmlTypeFolder, DssXmlObjectAncestors, 0, , -1)
        If Err.Number <> 0 Then
			sErrDescription = Err.Description
           Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.source, "PromptDisplayCuLib.asp", "CO_GetPathForFolder", "oObjServer.FindObject", "Error accessing DSS Server's object server", LogLevelError)
        End If
    End If

    Set oObjServer = Nothing

    CO_GetPathForFolder = Err.Number
    Err.Clear
End Function

Function CreatePromptsHelperObject(aConnectionInfo, oAllPrompts, sErrDescription)
'*******************************************************************************
'Purpose:   Gets the prompts helper object
'Inputs:	aConnectionInfo
'Outputs:	oAllPrompts, lErrNumber, sErrDescription
'*******************************************************************************
    On Error Resume Next
    Dim lErrNumber
    lErrNumber = 0
    Set oAllPrompts = Server.CreateObject("WebAPIHelper.DSSXMLPrompts")
    lErrNumber = Err.number
    If lErrNumber <> NO_ERR Then
		sErrDescription = Err.description
		Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "SearchCoLib.asp", "CreatePromptsHelperObject", "", "Error after calling Server.CreateObject(""WebAPIHelper.DSSXMLSearch"")", LogLevelError)
    Else
		oAllPrompts.SessionID = aConnectionInfo(S_TOKEN_CONNECTION)
    End If

    CreatePromptsHelperObject = lErrNumber
    Err.Clear
End Function

Function DisplayHiddenValues(oRequest, aPromptGeneralInfo)
'*******************************************************************************
'Purpose: To display all the <INPUT> tags of type HIDDEN
'Inputs:  oRequest, aPromptGeneralInfo
'*******************************************************************************
    On Error Resume Next

    Dim sAnswerXML
    Dim sEntireString
    Dim sValue
	Dim sName
	Dim lTotalCount
	Dim lCount
	Dim lIndex
	Dim lLogicalSplit


	If Len(oRequest("frommysubscriptions")) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FromMySubscriptions"" VALUE=""True"" />"
	End If
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MsgID"" VALUE=""" & aPromptGeneralInfo(PROMPT_S_MSGID) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportID"" VALUE=""" & aPromptGeneralInfo(PROMPT_S_REPORTID) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentID"" VALUE=""" & aPromptGeneralInfo(PROMPT_S_DOCUMENTID) & """ />"
	If aPromptGeneralInfo(PROMPT_B_ISDOC) Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""doc"" VALUE=""true"" />"
	End If
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""view"" VALUE=""" & aPromptGeneralInfo(PROMPT_S_VIEWMODE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Page"" VALUE=""" & aPageInfo(N_ALIAS_PAGE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Server"" VALUE=""" & aConnectionInfo(S_SERVER_NAME_CONNECTION) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Project"" VALUE=""" & aConnectionInfo(S_PROJECT_CONNECTION) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Port"" VALUE=""" & aConnectionInfo(N_PORT_CONNECTION) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Uid"" VALUE=""" & aConnectionInfo(S_UID_CONNECTION) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UMode"" VALUE=""" & aConnectionInfo(N_USER_MODE_CONNECTION) & """ />"
	If Len(oRequest("xsl")) > 0 And Len(oRequest("css")) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""xsl"" VALUE=""" & oRequest("xsl") & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""css"" VALUE=""" & oRequest("css") & """ />"
	End If
	If StrComp(CStr(oRequest("reprompt")), "1", vbBinaryCompare) = 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""reprompt"" VALUE=""1"" />"
	End If

	sAnswerXML = cleanXML(aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).ShortAnswerXML)
	If Len(sAnswerXML) > lBlockSize  Then
			sEntireString = sAnswerXML
			lTotalCount = Len(sEntireString)
			lCount = 1
			lIndex = 0
			Do
				sName = "nuXML_split_" & lIndex & "_" & "nuXML_AnswerXML"
				lLogicalSplit = 0
				If (lCount + lBlockSize < lTotalCount) Then
					lLogicalSplit = Instr(lCount + lBlockSize,sEntireString,"><",vbTextCompare)
					lLogicalSplit = lLogicalSplit - lCount
				End If
				If lLogicalSplit = 0 Then
					lLogicalSplit = lTotalCount - lCount
				End If
				sValue = Mid(sEntireString,lCount,lLogicalSplit + 1)
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & sName & """ VALUE=""" & sValue  & """ />"
				lCount = lCount + Len(sValue)
				lIndex = lIndex + 1
			Loop While 	lCount < lTotalCount
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""split"" VALUE=""" & "nuXML_AnswerXML" &"|"& lIndex-1   & """ />"
	Else
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""nuXML_AnswerXML"" VALUE=""" & sAnswerXML  & """ />"
	End If

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""nuXML_TempAnswerXML"" VALUE=""" & cleanXML(aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML).xml) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserSelections"" VALUE="""" />"
	If Len(oRequest("xml")) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""xml"" VALUE=""" & oRequest("xml") & """ />"
	End If
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DisplayPin"" VALUE=""" & aPromptGeneralInfo(PROMPT_S_CURORDER) & """ />"
	If aPromptGeneralInfo(PROMPT_B_SUMMARY) Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""summary"" VALUE=""1"" />"
	End If
	If aPromptGeneralInfo(B_ADD_SUBSCRIPTION_PROMPT) Then
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AddSubscription"" VALUE=""1"" />"
	End If
	If aPromptGeneralInfo(B_EDIT_SUBSCRIPTION_PROMPT) Then
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EditSubscription"" VALUE=""" & aPromptGeneralInfo(B_EDIT_SUBSCRIPTION_PROMPT) & """ />"
	End If
	If aPromptGeneralInfo(B_TRIGGERS_ONLY_PROMPT) Then
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TriggersOnly"" VALUE=""" & aPromptGeneralInfo(B_TRIGGERS_ONLY_PROMPT) & """ />"
	End If
	If Len(CStr(aPromptGeneralInfo(S_TRIGGER_ID_PROMPT))) > 0 Then
		If Not aPromptGeneralInfo(B_DISPLAY_TRIGGER_PROMPT) Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TriggerID"" VALUE=""" & aPromptGeneralInfo(S_TRIGGER_ID_PROMPT) & """ />"
		End If
	End If
	If Len(aPromptGeneralInfo(S_OLD_TRIGGER_ID_PROMPT)) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OldTriggerID"" VALUE=""" & aPromptGeneralInfo(S_OLD_TRIGGER_ID_PROMPT) & """ />"
	End If
	If Len(aPromptGeneralInfo(PROMPT_S_FILTERID)) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FilterID"" VALUE=""" & aPromptGeneralInfo(PROMPT_S_FILTERID) & """ />"
	End If
	If Len(aPromptGeneralInfo(PROMPT_S_TEMPLATEID)) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TemplateID"" VALUE=""" & aPromptGeneralInfo(PROMPT_S_TEMPLATEID) & """ />"
	End If
	If Len(oRequest("ExecMode")) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExecMode"" VALUE=""" & oRequest("ExecMode") & """ />"
	End If
	If Len(oRequest("index")) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""index"" VALUE=""" & oRequest("index") & """ />"
	End If
	If aPromptGeneralInfo(PROMPT_B_REEXECUTED) Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Reexecuted"" VALUE=""1"" />"
	End If
	If aPromptGeneralInfo(PROMPT_B_DISABLE_SAVE) Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DisableSave"" VALUE=""1"" />"
	End If
	If Len(oRequest("OrigFolderID")) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OrigFolderID"" VALUE=""" & oRequest("OrigFolderID") & """ />"
	End If
	If Len(oRequest("OrigView")) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OrigView"" VALUE=""" & oRequest("OrigView") & """ />"
	End if
	If Len(oRequest("FromSaveAsPage")) > 0 Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FromSaveAsPage"" VALUE=""" & oRequest("FromSaveAsPage") & """ />"
		'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""New"" VALUE=""New"" />"
	End if

	'Hydra
    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sToken"" VALUE=""" & aConnectionInfo(S_TOKEN_CONNECTION) & """ />"
    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""subGUID"" VALUE=""" & aPromptGeneralInfo(PROMPT_S_SUBSCRIPTIONGUID) & """ />"
    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""action"" VALUE=""" & CStr(oRequest("action")) & """ />"
    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""QOID"" VALUE=""" & CStr(aPromptGeneralInfo(PROMPT_S_QUESTIONOBJECT_ID)) & """ />"
    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""InformationSourceID"" VALUE=""" & CStr(aPromptGeneralInfo(PROMPT_S_INFORMATIONSOURCE_ID)) & """ />"
    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""src"" VALUE=""" & CStr(oRequest("src")) & """ />"
    If Len(oRequest("customQO")) > 0 Then
        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""customQO"" VALUE=""" & oRequest("customQO") & """ />"
    End If

    DisplayHiddenValues = Err.number
    Err.Clear
End Function


Function DisplayGoToAnchorDHTML(bDisplayHeaderTag)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
    On Error Resume Next

	If bDisplayHeaderTag Then
		Response.Write "<SCRIPT language=""JavaScript"">"
	End If

	Response.Write "function gotoAnchor(sCurPin)"
	Response.Write "{"
	Response.Write "	var isNav, isIE;"

	Response.Write "	if ((navigator != null) && (sCurPin > 1)) "
	Response.Write "	{"
	Response.Write "		if ((navigator.appName!=null) && (navigator.appVersion!=null))"
	Response.Write "		{"
	Response.Write "			isNav = (navigator.appName == ""Netscape"");"
	Response.Write "			isIE  = (navigator.appName.indexOf(""Internet Explorer"") != -1);"
	Response.Write "			if (isNav || (isIE && parseInt(navigator.appVersion)>=4 ))"
	Response.Write "			{"
	Response.Write "				self.location = ""#"" + sCurPin;"
	Response.Write "			}"
	Response.Write "		}"
	Response.Write "	}"
	Response.Write "}"

	If bDisplayHeaderTag Then
		Response.Write "</SCRIPT>"
	End If

	Err.Clear
End Function

Function isSupportedMDXPrompt(oSinglePromptQuestionXML)
'*******************************************************************************
'Purpose: checks if this MDX prompt is supported in portal
'Inputs: Prompt Question XML in DOM object
'Outputs:boolean
'*******************************************************************************
    On Error Resume Next

    Dim oOriginNode

    Set oOriginNode = oSinglePromptQuestionXML.selectSingleNode("or/dm/mi")

    if oOriginNode is nothing then
    	isSupportedMDXPrompt = false
    else
    	isSupportedMDXPrompt = true
    end if

End Function


Function ChangeLevelDisplayXML(aConnectionInfo, aPromptInfo, sPin, oAvailable, oSelected)
'**********************************************************************
'Purpose:   change the displayXML structure to save XSL transform time
'Inputs:    aConnectionInfo, oAvailable, oSelected
'Outputs:   oAvailable, oSelected
'**********************************************************************
    On Error Resume Next
    Dim oObj
	Dim lAvailableCount
	Dim sDID
	Dim oSelectedOBJ

	Set oOBJ = oAvailable.selectSingleNode("oi")
    if oOBJ is nothing then
		Call oAvailable.setAttribute("acc", "0")
	else
		'not be selected
		lAvailableCount = 0
		for each oOBJ in oAvailable.selectNodes("oi")
			sDID = oOBJ.getAttribute("did")
			Set oSelectedOBJ = oSelected.selectSingleNode("oi[@did='" & sDID & "']")
			if oSelectedOBJ is nothing then
				Call oOBJ.setAttribute("selected","0")
				lAvailableCount = lAvailableCount + 1
				If lAvailableCount = 1 Then
    				Call oOBJ.setAttribute("first", "1")
				End If
			Else
				Call oOBJ.setAttribute("selected","1")
			end if
		next
		Call oAvailable.setAttribute("acc", CStr(lAvailableCount))
	End if

    Set oOBJ = nothing
	Set oSelectedOBJ = nothing

    ChangeLevelDisplayXML = Err.number
    Err.Clear
End Function
%>
