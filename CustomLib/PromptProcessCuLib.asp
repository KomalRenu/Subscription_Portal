<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%


Function ProcessAllPromptSelectionsForJob(aConnectionInfo, oSession, aPromptGeneralInfo, aPromptInfo, oRequest)
'******************************************************************************************************************
'Purpose:   process all prompts
'Inputs:    aConnectionInfo, oSession, aPromptGeneralInfo, aPromptInfo, oRequest
'Outputs:   aPromptGeneralInfo
'******************************************************************************************************************
    On Error Resume Next
    Dim oSinglePromptQuestionXML
    Dim oSinglePromptTempXML
    Dim oPromptTempAnswersXML
    Dim sPin
    Dim lPType
    Dim lErrNumber
    Dim sUsed
    Dim sClosed
    Dim asSelections()
    Dim temArray
    Dim i
    Dim lPin
    Dim lCurrentPin
    Dim sReq
    Dim oSinglePrompt
    Dim lOrder
    Dim bCurrent

    ReDim asSelections(aPromptGeneralInfo(PROMPT_L_MAXPIN))

    If lErrNumber = NO_ERR Then
        Call CO_RemoveCurrentPin(aPromptGeneralInfo(PROMPT_O_TEMPANSWERSXML))   'remove previous curPin
        'parse UserSelections
        If Len(CStr(oRequest("userselections"))) > 0 Then
            temArray = Split(CStr(oRequest("userselections")), Chr(29), -1, vbBinaryCompare)
            For i = 0 To (UBound(temArray)) / 2 - 1
                asSelections(temArray(2 * i)) = temArray(2 * i + 1)
            Next
        Else
            For i = 0 To aPromptGeneralInfo(PROMPT_L_MAXPIN)
                asSelections(i) = ""
            Next
        End If
    End If

    If lErrNumber = NO_ERR Then
        bCurrent = False
        For lOrder = 1 To aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT)
            Call GetPinbyOrder(aPromptGeneralInfo, aPromptInfo, lOrder, lPin)
            sPin = CStr(lPin)
            Set oSinglePromptQuestionXML = aPromptInfo(lPin, PROMPTINFO_O_QUESTION)
            Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)
            If oSinglePrompt.Used And Not (oSinglePrompt.Closed) Then
                lPType = oSinglePrompt.PromptType
                Call GetPinbyOrder(aPromptGeneralInfo, aPromptInfo, aPromptGeneralInfo(PROMPT_S_CURORDER), lCurrentPin)
                If (aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) Or sPin = CStr(lCurrentPin)) And IsEmpty(oRequest("summary")) Then
                    Set oSinglePromptTempXML = aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER)
                    If lErrNumber = NO_ERR Then
                        Call CO_SetPromptError(oSinglePromptTempXML, "0") 'clear previous PromptError
                        lErrNumber = ProcessSinglePromptSelectionsForJob(aConnectionInfo, oSession, oSinglePrompt, aPromptInfo, sPin, lPType, oSinglePromptQuestionXML, oRequest, asSelections(CInt(sPin)), oSinglePromptTempXML)
                        If lErrNumber <> 0 Then
                            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), "", Err.Source, "PromptProcessCuLib.asp", "ProcessAllPromptSelectionsForJob", "", "Error in call to ProcessSinglePromptSelectionsForJob", LogLevelTrace)
                        End If
                        If lErrNumber = NO_ERR And aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) Then
                            If Not bCurrent Then
                                Call CO_IsCurrentPin(oSinglePromptTempXML, bCurrent)
                                If bCurrent Then
                                    aPromptGeneralInfo(PROMPT_S_CURORDER) = CStr(lOrder)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Else
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), "", Err.Source, "PromptProcessCuLib.asp", "ProcessAllPromptSelectionsForJob", "", "Error in call to CO_RemoveCurrentPin", LogLevelTrace)
    End If

    lErrNumber = CheckUnsubmited(aConnectionInfo, oRequest, aPromptGeneralInfo, aPromptInfo)
    If lErrNumber <> 0 Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessAllPromptSelectionsForJob", "", "Error in call to CheckUnsubmited", LogLevelTrace)
    End If

    If Len(CStr(oRequest("promptnext"))) > 0 Then
        aPromptGeneralInfo(PROMPT_S_CURORDER) = CStr(CLng(aPromptGeneralInfo(PROMPT_S_CURORDER)) + 1)
        If CLng(aPromptGeneralInfo(PROMPT_S_CURORDER)) > aPromptGeneralInfo(PROMPT_L_MAXPIN) Then
            aPromptGeneralInfo(PROMPT_S_CURORDER) = CStr(aPromptGeneralInfo(PROMPT_L_MAXPIN))
        End If
    ElseIf Len(CStr(oRequest("promptprev"))) > 0 Then
        aPromptGeneralInfo(PROMPT_S_CURORDER) = CStr(CLng(aPromptGeneralInfo(PROMPT_S_CURORDER)) - 1)
        If CLng(aPromptGeneralInfo(PROMPT_S_CURORDER)) < 1 Then
            aPromptGeneralInfo(PROMPT_S_CURORDER) = "1"
        End If
    ElseIf Len(CStr(oRequest("indexprev.x"))) > 0 Or Len(CStr(oRequest("indexprev.y"))) > 0 Then
        aPromptGeneralInfo(PROMPT_S_CURORDER) = CStr(oRequest("indexprevvalue"))
    ElseIf Len(CStr(oRequest("indexnext.x"))) > 0 Or Len(CStr(oRequest("indexnext.y"))) > 0 Then
        aPromptGeneralInfo(PROMPT_S_CURORDER) = CStr(oRequest("indexnextvalue"))
    ElseIf Len(CStr(oRequest("indexfirst.x"))) > 0 Or Len(CStr(oRequest("indexfirst.y"))) > 0 Then
        aPromptGeneralInfo(PROMPT_S_CURORDER) = "1"
    ElseIf Len(CStr(oRequest("indexlast.x"))) > 0 Or Len(CStr(oRequest("indexlast.y"))) > 0 Then
        aPromptGeneralInfo(PROMPT_S_CURORDER) = CStr(Int(CLng(aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) - 1) / 5) * 5 + 1)
    Else
        For Each sReq In oRequest
			If Not IsEmpty(oRequest(sReq)) Then
			    If StrComp(Left(sReq, 11), "promptcurr_", vbTextCompare) = 0 Then
			        aPromptGeneralInfo(PROMPT_S_CURORDER) = Mid(sReq, 12, Len(sReq) - 13)
			        Exit For
			    End If
			End If
        Next
    End If

    Set oSinglePromptTempXML = Nothing

    ProcessAllPromptSelectionsForJob = lErrNumber
    Err.Clear
End Function

Function CheckUnsubmited(aConnectionInfo, oRequest, aPromptGeneralInfo, aPromptInfo)
'*************************************************************************
'Purpose:   check unsubmitted form components, like FilterOP: And/or
'Inputs:    aConnectionInfo, oRequest, aPromptGeneralInfo
'Outputs:   aPromptGeneralInfo
'*************************************************************************
    On Error Resume Next
    Dim oSinglePromptTempXML
    Dim lPin
    Dim sFilterOperator
    Dim sMatchCase
    Dim oSinglePrompt
    Dim oExpression
    Dim lRootOperator
    Dim sDisplayUnknownDef

    For lPin = 1 To aPromptGeneralInfo(PROMPT_L_MAXPIN)
        Set oSinglePromptTempXML = aPromptInfo(lPin, PROMPTINFO_O_TEMPANSWER)
        Call CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)

        If StrComp(sDisplayUnknownDef, "0") = 0 Then
            sFilterOperator = CStr(oRequest("filteroperator_" & CStr(lPin)))
            If Len(sFilterOperator) > 0 Then
                Set oSinglePrompt = aPromptGeneralInfo(PROMPT_O_PROMPTSOBJECT).Item(lPin)
                Set oExpression = oSinglePrompt.ExpressionObject
                lRootOperator = DssXmlFunctionAnd
                If StrComp(sFilterOperator, "AND", vbTextCompare) = 0 Then
                    lRootOperator = DssXmlFunctionAnd
                ElseIf StrComp(sFilterOperator, "OR", vbTextCompare) = 0 Then
                    lRootOperator = DssXmlFunctionOr
                End If
                oExpression.RootNode.Operator = lRootOperator
            End If
        End If

        sMatchCase = CStr(oRequest("case_" & CStr(lPin)))

        If Len(sMatchCase) = 0 Then
			sMatchCase = "0"
        End If

        Call CO_SetMatchCase(oSinglePromptTempXML, sMatchCase)
    Next

    Set oSinglePromptTempXML = Nothing

    CheckUnsubmited = Err.Number
    Err.Clear
End Function

Function ProcessSinglePromptSelectionsForJob(aConnectionInfo, oSession, oSinglePrompt, aPromptInfo, sPin, lPType, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML)
'*************************************************************************************************************************************
'Purpose:   build oSinglePromptTempXML For a single prompt question
'Inputs:    aConnectionInfo, oSession, aPromptInfo, sPin, lPType, oSinglePromptQuestionXML, oRequest, sUserSelections
'Outputs:   oSinglePromptTempXML
'*************************************************************************************************************************************
    On Error Resume Next
    Dim lErrNumber

    Select Case lPType
    Case DssXmlPromptLong, DssXmlPromptString, DssXmlPromptDouble, DssXmlPromptDate
        lErrNumber = ProcessSelectionsForConstantPrompt(aConnectionInfo, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML)

    Case DssXmlPromptObjects
        lErrNumber = ProcessSelectionsForObjectPrompt(aConnectionInfo, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, aPromptGeneralInfo, oSinglePromptTempXML)

    Case DssXmlPromptElements
        lErrNumber = ProcessSelectionsForElementPrompt(aConnectionInfo, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML)

    Case DssXmlPromptExpression
        If oSinglePrompt.ExpressionType = DssXmlFilterAllAttributeQual or oSinglePrompt.ExpressionType = DssXmlExpressionMDXSAPVariable Then
            lErrNumber = ProcessSelectionsforHierachicalPrompt(aConnectionInfo, oSession, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML)
        Else
            lErrNumber = ProcessSelectionsForExpressionPrompt(aConnectionInfo, oSession, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, aPromptGeneralInfo, oSinglePromptTempXML)
        End If

    Case DssXmlPromptDimty
        lErrNumber = ProcessSelectionsForLevelPrompt(aConnectionInfo, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML)

    Case Else
        lErrNumber = ERR_CUSTOM_UNKNOWN_PROMPT_TYPE
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), "Unknown prompt type", Err.Source, "PromptProcessCuLib.asp", "ProcessSinglePromptSelectionsForJob", "", "Error in Call to CO_AddDefaultToSinglePromptTempXML", LogLevelError)
    End Select

    If lErrNumber <> 0 Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSinglePromptSelectionsForJob", "", "Error in call to ProcessSelectionsForPrompt, PType=" & lPType, LogLevelTrace)
    End If

    ProcessSinglePromptSelectionsForJob = lErrNumber
    Err.Clear
End Function

Function ProcessSelectionsForConstantPrompt(aConnectionInfo, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML)
'*******************************************************************************************************************************
'Purpose:   read what the user typed into the textbox, And update the oSinglePromptTempXML
'Inputs:    aConnectionInfo, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'*******************************************************************************************************************************
    On Error Resume Next
    Dim oAvailable
    Dim aAvailable

    'If no form Is submitted, last answer should be kept


	'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)

	aAvailable = SplitRequest(oRequest("available_" & sPin))

    If UBound(aAvailable) > -1 Then

        For Each oAvailable In aAvailable
        	If (oSinglePrompt.PromptType = DssXmlPromptDouble Or oSinglePrompt.PromptType = DssXmlPromptLong) Then
        		oSinglePrompt.Value = CStr(oAvailable)
        		If Len(oAvailable) > 0 And Not IsNumeric(oAvailable) Then
        			Call CO_SetPromptError(oSinglePromptTempXML, ERR_ANSWER_MUST_BE_NUMERIC)
					Exit For
        		End If
        	ElseIf  oSinglePrompt.PromptType = DssXmlPromptDate Then
        		If Len(oAvailable) > 0 Then
        			If IsValidDate(oAvailable) Then
        				oSinglePrompt.Value = CStr(oAvailable)
        			Else
        				'this means it's a dynamic date.  portal don't currently support it
        				'don't set the value, just use the old pref set previously.
        			End If
        		End If
			Else
				oSinglePrompt.Value = CStr(oAvailable)
        	End If
        Next
    Else
		If Not IsEmpty(bUseDynamicDateAsDefaultAnswer) And bUseDynamicDateAsDefaultAnswer Then
		' We need to clear the answers for Date prompts if the user does not answer
			If oSinglePrompt.PromptType = DssXmlPromptDate Then
				oSinglePrompt.Value = ""
			End If
		End If
    End If

    ProcessSelectionsForConstantPrompt = Err.Number
    Err.Clear
End Function

Function ProcessSelectionsForObjectPrompt(aConnectionInfo, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sObjectSelections, aPromptGeneralInfo, oSinglePromptTempXML)
'********************************************************************************************************************************
'Purpose:  read user selections, And update the oSinglePromptTempXML
' Inputs:   aConnectionInfo, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sObjectSelections, oSinglePromptTempXML, aPromptGeneralInfo(PROMPT_B_DHTML)
' Outputs:  oSinglePromptTempXML
'********************************************************************************************************************************
    On Error Resume Next
    Dim lResult
    Dim sFDDid

    If Len(CStr(oRequest("prev_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("prev_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbprev_" & sPin)))
    ElseIf Len(CStr(oRequest("next_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("next_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbnext_" & sPin)))
    ElseIf Len(CStr(oRequest("bbcurr_" & sPin))) > 0 Then
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbcurr_" & sPin)))
    End If

    ' handle UserSelections For DHTML version
    If Err.Number = NO_ERR Then
        If aPromptGeneralInfo(PROMPT_B_DHTML) Then

			If oSinglePrompt.PromptType = DssXmlPromptExpression Then
				Call oSinglePrompt.ExpressionObject.Reset()
			ElseIf oSinglePrompt.PromptType = DssXmlPromptObjects Then
				Call oSinglePrompt.FolderObject.Clear()
			Else
				Call oSinglePrompt.Reset()
			End If

            If Len(CStr(oRequest("userselections"))) > 0 Then
                lResult = BuildObjectAnswerXMLFromUserSelections(aConnectionInfo, sPin, sObjectSelections, oSinglePrompt)
                If lResult <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to BuildElementAnswerFromUserSelections", LogLevelTrace)
                End If
            End If
        End If
    End If

    If aPromptInfo(CLng(sPin), PROMPTINFO_B_ISCART) Then
        If Len(CStr(oRequest("add_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("add_" & sPin & ".y"))) > 0 Then
            Call CO_SetCurrentPin(oSinglePromptTempXML)
            lResult = AddSelectionsToHelperObjectForObjectPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
            If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForObjectPrompt", "", "Error in call to AddSelectionsToAnswerXMLForObjectPrompt", LogLevelTrace)
            End If
        ElseIf (Len(CStr(oRequest("remove_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("remove_" & sPin & ".y"))) > 0) Then
              Call CO_SetCurrentPin(oSinglePromptTempXML)
              lResult = RemoveSelectionsFromObjectPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
              If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForObjectPrompt", "", "Error in call to RemoveSelectionsFromObjectPrompt", LogLevelTrace)
              End If
        ElseIf Len(CStr(oRequest("find_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("find_" & sPin & ".y"))) > 0 Then
            Call CO_SetCurrentPin(oSinglePromptTempXML)
            Call CO_SetSearchField(oSinglePromptTempXML, CStr(oRequest("search_" & sPin)))
            Call CO_SetBlockBegin(oSinglePromptTempXML, "1")
        ElseIf Len(CStr(oRequest("hilinkgo_" & sPin))) > 0 Then
            If CStr(oRequest("hilink_" & sPin)) = "-none-" Then
            'Nothing
            Else
                sFDDid = CStr(oRequest("hilink_" & sPin))
                Call CO_SetCurrentPin(oSinglePromptTempXML)
	            Call CO_SetBlockBegin(oSinglePromptTempXML, "1")
                Call CO_SetSearchField(oSinglePromptTempXML, "")
                Call CO_SetHiLinkforObjectPrompt(oSinglePromptTempXML, sFDDid)
            End If
        ElseIf Len(CStr(oRequest("hiparentgo_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("hiparentgo_" & sPin & ".y"))) > 0 Then
            Call CO_SetBlockBegin(oSinglePromptTempXML, "1")
            sFDDid = CStr(oRequest("hiparent_" & sPin))
            Call CO_SetCurrentPin(oSinglePromptTempXML)
            Call CO_SetHiLinkforObjectPrompt(oSinglePromptTempXML, sFDDid)
        End If
    Else    'other styles are of same code

        If oSinglePrompt.PromptType = DssXmlPromptExpression Then
			Call oSinglePrompt.ExpressionObject.Reset()
		ElseIf oSinglePrompt.PromptType = DssXmlPromptObjects Then
			Call oSinglePrompt.FolderObject.Clear()
		Else
			Call oSinglePrompt.Reset()
		End If

        If lResult <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForObjectPrompt", "", "Error in call to ClearAllSelectionsForObjectPrompt", LogLevelTrace)
        End If

        If lResult = NO_ERR Then
            lResult = AddSelectionsToHelperObjectForObjectPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
            If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForObjectPrompt", "", "Error in call to AddSelectionsToAnswerXMLForObjectPrompt", LogLevelTrace)
            End If
        End If
    End If

    ProcessSelectionsForObjectPrompt = lResult
    Err.Clear
End Function

Function ProcessSelectionsForLevelPrompt(aConnectionInfo, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML)
'********************************************************************************************************************************
'Purpose:   read user selections, And update the oSinglePromptTempXML
'Inputs:    aConnectionInfo, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sLevelSelections, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'********************************************************************************************************************************
    On Error Resume Next
    Dim lResult
    Dim sFDDid

    If Len(CStr(oRequest("prev_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("prev_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbprev_" & sPin)))
    ElseIf Len(CStr(oRequest("next_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("next_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbnext_" & sPin)))
    ElseIf Len(CStr(oRequest("bbcurr_" & sPin))) > 0 Then
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbcurr_" & sPin)))
    End If

    If aPromptInfo(CLng(sPin), PROMPTINFO_B_ISCART) Then
        If Len(CStr(oRequest("add_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("add_" & sPin & ".y"))) > 0 Then
            Call CO_SetCurrentPin(oSinglePromptTempXML)
            lResult = AddSelectionsToLevelPrompt(aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML)
            If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForLevelPrompt", "", "Error in call to AddSelectionsToLevelPrompt", LogLevelTrace)
            End If
        ElseIf Len(CStr(oRequest("remove_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("remove_" & sPin & ".y"))) > 0 Then
            Call CO_SetCurrentPin(oSinglePromptTempXML)
            lResult = RemoveSelectionsFromLevelPrompt(aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML)
            If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForLevelPrompt", "", "Error in call to RemoveSelectionsFromLevelPrompt", LogLevelTrace)
            End If
        End If
    Else        'other styles are of same code
        lResult = ClearAllSelections(aConnectionInfo, oSinglePromptTempXML)
        If lResult <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForLevelPrompt", "", "Error in call to ClearAllSelectionsForLevelPrompt", LogLevelTrace)
        End If

        If lResult = NO_ERR Then
            lResult = AddSelectionsToLevelPrompt(aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML)
            If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForLevelPrompt", "", "Error in call to AddSelectionsToLevelPrompt", LogLevelTrace)
            End If
        End If
    End If

    ProcessSelectionsForLevelPrompt = lResult
    Err.Clear
End Function

Function AddSelectionsToLevelPrompt(aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given level prompt, add them to oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim oRootXML
    Dim sSelection
    Dim temArray
    Dim sObjectDID
    Dim sObjType
    Dim sName
    Dim oExistOI
    Dim oNewOI
    Dim aAvailable

    Set oRootXML = oSinglePromptTempXML.selectSingleNode("/")
    'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)

    aAvailable = SplitRequest(oRequest("available_" & sPin))

    If UBound(aAvailable) > -1 Then

        For Each sSelection In aAvailable
            If Len(sSelection) > 0 And StrComp(sSelection, "-none-", vbTextCompare) <> 0 Then
                temArray = Split(CStr(sSelection), Chr(30), -1, vbBinaryCompare)
                sObjectDID = CStr(temArray(0))
                sObjType = CLng(temArray(1))
                sName = CStr(temArray(2))

                Set oExistOI = oSinglePromptTempXML.selectSingleNode("./oi[@did='"&sObjectDID&"']")
                If oExistOI Is Nothing Then
                    Set oNewOI = oRootXML.createElement("oi")
                    Call oNewOI.setAttribute("did", sObjectDID)
                    Call oNewOI.setAttribute("tp", sObjType)
                    Call oNewOI.setAttribute("n", sName)

                    Call oSinglePromptTempXML.appendChild(oNewOI)
                End If
            End If
        Next
    End If

    Set oRootXML = Nothing
    Set oExistOI = Nothing
    Set oNewOI = Nothing

    AddSelectionsToLevelPrompt = Err.Number
    Err.Clear
End Function

Function RemoveSelectionsFromLevelPrompt(aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given level prompt, remove them from oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim temArray
    Dim sSelection
    Dim sObjectDID
    Dim oOI
    Dim aSelected

    'aSelected = Split(oRequest("selected_" & sPin), ", ", -1, vbBinaryCompare)

    aSelected = SplitRequest(oRequest("selected_" & sPin))
    If UBound(aSelected) > -1 Then

        For Each sSelection In aSelected
            If StrComp(sSelection, "-none-", vbTextCompare) = 0 Then
                Exit For
            End If

            temArray = Split(CStr(sSelection), Chr(30), -1, vbBinaryCompare)
            sObjectDID = CStr(temArray(0))

            Set oOI = oSinglePromptTempXML.selectSingleNode("./oi[@did='"&sObjectDID&"']" )     'find selected <o> node
            If Not (oOI Is Nothing) Then
                Call oSinglePromptTempXML.RemoveChild(oOI)
            End If
        Next
        lErrNumber = Err.Number
    End If

    Set oOI = Nothing
    RemoveSelectionsFromLevelPrompt = Err.Number
    Err.Clear
End Function

Function ProcessSelectionsForElementPrompt(aConnectionInfo, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sElementSelections, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   read user selections, And update the oSinglePromptTempXML
'Inputs:    aConnectionInfo, sPin, aPromptInfo, oSinglePromptQuestionXML, oRequest, sElementSelections, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim lResult

    If Len(CStr(oRequest("prev_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("prev_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbprev_" & sPin)))
    ElseIf Len(CStr(oRequest("next_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("next_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbnext_" & sPin)))
    ElseIf Len(CStr(oRequest("bbcurr_" & sPin))) > 0 Then
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbcurr_" & sPin)))
    End If

    lResult = Err.Number
    If lResult = NO_ERR Then
        If aPromptGeneralInfo(PROMPT_B_DHTML) Then
			If oSinglePrompt.PromptType = DssXmlPromptElements Then
				Call oSinglePrompt.ElementsObject.Clear()
			Else
				Call oSinglePrompt.Reset
			End If

            If Len(CStr(oRequest("userselections"))) > 0 Then
                lResult = BuildElementAnswerXMLFromUserSelections(aConnectionInfo, sElementSelections, oSinglePrompt)
                If lResult <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to BuildElementAnswerFromUserSelections", LogLevelTrace)
                    'Err.Raise lResult
                End If
            End If
        End If
    End If

    If lResult = NO_ERR Then
        If aPromptInfo(CLng(sPin), PROMPTINFO_B_ISCART) Then
            If Len(CStr(oRequest("add_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("add_" & sPin & ".y"))) > 0 Then
                Call CO_SetCurrentPin(oSinglePromptTempXML)
                lResult = AddSelectionsToHelperObjectForElementPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
                If lResult <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to AddSelectionsToAnswerXMLForElementPrompt", LogLevelTrace)
                    'Err.Raise lResult
                End If
            ElseIf Len(CStr(oRequest("remove_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("remove_" & sPin & ".y"))) > 0 Then
                Call CO_SetCurrentPin(oSinglePromptTempXML)
                lResult = RemoveSelectionsFromElementPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
                If lResult <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to RemoveSelectionsFromElementPrompt", LogLevelTrace)
                    'Err.Raise lResult
                End If
            ElseIf Len(CStr(oRequest("find_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("find_" & sPin & ".y"))) > 0 Then
                Call CO_SetCurrentPin(oSinglePromptTempXML)
                Call CO_SetSearchField(oSinglePromptTempXML, CStr(oRequest("search_" & sPin)))
                Call CO_SetBlockBegin(oSinglePromptTempXML, "1")
            End If
        Else    'other styles are of same code
			If Not aPromptGeneralInfo(PROMPT_B_DHTML) Then
				If oSinglePrompt.PromptType = DssXmlPromptElements Then
					Call oSinglePrompt.ElementsObject.Clear()
				Else
					Call oSinglePrompt.Reset
				End If
			End If

            If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to ClearAllSelectionsForElementPrompt", LogLevelTrace)
            End If
            If lResult = NO_ERR Then
                lResult = AddSelectionsToHelperObjectForElementPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
                If lResult <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to AddSelectionsToAnswerXMLForElementPrompt", LogLevelTrace)
                End If
            End If
        End If
    Else
        Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in Call to CO_SetBlockBegin", LogLevelTrace)
    End If

    ProcessSelectionsForElementPrompt = lResult
    Err.Clear
End Function

Function AddSelectionsToHelperObjectForElementPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given element prompt, add them to oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim sSelection
    Dim lErrNumber
    Dim oElements
    Dim sEI
    Dim sName
    Dim temArray
    Dim aAvailable

    'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)

    aAvailable = SplitRequest(oRequest("available_" & sPin))
    If UBound(aAvailable) > -1 Then  'At least there is one
        Set oElements = oSinglePrompt.ElementsObject

        For Each sSelection In aAvailable
            If Len(sSelection) > 0 And StrComp(sSelection, "-none-", vbTextCompare) <> 0 Then
                lErrNumber = AddSingleSelectionToHelperObjectForElementPrompt(aConnectionInfo, sSelection, oElements)
                If lErrNumber <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to ClearAllSelectionsForElementPrompt", LogLevelTrace)
                End If
            End If
        Next
    End If

    AddSelectionsToHelperObjectForElementPrompt = lErrNumber
    Err.Clear
End Function

Function AddSingleSelectionToHelperObjectForElementPrompt(aConnectionInfo, sSelection, oElements)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given element prompt, add them to oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim sEI
    Dim sName
    Dim temArray

    temArray = Split(CStr(sSelection), Chr(30), -1, vbBinaryCompare)
    sEI = CStr(temArray(0))
    sName = CStr(temArray(1))
    Call oElements.Add(sEI, 1, sName)

    AddSingleSelectionToHelperObjectForElementPrompt = Err.Number
    Err.Clear
End Function

Function AddSelectionsToHelperObjectForObjectPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given element prompt, add them to oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim sSelection
    Dim lErrNumber
    Dim oFolder
    Dim sEI
    Dim sName
    Dim temArray
    Dim sBefore
    Dim aAvailable

    If IsEmpty(oRequest("selected_" & sPin))Then
        sBefore = "-none-"
    ElseIf Len(oRequest("selected_" & sPin)) = 0 Then
        sBefore = "-none-"
    Else
        'insert just under the item selected in right-side box
        If InStr(1, oRequest("selected_" & sPin), ",", vbBinaryCompare) > 0 Then
			sBefore = Left(oRequest("selected_" & sPin), InStr(1, oRequest("selected_" & sPin), ",", vbBinaryCompare) - 1)  'get first highlighted item
		Else
			sBefore	= oRequest("selected_" & sPin)
		End If
    End If
    'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)
    aAvailable = SplitRequest(oRequest("available_" & sPin))
    If UBound(aAvailable) > -1 Then
        Set oFolder = oSinglePrompt.FolderObject

        For Each sSelection In aAvailable
            If Len(sSelection) > 0 And StrComp(sSelection, "-none-", vbTextCompare) <> 0 Then
                lErrNumber = AddSingleSelectionToHelperObjectForObjectPrompt(aConnectionInfo, sSelection, sBefore, oFolder)
                If lErrNumber <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to ClearAllSelectionsForElementPrompt", LogLevelTrace)
                End If
            End If
        Next
    End If

    AddSelectionsToHelperObjectForObjectPrompt = lErrNumber
    Err.Clear
End Function

Function AddSingleSelectionToHelperObjectForObjectPrompt(aConnectionInfo, sSelection, sBefore, oFolder)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given element prompt, add them to oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim sObjectDID
    Dim sName
    Dim temArray
    Dim lObjType
    Dim lBeforeID

    temArray = Split(CStr(sSelection), Chr(30), -1, vbBinaryCompare)
    sObjectDID = CStr(temArray(0))
    lObjType = CLng(temArray(1))

    sName = CStr(temArray(2))
    If InStr(sName, "&#8364;") > 0 Then
		sName = DecodeStr(sName)
	End If

    If StrComp(sBefore, "-none-", vbTextCompare) <> 0 Then
        temArray = Null
        temArray = Split(CStr(sBefore), Chr(30), -1, vbBinaryCompare)
        lBeforeID = CLng(temArray(3))
    Else
        lBeforeID = -1
    End If
    Call oFolder.AddCopy(sObjectDID, lObjType, sName, lBeforeID)

    AddSingleSelectionToHelperObjectForObjectPrompt = Err.Number
    Err.Clear
End Function

Function RemoveSelectionsFromElementPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given element prompt, remove them from oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim temArray
    Dim sSelection
    Dim lID
    Dim oElements
    Dim lErrNumber
    Dim aSelected
    'aSelected = Split(oRequest("selected_" & sPin), ", ", -1, vbBinaryCompare)
    aSelected = SplitRequest(oRequest("selected_" & sPin))
    If UBound(aSelected) > -1 Then

        For Each sSelection In aSelected
            If StrComp(sSelection, "-none-", vbTextCompare) = 0 Then
                Exit For
            End If

            temArray = Split(CStr(sSelection), Chr(30), -1, vbBinaryCompare)
            lID = CLng(temArray(2))
            Set oElements = oSinglePrompt.ElementsObject
            Call oElements.Remove(lID)
        Next
        lErrNumber = Err.Number
    End If

    RemoveSelectionsFromElementPrompt = lErrNumber
    Err.Clear
End Function

Function RemoveSelectionsFromObjectPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given element prompt, remove them from oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim temArray
    Dim sSelection
    Dim lID
    Dim oFolder
    Dim lErrNumber
    Dim aSelected

    'aSelected = Split(oRequest("selected_" & sPin), ", ", -1, vbBinaryCompare)

    aSelected = SplitRequest(oRequest("selected_" & sPin))
    If UBound(aSelected) > -1 Then

        For Each sSelection In aSelected
			Set	oFolder = oSinglePrompt.FolderObject

            If StrComp(sSelection, "-none-", vbTextCompare) <> 0 Then
				temArray = Split(CStr(sSelection), Chr(30), -1, vbBinaryCompare)
				lID = CLng(temArray(3))
				Call oFolder.Remove(lID)
			Else
				Call oFolder.clear()
            End If
        Next
        lErrNumber = Err.Number
    End If

    RemoveSelectionsFromObjectPrompt = lErrNumber
    Err.Clear
End Function

Function ClearAllSelections(aConnectionInfo, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   clear all selections in oSinglePromptTempXML of element prompt
'Inputs:    aConnectionInfo, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oParent
    Dim oPIF
    Dim lErrNumber

    Set oPIF = oSinglePromptTempXML.cloneNode(False)
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Call oPIF.appendChild(oTemp.cloneNode(True))
    End If
    Set oParent = oSinglePromptTempXML.parentNode
    Call oParent.RemoveChild(oSinglePromptTempXML)
    Call oParent.appendChild(oPIF)

    Set oSinglePromptTempXML = oPIF
    Set oTemp = Nothing
    Set oParent = Nothing

    lErrNumber = Err.Number

    ClearAllSelections = lErrNumber
    Err.Clear
End Function

Function ProcessSelectionsForExpressionPrompt(aConnectionInfo, oSession, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, aPromptGeneralInfo, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   read user selections, And update the oSinglePromptTempXML for expression prompt
'Inputs:    aConnectionInfo, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, aPromptGeneralInfo(PROMPT_B_DHTML), oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim lResult
    Dim lExpType
    Dim sFilterOperator
    Dim oExpression
    Dim lRootOperator
    Dim sDisplayUnknownDef
    Dim bUnknownDef

    lExpType = oSinglePrompt.ExpressionType
    Call CO_ClearRemove(oSinglePromptTempXML)

    If aPromptGeneralInfo(PROMPT_B_DHTML) And aPromptInfo(CLng(sPin), PROMPTINFO_B_ISCART) And lExpType <> DssXmlFilterSingleMetricQual Then
        If StrComp(aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE), "PromptExpression_textbox.xsl", vbTextCompare) <> 0 Then
            Call oSinglePrompt.Reset
            If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to ClearAllSelectionsForElementPrompt", LogLevelTrace)
            End If
            If Len(sUserSelections) > 0 And InStr(1, sUserSelections, "-default-", vbTextCompare) = 0 Then
                lResult = BuildExpressionAnswerFromUserSelections(aConnectionInfo, oRequest, sPin, lExpType, sUserSelections, aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE), oSinglePromptQuestionXML, oSinglePromptTempXML, oSinglePrompt)
                If lResult <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to BuildExpressionAnswerFromUserSelections", LogLevelTrace)
                End If
                Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "0")
            ElseIf InStr(1, sUserSelections, "-default-", vbTextCompare) > 0 Then
                Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "1")
                Call oSinglePrompt.SetDefaultAsAnswer
            End If
        End If
    End If

    If lResult = 0 Then
        Select Case lExpType
        Case DssXmlFilterSingleMetricQual, DssXmlFilterAttributeIDQual, DssXmlFilterAttributeDESCQual
            If Len(CStr(oRequest("prev_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("prev_" & sPin & ".y"))) > 0 Then
                Call CO_SetCurrentPin(oSinglePromptTempXML)
                Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbprev_" & sPin)))
            ElseIf Len(CStr(oRequest("next_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("next_" & sPin & ".y"))) > 0 Then
                Call CO_SetCurrentPin(oSinglePromptTempXML)
                Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbnext_" & sPin)))
            ElseIf Len(CStr(oRequest("bbcurr_" & sPin))) > 0 Then
                Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbcurr_" & sPin)))
            End If

            If aPromptInfo(CLng(sPin), PROMPTINFO_B_ISCART) Then
                If Len(CStr(oRequest("add_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("add_" & sPin & ".y"))) > 0 Then
                    Call CO_SetCurrentPin(oSinglePromptTempXML)
                    'Call CO_GetbUnknownDef(oSinglePromptTempXML, bUnknownDef)
                    Call CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)
                    If StrComp(sDisplayUnknownDef, "1") = 0 Then 'bUnknownDef and
                        Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "0")
                        Call oSinglePrompt.Reset
                    End If
                    lResult = AddSelectionsToHelperObjectForExpressionPrompt(aConnectionInfo, aPromptInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
                    If lResult <> 0 Then
                        Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForExpressionPrompt", "", "Error in call to AddSelectionsToExpressionPrompt", LogLevelTrace)
                    End If
                ElseIf Len(CStr(oRequest("remove_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("remove_" & sPin & ".y"))) > 0 Then
                    Call CO_SetCurrentPin(oSinglePromptTempXML)
                    lResult = RemoveSelectionsFromExpressionPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
                    If lResult <> 0 Then
                        Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForExpressionPrompt", "", "Error in call to RemoveSelectionsFromExpressionPrompt", LogLevelTrace)
                    End If
                ElseIf Len(CStr(oRequest("loadfilego_" & sPin))) > 0 Then
                    Call CO_SetCurrentPin(oSinglePromptTempXML)
                    Call CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)
                    If StrComp(sDisplayUnknownDef, "1") = 0 Then 'bUnknownDef and
                        Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "0")
                        Call oSinglePrompt.Reset
                    End If
                    lResult = LoadTextFile(aConnectionInfo, oRequest, sPin, lExpType, aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE), oSinglePrompt, oSinglePromptTempXML)
                    If lResult <> NO_ERR Then
                        Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForExpressionPrompt", "", "Error in call to RemoveSelectionsFromExpressionPrompt", LogLevelTrace)
                    End If
                End If
            Else        'other styles are of same code
                'use default for non-cart style
                Call CO_GetbUnknownDef(oSinglePromptTempXML, bUnknownDef)

				'This way, default value will remain intact when
				'Prompt summary is selected and no values in input textfield are entered
				If bUnknownDef Then
					If Len(oRequest("input_" & sPin)) = 0 or Len(oRequest("available_" & sPin)) = 0 Then
						'If there's no input data provided, make sure the default value is shown
						Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "1")
						'Reset this prompt to its default value.
						Call oSinglePrompt.SetDefaultAsAnswer()
						Exit Function
					End If

					'At this point, input was entered by user. Need then to
					'display current selection. Don't display unknown default value
					Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "0")
				End If

                Call oSinglePrompt.Reset
                If lResult <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForExpressionPrompt", "", "Error in call to CO_ClearAllSelectionsForExpressionPrompt", LogLevelTrace)
                End If

                If lResult = NO_ERR Then
                    lResult = AddSelectionsToHelperObjectForExpressionPrompt(aConnectionInfo, aPromptInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
                    If Len(CStr(oRequest("PromptGo"))) = 0 And Len(CStr(oRequest("SubscribeGo"))) = 0 Then
                        Call CO_SetPromptError(oSinglePromptTempXML, "0")
                        lResult = 0
                    ElseIf lResult <> 0 Then
                        Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForExpressionPrompt", "", "Error in call to AddSelectionsToExpressionPrompt", LogLevelTrace)
                    End If
                End If
            End If
        Case Else
            lResult = ERR_CUSTOM_UNKNOWN_EXPPROMPT_TYPE
            Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForExpressionPrompt", "", "Unknown Prompt Type", LogLevelError)
        End Select
    End If

    ProcessSelectionsForExpressionPrompt = lResult
    Err.Clear
End Function

Function ProcessSelectionsforHierachicalPrompt(aConnectionInfo, oSession, oSinglePrompt, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   Process Hierachical prompt
'Inputs:    aConnectionInfo, oSession, aPromptInfo, sPin, oSinglePromptQuestionXML, oRequest, sUserSelections, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim sCur
    Dim temArray
    Dim sCurATName
    Dim sCurATDID
    Dim sOldATName
    Dim sOldATDid
    Dim sCurDrillDir
    Dim lResult
    Dim sHIDid
    Dim sHierachyXML
    Dim oFilterXML_Drill
    Dim oF
    Dim sFiltered
    Dim lExpType
    Dim sHIFlag
    Dim oElementSource
    Dim sDisplayUnknownDef
    Dim oFilterExp_Drill

    lExpType = oSinglePrompt.ExpressionType
    Call CO_ClearRemove(oSinglePromptTempXML)

    If aPromptGeneralInfo(PROMPT_B_DHTML) And aPromptInfo(CLng(sPin), PROMPTINFO_B_ISCART) Then   'And (Not aPromptGeneralInfo(PROMPT_B_SUMMARY)) 'And aPromptInfo(Clng(sPin), PROMPTINFO_B_ISCART)
        If Len(sUserSelections) = 0 Then
            Call oSinglePrompt.Reset
        ElseIf Len(sUserSelections) > 0 And InStr(1, sUserSelections, "-default-", vbTextCompare) = 0 Then
            Call oSinglePrompt.Reset
            lResult = BuildExpressionAnswerFromUserSelections(aConnectionInfo, oRequest, sPin, lExpType, sUserSelections, aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE), oSinglePromptQuestionXML, oSinglePromptTempXML, oSinglePrompt)
            If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to BuildExpressionAnswerFromUserSelections", LogLevelTrace)
            End If
            Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "0")
        ElseIf InStr(1, sUserSelections, "-default-", vbTextCompare) > 0 Then
            Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "1")
            Call oSinglePrompt.SetDefaultAsAnswer
        End If
    End If

    If Len(CStr(oRequest("attributego_" & sPin))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, "1")
        Call CO_SetSearchField(oSinglePromptTempXML, "")

        sCur = oRequest("attribute_" & sPin)
        temArray = Split(CStr(sCur), Chr(30), -1, vbBinaryCompare)
        sCurATDID = CStr(temArray(0))
        sCurATName = CStr(temArray(1))
        If UBound(temArray) > 2 Then
            sFiltered = CStr(temArray(3))
        Else
            sFiltered = ""
        End If

        Call CO_SetAttributeforHIPrompt(oSinglePromptTempXML, sCurATName, sCurATDID)
        If Len(sFiltered) = 0 Then
            Call CO_ClearFilterXMLForDrillInHIPrompt(oSinglePromptTempXML)
        End If

    ElseIf Len(CStr(oRequest("find_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("find_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetSearchField(oSinglePromptTempXML, CStr(oRequest("search_" & sPin)))
        Call CO_SetBlockBegin(oSinglePromptTempXML, "1")

    ElseIf Len(CStr(oRequest("add_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("add_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)

        Call CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)
        If StrComp(sDisplayUnknownDef, "1") = 0 Then 'bUnknownDef and
            Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "0")
            Call oSinglePrompt.Reset
        End If

        Call CO_GetAttributeforHIPrompt(oSinglePromptTempXML, sCurATName, sCurATDID)
        lResult = AddSelectionsToHelperObjectForHIPrompt(aConnectionInfo, aPromptInfo, oRequest, sPin, sCurATName, sCurATDID, oSinglePrompt, oSinglePromptTempXML)
        If lResult <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsforHierachicalPrompt", "", "Error in call to AddSelectionsToExpressionPrompt", LogLevelTrace)
        End If

    ElseIf Len(CStr(oRequest("remove_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("remove_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        lResult = RemoveSelectionsFromExpressionPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
        If lResult <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsforHierachicalPrompt", "", "Error in call to RemoveSelectionsFromExpressionPrompt", LogLevelTrace)
            'Err.Raise lResult
        End If

    ElseIf Len(CStr(oRequest("drillgo_" & sPin))) > 0 Then
        If CStr(oRequest("drill_" & sPin)) = "-none-" Then
            'Nothing
        ElseIf IsEmpty(oRequest("available_" & sPin))Then
            Call CO_SetPromptError(oSinglePromptTempXML, ERR_SELECTELEM_BEFOREDRILL)
        ElseIf Len(oRequest("available_" & sPin)) = 0 Then
            Call CO_SetPromptError(oSinglePromptTempXML, ERR_SELECTELEM_BEFOREDRILL)
        Else
			If StrComp(oRequest("available_" & sPin), "-none-") <> 0 Then
				Call CO_SetCurrentPin(oSinglePromptTempXML)
				Call CO_SetBlockBegin(oSinglePromptTempXML, "1")
				Call CO_SetSearchField(oSinglePromptTempXML, "")
				Call CO_ClearFilterXMLForDrillInHIPrompt(oSinglePromptTempXML)

				Call CO_GetAttributeforHIPrompt(oSinglePromptTempXML, sOldATName, sOldATDid)
				Call CO_SetAttBeforeDrillforHIPrompt(oSinglePromptTempXML, sOldATName, sOldATDid)

				sCur = CStr(oRequest("drill_" & sPin))
				temArray = Split(CStr(sCur), Chr(30), -1, vbBinaryCompare)
				sCurATName = CStr(temArray(0))
				sCurATDID = CStr(temArray(1))
				sCurDrillDir = CStr(temArray(2))
				Call CO_SetAttributeforHIPrompt(oSinglePromptTempXML, sCurATName, sCurATDID)

                If StrComp(sCurDrillDir, "down") = 0 Then
					lResult = BuildFilterXMLforDrillinHIPrompt(aConnectionInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oFilterExp_Drill)
				Else
					lResult = BuildFilterXMLforDrillupHIPrompt(aConnectionInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, sCurATDID, oFilterExp_Drill, oRequest)
				End If

				If lResult <> 0 Then
					Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsforHierachicalPrompt", "BuildFilterXMLforDrillinHIPrompt", "Error in call to BuildFilterXMLforDrillinHIPrompt", LogLevelTrace)
				Else
					If oFilterExp_Drill Is Nothing Then
						Call CO_SetFilterXMLForDrillInHIPrompt(oSinglePromptTempXML, "")
					Else
						Call CO_SetFilterXMLForDrillInHIPrompt(oSinglePromptTempXML, oFilterExp_Drill.XML)
					End If
				End If
			End If
        End If

    ElseIf Len(CStr(oRequest("list_" & sPin))) > 0 Or Len(CStr(oRequest("list_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("list_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, "1")
        Call CO_ClearAttributeforHIPrompt(oSinglePromptTempXML)
        Call CO_ClearFilterXMLForDrillInHIPrompt(oSinglePromptTempXML)
        Call CO_SetHiFlagForHIPrompt(oSinglePromptTempXML, "ELEM")

    ElseIf Len(CStr(oRequest("qualify_" & sPin))) > 0 Or Len(CStr(oRequest("qualify_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("qualify_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, "1")
        Call CO_ClearAttributeforHIPrompt(oSinglePromptTempXML)
        Call CO_ClearFilterXMLForDrillInHIPrompt(oSinglePromptTempXML)
        Call CO_SetHiFlagForHIPrompt(oSinglePromptTempXML, "QUAL")

    ElseIf Len(CStr(oRequest("higo_" & sPin))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_GetHiFlagForHIPrompt(oSinglePromptTempXML, aPromptInfo(CLng(sPin), PROMPTINFO_B_ISALLDIMENSION), sHIFlag)
        If sHIFlag = "PICK_ELEM" Then
            Call CO_SetHiFlagForHIPrompt(oSinglePromptTempXML, "ELEM")
        ElseIf sHIFlag = "PICK_QUAL" Then
            Call CO_SetHiFlagForHIPrompt(oSinglePromptTempXML, "QUAL")
        End If
        Call CO_SetBlockBegin(oSinglePromptTempXML, "1")
        Call CO_SetSearchField(oSinglePromptTempXML, "")
        Call CO_ClearAttributeforHIPrompt(oSinglePromptTempXML)
        Call CO_ClearFilterXMLForDrillInHIPrompt(oSinglePromptTempXML)
        temArray = Split(CStr(oRequest("hi_" & sPin)), Chr(30), -1, vbBinaryCompare)
        sHIDid = CStr(temArray(1))
        Call CO_SetHierachyDIDForHIPrompt(oSinglePromptTempXML, sHIDid)

    ElseIf Len(CStr(oRequest("sfgo_" & sPin))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_ClearAttributeforHIPrompt(oSinglePromptTempXML)
        Call CO_ClearHierachyDIDForHIPrompt(oSinglePromptTempXML)
        Call CO_ClearHierachyForHIPrompt(oSinglePromptTempXML)
        Call CO_SetSubFolderForPromptAllDimensions(oSinglePromptTempXML, CStr(oRequest("sf_" & sPin)))
        Call CO_GetHiFlagForHIPrompt(oSinglePromptTempXML, aPromptInfo(CLng(sPin), PROMPTINFO_B_ISALLDIMENSION), sHIFlag)
        If sHIFlag = "ELEM" Then
            Call CO_SetHiFlagForHIPrompt(oSinglePromptTempXML, "PICK_ELEM")
        ElseIf sHIFlag = "QUAL" Then
            Call CO_SetHiFlagForHIPrompt(oSinglePromptTempXML, "PICK_QUAL")
        End If

    ElseIf Len(CStr(oRequest("prev_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("prev_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbprev_" & sPin)))

    ElseIf Len(CStr(oRequest("next_" & sPin & ".x"))) > 0 Or Len(CStr(oRequest("next_" & sPin & ".y"))) > 0 Then
        Call CO_SetCurrentPin(oSinglePromptTempXML)
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbnext_" & sPin)))

    ElseIf Len(CStr(oRequest("bbcurr_" & sPin))) > 0 Then
        Call CO_SetBlockBegin(oSinglePromptTempXML, CStr(oRequest("bbcurr_" & sPin)))

    Else
        Call CO_SetBlockBegin(oSinglePromptTempXML, "1")

    End If

    If Err.Number <> 0 Then
        lResult = Err.Number
    End If

    ProcessSelectionsforHierachicalPrompt = lResult
    Err.Clear
End Function

Function AddSelectionsToHelperObjectForExpressionPrompt(aConnectionInfo, aPromptInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
'***************************************************************************************************
'Purpose:   append single answer to oSinglePrompt for Element prompt
'Inputs:    aConnectionInfo, aPromptInfo, oRequest, sPin
'Outputs:   oSinglePrompt
'***************************************************************************************************
    On Error Resume Next
    Dim lExpType

    lExpType = oSinglePrompt.ExpressionType

    Select Case oSinglePrompt.ExpressionType
    Case DssXmlFilterSingleMetricQual
        lErrNumber = AddSelectionsToHelperObjectForMetricQual(aConnectionInfo, oRequest, sPin, lExpType, oSinglePrompt, oSinglePromptTempXML)
    Case DssXmlFilterAllAttributeQual, DssXmlFilterAttributeIDQual, DssXmlFilterAttributeDESCQual
        lErrNumber = AddSelectionsToHelperObjectForAttributeQual(aConnectionInfo, aPromptInfo, oRequest, sPin, AQ_ATTRQUAL, lExpType, oSinglePrompt, oSinglePromptTempXML)
    End Select

    AddSelectionsToHelperObjectForExpressionPrompt = lErrNumber
    Err.Clear
End Function

Function AddSelectionsToHelperObjectForHIPrompt(aConnectionInfo, aPromptInfo, oRequest, sPin, sCurATName, sCurATDID, oSinglePrompt, oSinglePromptTempXML)
'***************************************************************************************************
'Purpose:   append single answer to oSinglePrompt for Element prompt
'Inputs:    aConnectionInfo, aPromptInfo, oRequest, sPin, sCurATName, sCurATDID
'Outputs:   oSinglePrompt
'***************************************************************************************************
    On Error Resume Next
    Dim lExpType
    Dim lErrNumber

    lExpType = oSinglePrompt.ExpressionType
    If IsEmpty(oRequest("operator_" & sPin)) Then
        lErrNumber = AddSelectionsToHelperObjectForElementListinHI(aConnectionInfo, oRequest, sPin, sCurATName, sCurATDID, oSinglePrompt)
    Else
        lErrNumber = AddSelectionsToHelperObjectForAttributeQual(aConnectionInfo, aPromptInfo, oRequest, sPin, AQ_HIPROMPT, lExpType, oSinglePrompt, oSinglePromptTempXML)
    End If
    AddSelectionsToHelperObjectForHIPrompt = lErrNumber
    Err.Clear
End Function

Function AddSelectionsToHelperObjectForElementListinHI(aConnectionInfo, oRequest, sPin, sATName, sATDID, oSinglePrompt)
'***************************************************************************************************
'Purpose:   append single answer to oSinglePrompt for Element prompt
'Inputs:    aConnectionInfo, sEI, oSinglePrompt
'Outputs:   oSinglePrompt
'***************************************************************************************************
    On Error Resume Next
    Dim oExpression
    Dim oOperatorNode
    Dim oElementList
    Dim sSelection
    Dim oElements
    Dim oShortCutNode
    Dim oRootNode
    Dim oExpNode
    Dim lIndex
    Dim bExist
    Dim oNode
    Dim aAvailable

    bExist = False
    Set oExpression = oSinglePrompt.ExpressionObject
    Set oRootNode = oExpression.RootNode
    If oRootNode.HasChildNodes Then
        For lIndex = 1 To oRootNode.ChildCount
            Set oNode = oRootNode.Child(lIndex)
            If oNode.ExpressionType = DssXmlFilterListQual Then
                If oNode.FirstChild.ShortcutID = sATDID Then
                    Set oElements = oNode.Child(2).ElementsObject
                    bExist = True
                End If
            End If
        Next
    End If

    If Not (bExist) Then
        Set oOperatorNode = oExpression.CreateOperatorNode(DssXmlFilterListQual, DssXmlFunctionIn)
        Set oShortCutNode = oExpression.CreateShortCutNode(sATDID, DssXmlTypeAttribute, oOperatorNode)
        Set oElementList = oExpression.CreateElementListNode(sATDID, oOperatorNode)
        Set oElements = oElementList.ElementsObject
        oShortCutNode.DisplayName = sATName
    End If

    lErrNumber = Err.Number

    'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)

    aAvailable = SplitRequest(oRequest("available_" & sPin))

    If UBound(aAvailable) > -1 Then

        For Each sSelection In aAvailable
            If Len(sSelection) > 0 And StrComp(sSelection, "-none-", vbTextCompare) <> 0 Then
                lErrNumber = AddSingleSelectionToHelperObjectForElementPrompt(aConnectionInfo, sSelection, oElements)
                If lErrNumber <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to ClearAllSelectionsForElementPrompt", LogLevelTrace)
                End If
            End If
        Next
    End If

    AddSelectionsToHelperObjectForElementListinHI = lErrNumber
    Err.Clear
End Function

Function AddSelectionsToHelperObjectForElementListinHIFromUserSelection(aConnectionInfo, sPin, sATDID, sATName, asSelectionArray, lBeginIndex, lEndIndex, oSinglePrompt)
'***************************************************************************************************
'Purpose:   append single answer to oSinglePrompt for Element prompt
'Inputs:    aConnectionInfo, sEI, oSinglePrompt
'Outputs:   oSinglePrompt
'***************************************************************************************************
    On Error Resume Next
    Dim oExpression
    Dim oOperatorNode
    Dim oElementList
    Dim sSelection
    Dim oElements
    Dim oShortCutNode
    Dim oRootNode
    Dim oExpNode
    Dim lIndex
    Dim bExist
    Dim oNode

    bExist = False
    Set oExpression = oSinglePrompt.ExpressionObject
    Set oRootNode = oExpression.RootNode
    If oRootNode.HasChildNodes Then
        For lIndex = 1 To oRootNode.ChildCount
            Set oNode = oRootNode.Child(lIndex)
            If oNode.ExpressionType = DssXmlFilterListQual Then
                If oNode.FirstChild.ShortcutID = sATDID Then
                    Set oElements = oNode.Child(2).ElementsObject
                    bExist = True
                End If
            End If
        Next
    End If

    If Not (bExist) Then
        Set oOperatorNode = oExpression.CreateOperatorNode(DssXmlFilterListQual, DssXmlFunctionIn)
        Set oShortCutNode = oExpression.CreateShortCutNode(sATDID, DssXmlTypeAttribute, oOperatorNode)
        Set oElementList = oExpression.CreateElementListNode(sATDID, oOperatorNode)
        Set oElements = oElementList.ElementsObject
        oShortCutNode.DisplayName = sATName
    End If

    lErrNumber = Err.Number
    'If oRequest("available_" & sPin).Count > 0 Then
        For lIndex = lBeginIndex To lEndIndex
            If Len(asSelectionArray(lIndex)) > 0 And StrComp(asSelectionArray(lIndex), "-none-", vbTextCompare) <> 0 Then
                lErrNumber = AddSingleSelectionToHelperObjectForElementPrompt(aConnectionInfo, asSelectionArray(lIndex), oElements)
                If lErrNumber <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "ProcessSelectionsForElementPrompt", "", "Error in call to ClearAllSelectionsForElementPrompt", LogLevelTrace)
                End If
            End If
        Next
    'End If

    AddSelectionsToHelperObjectForElementListinHIFromUserSelection = lErrNumber
    Err.Clear
End Function

Function AddSingleSelectionToHelperObjectForAttributeQual(aConnectionInfo, sATName, sATDID, sFMName, sFMDID, lDDT, sOperator, sInput, lExpType, oSinglePrompt, oSinglePromptTempXML, sXsl)
'***************************************************************************************************
'Purpose:   append single answer to oSinglePrompt for Element prompt
'Inputs:    aConnectionInfo, sEI, oSinglePrompt
'Outputs:   oSinglePrompt
'***************************************************************************************************
    On Error Resume Next
    Dim oExpression
    Dim oOperatorNode
    Dim sOperatorType
    Dim sOperatorValue
    Dim aInput
    Dim lIndex
    Dim sExpressionText
    Dim oFormShortCutNode
    Dim lErrNumber
    Dim lInputArraySize, sOP
    Dim sInputAux
    Dim bRemoveSubExp
    Dim bValidationError

    Set oExpression = oSinglePrompt.ExpressionObject
    sOperatorValue = Mid(sOperator, 2, Len(sOperator) - 1)

    If (Clng(sOperatorValue) = DssXmlFunctionIn) Or (Clng(sOperatorValue) = DssXmlFunctionNotIn) Then
    'If sOperator = OperatorType_Metric & CStr(DssXmlFunctionIn) Then
		If Len(sFMDID)>0 Then
			Set oOperatorNode = oExpression.CreateOperatorNode(DssXmlFilterListFormQual, CInt(sOperatorValue))
		Else
			Set oOperatorNode = oExpression.CreateOperatorNode(DssXmlFilterListQual, CInt(sOperatorValue))
		End If
    Else
        Set oOperatorNode = oExpression.CreateOperatorNode(DssXmlFilterSingleBaseFormQual, CInt(sOperatorValue))
    End If

    If Len(sFMDID)>0 Then
		Set oFormShortCutNode = oExpression.CreateFormShortCutNode(sATDID, sFMDID, oOperatorNode)
	Else
		Set oFormShortCutNode = oExpression.CreateShortCutNode(sATDID, DssXmlTypeAttribute, oOperatorNode)
	End If

    'oFormShortCutNode.DisplayName = sATName & Chr(30) & sFMName
    oFormShortCutNode.DisplayName = sATName & "\" & sFMName

    aInput = Split(sInput, ";", -1, vbBinaryCompare)
    sInputAux = ""
    lInputArraySize = UBound(aInput)
    lErrNumber = AreMultipleAnswersSupportedByOperator(sOperator, lInputArraySize)

    If lErrNumber = -1 Then
		Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOMANY_SELECTIONS_EXPRESSIONPROMPT)
		bRemoveSubExp = True
	Else
		bRemoveSubExp = False
	End If

    If StrComp(sXsl, "PromptExpression_TextFile.xsl", vbTextCompare) = 0 Then
		bValidationError = True
    Else
		If (CLng(sOperatorValue) <> DssXmlFunctionIn) Or (CLng(sOperatorValue) <> DssXmlFunctionNotIn) Then
			bValidationError = True
		Else
			bValidationError = False
		End If
    End If


    If lErrNumber = NO_ERR Then
		For lIndex = 0 To lInputArraySize
			If Not ValidateSegmentDataType(aInput(lIndex), CStr(lDDT)) Then
				If bValidationError Then
					If lInputArraySize > 1 Then
						If lDDT = DssXmlDataTypeDate Then
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_ALL_ANSWERS_MUST_BE_A_DATE)
						Else
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_ALL_ANSWERS_MUST_BE_NUMERIC)
						End If
					Else
						If lDDT = DssXmlDataTypeDate Then
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_ANSWER_MUST_BE_A_DATE)
						Else
							Call CO_SetPromptError(oSinglePromptTempXML, ERR_ANSWER_MUST_BE_NUMERIC)
						End If
					End If
				End If
			Else
				Select Case lDDT
					Case DssXmlDataTypeTimeStamp, DssXmlDataTypeTime, DssXmlDataTypeDate
						Call oExpression.CreateTimeNode(aInput(lIndex), lDDT, oOperatorNode)
					Case Else
						Call oExpression.CreateConstantNode(aInput(lIndex), lDDT, oOperatorNode)
				End	Select

				If Len(sInputAux) = 0 Then
					sInputAux = aInput(lIndex)
				Else
					sInputAux = sInputAux & ", " & aInput(lIndex)
				End If
			End If
		Next

		If Len(sInputAux) = 0 Then
			bRemoveSubExp = True
		End If
	End If

	If Not bRemoveSubExp Then
		Call GetExpressionTextforQual(aConnectionInfo, sATName, sFMName, "", sOperator, sInputAux, lExpType, sExpressionText)
		oOperatorNode.DisplayName = sExpressionText
	Else
		Call oSinglePrompt.ExpressionObject.RootNode.RemoveChild(oOperatorNode)
		Call oOperatorNode.clear()
	End If

	If lErrNumber <> NO_ERR Then
		AddSingleSelectionToHelperObjectForAttributeQual = lErrNumber
	Else
		AddSingleSelectionToHelperObjectForAttributeQual = Err.Number
	End If

    Err.Clear
End Function


Function AreMultipleAnswersSupportedByOperator(sOperator, lInputArraySize)
'***************************************************************************************************
'Purpose:   To determine whether this operator supports multiple answers in ;-delimited list format
'Inputs:	sOperator, lInputArraySize
'Outputs:	Error Value
'***************************************************************************************************
	Dim lErrNumber

	lErrNumber = 0

	If lInputArraySize > 0 Then
		Select Case Clng(Right(sOperator, Len(sOperator)-1))
		        Case DssXmlFunctionEquals, DssXmlFunctionNotEqual, DssXmlFunctionGreater, _
		             DssXmlFunctionGreaterEqual, DssXmlFunctionLess, DssXmlFunctionLessEqual

						lErrNumber = -1
		End Select
	End If

	AreMultipleAnswersSupportedByOperator = lErrNumber
End Function

Function AddSingleSelectionToHelperObjectForMetricQual(aConnectionInfo, sMTName, sMTDID, lDDT, sOperator, sInput, lExpType, oSinglePrompt)
'***************************************************************************************************
'Purpose:   append single answer to oSinglePrompt for Element prompt
'Inputs:    aConnectionInfo, sEI, oSinglePrompt
'Outputs:   oSinglePrompt
'***************************************************************************************************
    On Error Resume Next
    Dim oExpression
    Dim oOperatorNode
    Dim sOperatorType
    Dim sOperatorValue
    Dim aInput
    Dim lIndex
    Dim sExpressionText
    Dim oShortCutNode
    Dim lInputArraySize
    Dim lErrNumber, sOP

    Set oExpression = oSinglePrompt.ExpressionObject
    sOperatorType = Left(sOperator, 1)
    sOperatorValue = Mid(sOperator, 2, Len(sOperator) - 1)

    Select Case sOperatorType
        Case OperatorType_Metric
            Set oOperatorNode = oExpression.CreateOperatorNode(DssXmlFilterSingleMetricQual, CInt(sOperatorValue))
            Set oShortCutNode = oExpression.CreateShortCutNode(sMTDID, DssXmlTypeMetric, oOperatorNode)
            oOperatorNode.Child(1).DisplayName = sMTName
        Case OperatorType_Rank
            Set oOperatorNode = oExpression.CreateMetricRankOperatorNode(CInt(sOperatorValue), sMTDID)
            oOperatorNode.Child(1).Child(1).DisplayName = sMTName
        Case OperatorType_Percent
            Set oOperatorNode = oExpression.CreateMetricPercentOperatorNode(CInt(sOperatorValue), sMTDID)
            oOperatorNode.Child(1).Child(1).DisplayName = sMTName
    End Select
    Call GetExpressionTextforQual(aConnectionInfo, "", "", sMTName, sOperator, sInput, lExpType, sExpressionText)
    oOperatorNode.DisplayName = sExpressionText

    aInput = Split(sInput, ";", -1, vbBinaryCompare)
    lInputArraySize = UBound(aInput)
    lErrNumber = AreMultipleAnswersSupportedByOperator(sOperator, lInputArraySize)

    If lErrNumber = 0 Then
		For lIndex = 0 To lInputArraySize
			If Len(aInput(lIndex)) > 0 Then
    			aInput(lIndex) = Replace(aInput(lIndex), " " ,"")
			End If
			Call oExpression.CreateConstantNode(aInput(lIndex), lDDT, oOperatorNode)
		Next
	Else
	    AddSingleSelectionToHelperObjectForMetricQual = -1
	    Call oSinglePrompt.ExpressionObject.RootNode.RemoveChild(oSinglePrompt.ExpressionObject.RootNode.LastChild)
	    'To prevent memory leak, GA2
		Call oSinglePrompt.ExpressionObject.RootNode.LastChild.Clear()

        Exit Function
	End If

    AddSingleSelectionToHelperObjectForMetricQual = Err.Number
    Err.Clear
End Function

Function RemoveSelectionsFromExpressionPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   Remove user selections from oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, lExpType, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim lResult
    Dim oDefault
    Dim lExpType
    Dim bUnknownDef

    lExpType = oSinglePrompt.ExpressionType
    If CStr(oRequest("selected_" & sPin)) = "-default-" Then
        Exit Function
    End If

    Select Case oSinglePrompt.ExpressionType
    Case DssXmlFilterSingleMetricQual, DssXmlFilterAttributeIDQual, DssXmlFilterAttributeDESCQual
        lResult = RemoveSelectionsFromAQMQPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
        If lResult <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "RemoveSelectionsFromExpressionPrompt", "", "Error in call to RemoveSelectionsFromAQMQ", LogLevelTrace)
        End If

    Case DssXmlFilterAllAttributeQual
        lResult = RemoveSelectionsFromHIPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
        If lResult <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "RemoveSelectionsFromExpressionPrompt", "", "Error in call to RemoveSelectionsFromHI", LogLevelTrace)
        End If
    End Select

    If lResult = NO_ERR Then
        Call CO_GetbUnknownDef(oSinglePromptTempXML, bUnknownDef)
        If bUnknownDef And oSinglePrompt.ExpressionObject.Count = 0 Then
            Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "1")
            Call oSinglePrompt.SetDefaultAsAnswer
        End If
    End If

    RemoveSelectionsFromExpressionPrompt = lResult
    Err.Clear
End Function

Function RemoveSelectionsFromAQMQPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   Remove user selections from oSinglePromptTempXML For AQ/MQ expression prompt
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim oExpItem
    Dim oEXP
    Dim sSelection
    Dim sATDID
    Dim sFMDID
    Dim sMTDID
    Dim sOP
    Dim lErrNumber
    Dim temArray
    Dim sID
    Dim oExpression
    Dim oNode
    Dim sOPTP
    Dim aSelected

	'aSelected = Split(oRequest("selected_" & sPin), ", ", -1, vbBinaryCompare)

	aSelected = SplitRequest(oRequest("selected_" & sPin))

    If UBound(aSelected) > -1 Then
        For Each sSelection In aSelected
            temArray = Split(CStr(sSelection), Chr(27), -1, vbBinaryCompare)
            If UBound(temArray) > 1 Then
                sID = temArray(UBound(temArray))
                Set oExpression = oSinglePrompt.ExpressionObject
                Set oNode = oExpression.FindNodeByDisplayID(CLng(sID))
                If oNode.ExpressionType = DssXmlFilterSingleBaseFormQual Then
                    sATDID = oNode.AttributeID
                    sFMDID = oNode.AttributeFormID
                    sMTDID = ""
                ElseIf oNode.ExpressionType = CLng(ND_DssXmlFilterSingleMetricQual) Then
                    sATDID = ""
                    sFMDID = ""
                    sMTDID = oNode.Child(1).ShortcutID
                    If Len(sMTDID) = 0 Then
                        sMTDID = oNode.Child(1).Child(1).ShortcutID
                    End If
                End If
                sOPTP = CStr(oNode.OperatorType)
                Select Case sOPTP
                    Case "1", ""
                        sOPTP = OperatorType_Metric
                        sOP = oNode.Operator
                    Case "2"
                        sOPTP = OperatorType_Rank
                        sOP = oNode.MRPOperator
                    Case "3"
                        sOPTP = OperatorType_Percent
                        sOP = oNode.MRPOperator
                End Select
                sOP = sOPTP & sOP
                Call CO_SetRemoveByName(sATDID, sFMDID, sMTDID, sOP, oSinglePromptTempXML)
                Call oNode.Parent.RemoveChild(oNode)
                'To prevent memory leak, GA2
				Call oNode.Clear()
            End If
        Next
    End If
    lErrNumber = Err.Number

    Set oExpItem = Nothing
    Set oEXP = Nothing
    RemoveSelectionsFromAQMQPrompt = lErrNumber
    Err.Clear
End Function

Function RemoveSelectionsFromHIPrompt(aConnectionInfo, oRequest, sPin, oSinglePrompt, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   Remove user selections from oSinglePromptTempXML For Hierachical expression prompt
'Inputs:    aConnectionInfo, oRequest, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim oRootXML
    Dim sATID
    Dim oEXP
    Dim oExpItem
    Dim oE
    Dim sSelection
    Dim sID
    Dim sElementID
    Dim temArray
    Dim oExpression
    Dim oNode
    Dim oElements
    Dim sErrDescription
    Dim aSelected

	'aSelected = Split(oRequest("selected_" & sPin), ", ", -1, vbBinaryCompare)
	 aSelected = SplitRequest(oRequest("selected_" & sPin))

    If UBound(aSelected) > -1 Then

        For Each sSelection In aSelected
            If StrComp(sSelection, "-none-", vbTextCompare) = 0 Or StrComp(sSelection, "-default-", vbTextCompare) = 0 Then
                Exit For
            End If

            'differentiate attribute/element selection
            temArray = Null
            temArray = Split(CStr(sSelection), Chr(27), -1, vbBinaryCompare)
            If UBound(temArray) > 3 Then 'AQ qualification
                sID = CStr(temArray(4))
                Set oExpression = oSinglePrompt.ExpressionObject
                Set oNode = oExpression.FindNodeByDisplayID(CLng(sID))
                Call oNode.Parent.RemoveChild(oNode)
                'To prevent memory leak, GA2
				Call oNode.Clear()
            Else
                temArray = Null
                temArray = Split(CStr(sSelection), Chr(30), -1, vbBinaryCompare)
                If InStr(1, temArray(0), ":", vbBinaryCompare) = 0 Then 'Attribute
                    sATID = temArray(2)
                    Set oExpression = oSinglePrompt.ExpressionObject
                    Set oNode = oExpression.FindNodeByDisplayID(CLng(sATID))
                    Call oNode.Parent.RemoveChild(oNode)
                    'To prevent memory leak, GA2
					Call oNode.Clear()
                Else    'Element
                    sATID = temArray(2)
                    sElementID = temArray(3)
                    Set oExpression = oSinglePrompt.ExpressionObject
                    Set oNode = oExpression.FindNodeByDisplayID(CLng(sATID))
                    Set oElements = oNode.ElementsObject
                    Call oElements.Remove(CLng(sElementID))
                    If oElements.Count = 0 Then
                        Call oNode.Parent.Parent.RemoveChild(oNode.Parent)
						'To prevent memory leak, GA2
						Call oNode.Parent.Clear()
                    End If
                End If
            End If
        Next
    End If

    RemoveSelectionsFromHIPrompt = Err.Number
    Err.Clear
End Function



Function FindNodeByElementID(aConnectionInfo, oExpression, sElementID, oNode, sErrDescription)
'***********************************************************************************************************
'Purpose:
'Inputs:
'Outputs:
'***********************************************************************************************************
    On Error Resume Next
    Dim lErrorNumber
    Dim oExpressionXML

    Call GetXMLDOM(aConnectionInfo, oExpressionXML, sErrDescription)

    If Len(oExpression.xml) > 0 Then
		Call oExpressionXML.loadXML(oExpression.xml)
	Else
		Set oNode = Nothing
	End If

	FindNodeByElementID = lErrorNumber
	Err.Clear
End Function



Function AddSelectionsToHelperObjectForAttributeQual(aConnectionInfo, aPromptInfo, oRequest, sPin, sFlag, lExpType, oSinglePrompt, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   Build an expression <nd> For AttributeIDQual Or AttributeDESCQual Expression Type
'Inputs:    aConnectionInfo, aPromptInfo, oRequest, sPin, sFlag, lExpType
'Outputs:   oSinglePrompt
'***********************************************************************************************************
    On Error Resume Next
    Dim sOperator
    Dim sInput
    Dim sAttribute
    Dim lResult
    Dim sAvailable, oAvailable
    Dim sCleanInput
    Dim temArray
    Dim sOP
    Dim sATDID
    Dim sATName
    Dim sFMDID
    Dim sFMName
    Dim sDT
    Dim aInput
    Dim sFilterOperator
    Dim lRootOperator
    Dim bUnknownDef
    Dim sFormInfo
	Dim aFormInfo
	Dim aFormList
	Dim lCount
	Dim sDDT
	Dim aAvailable

    'Set oSinglePromptTempXML = aPromptInfo(CLng(sPin), PROMPTINFO_O_TEMPANSWER)
    lResult = Err.Number
    If lResult = NO_ERR Then

		'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)
		aAvailable = SplitRequest(oRequest("available_" & sPin))

        For Each oAvailable In aAvailable
            sAvailable = CStr(oAvailable)
        Next
        sAttribute = CStr(oRequest("attribute_" & sPin))
        sInput = CStr(oRequest("input_" & sPin))
        sOperator = CStr(oRequest("operator_" & sPin))

        Select Case sFlag
        Case AQ_ATTRQUAL
			Call GetAttributeFormInfoFromString(sAvailable, sATDID, sFMDID, sATName, sFMName, sDT)
        Case AQ_HIPROMPT
            Call GetAttributeFormInfoFromString(sAttribute, sATDID, sFMDID, sATName, sFMName, sDT)
        End Select

        Call CO_CleanInput(DssXmlFilterAllAttributeQual, sInput, sCleanInput)
        sInput = sCleanInput

        If StrComp(sOperator, OperatorType_Metric & CStr(DssXmlFunctionBetween), vbTextCompare) = 0 Or StrComp(sOperator, OperatorType_Metric & CStr(DssXmlFunctionNotBetween), vbTextCompare) = 0 Then
            aInput = Split(sInput, ";", -1, vbBinaryCompare)
            If UBound(aInput) < 1 Or UBound(aInput) >= 2 Then
                If StrComp(sOperator, OperatorType_Metric & CStr(DssXmlFunctionBetween), vbTextCompare) = 0 Then
                    Call CO_SetPromptError(oSinglePromptTempXML, ERR_BETWEEN_EXPECTS_TWOVALUES)
                Else
                    Call CO_SetPromptError(oSinglePromptTempXML, ERR_NOTBETWEEN_EXPECTS_TWOVALUES)
                End If
                Call CO_SetRemoveByName(sATDID, sFMDID, "", sOperator, oSinglePromptTempXML)
                lResult = ERR_BETWEEN_EXPECTS_TWOVALUES
            Else
                aInput(0) = Trim(aInput(0))
                aInput(1) = Trim(aInput(1))
                If Len(aInput(0)) = 0 Or Len(aInput(1)) = 0 Then
                    If StrComp(sOperator, OperatorType_Metric & CStr(DssXmlFunctionBetween), vbTextCompare) = 0 Then
                    Call CO_SetPromptError(oSinglePromptTempXML, ERR_BETWEEN_EXPECTS_TWOVALUES)
                Else
                    Call CO_SetPromptError(oSinglePromptTempXML, ERR_NOTBETWEEN_EXPECTS_TWOVALUES)
                End If
                Call CO_SetRemoveByName(sATDID, sFMDID, "", sOperator, oSinglePromptTempXML)
                    lResult = ERR_BETWEEN_EXPECTS_TWOVALUES
                End If
            End If
        End If

        If Len(sInput) > 0 And (lResult = 0 Or lResult = ERR_BETWEEN_EXPECTS_TWOVALUES) Then

			If Len(sFMDID)>0 And Len(sFMName)>0 Then
				lResult = AddSingleSelectionToHelperObjectForAttributeQual(aConnectionInfo, sATName, sATDID, sFMName, sFMDID, CLng(sDT), sOperator, sInput, lExpType, oSinglePrompt, oSinglePromptTempXML, aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE))
			Else
				sFormInfo = CStr(oRequest("form_" & sPin & "_" & sATDID))

				If Len(sFormInfo)>0 Then
					'aFormList = Split(sFormInfo, "," , -1, vbBinaryCompare)
					aFormList = SplitRequest(oRequest("form_" & sPin & "_" & sATDID))
					oSinglePrompt.ExpressionObject.RootNode.Operator = DssXmlFunctionOr

					For lCount=0 To UBound(aFormList)
						aFormInfo = Split(aFormList(lCount), Chr(30) , -1, vbBinaryCompare)
						sFMDID = aFormInfo(0)
						sFMName = aFormInfo(1)
						sDDT = Trim(aFormInfo(2))

						lResult = AddSingleSelectionToHelperObjectForAttributeQual(aConnectionInfo, sATName, sATDID, sFMName, sFMDID, CLng(sDDT), sOperator, sInput, lExpType, oSinglePrompt, oSinglePromptTempXML, aPromptInfo(CLng(sPin), PROMPTINFO_S_XSLFILE))
						If lResult <> 0 Then
							Break
						End If
					Next
				Else
					lResult = -1
				End If

			End If

            If lResult <> 0 Then
                Call CO_SetRemoveByName(sATDID, sFMDID, "", sOperator, oSinglePromptTempXML)
				lResult = ERR_TOOMANY_SELECTIONS_EXPRESSIONPROMPT
				Call CO_GetbUnknownDef(oSinglePromptTempXML, bUnknownDef)
			    If bUnknownDef Then
					Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "1")
					Call oSinglePrompt.SetDefaultAsAnswer
                End If
            End If
        ElseIf Len(sInput) = 0 And oSinglePrompt.Required Then
            Call CO_SetPromptError(oSinglePromptTempXML, ERR_OPERATOR_EXPECTS_VALUE)
            Call CO_SetRemoveByName(sATDID, sFMDID, "", sOperator, oSinglePromptTempXML)
            lResult = ERR_OPERATOR_EXPECTS_VALUE
        End If
    End If

    AddSelectionsToHelperObjectForAttributeQual = lResult
    Err.Clear
End Function

Function AddSelectionsToHelperObjectForMetricQual(aConnectionInfo, oRequest, sPin, lExpType, oSinglePrompt, oSinglePromptTempXML)
'***********************************************************************************************************
'Purpose:   Build an expression <nd> For AttributeIDQual Or AttributeDESCQual Expression Type
'Inputs:    aConnectionInfo, sPin, lExpType, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, sFlag
'Outputs:   oEXP
'***********************************************************************************************************
    On Error Resume Next
    Dim sOperator
    Dim sInput
    Dim sAttribute
    Dim lResult
    Dim sAvailable, oAvailable
    Dim sCleanInput
    Dim temArray
    Dim sOP
    Dim sMTDID
    Dim sMTName
    Dim aInput
    Dim sFilterOperator
    Dim bUnknownDef
    Dim aAvailable

    lResult = Err.Number
    If lResult = NO_ERR Then
		'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)
		aAvailable = SplitRequest(oRequest("available_" & sPin))

        For Each oAvailable In aAvailable
            sAvailable = CStr(oAvailable)
        Next
        'sAvailable = CStr(oRequest("available_" & sPin)(1))
        sInput = CStr(oRequest("input_" & sPin))
        sOperator = CStr(oRequest("operator_" & sPin))
        temArray = Split(sAvailable, Chr(30), -1, vbBinaryCompare)

        sMTDID = CStr(temArray(0))
        sMTName = CStr(temArray(1))

        Call CO_CleanInput(DssXmlFilterSingleMetricQual, sInput, sCleanInput)
        sInput = sCleanInput

        If StrComp(sOperator, OperatorType_Metric & CStr(DssXmlFunctionBetween), vbTextCompare) = 0 Or StrComp(sOperator, OperatorType_Metric & CStr(DssXmlFunctionNotBetween), vbTextCompare) = 0 Then
            aInput = Split(sInput, ";", -1, vbBinaryCompare)
            If UBound(aInput) < 1 Then
                If StrComp(sOperator, OperatorType_Metric & CStr(DssXmlFunctionBetween), vbTextCompare) = 0 Then
                    Call CO_SetPromptError(oSinglePromptTempXML, ERR_BETWEEN_EXPECTS_TWOVALUES)
                Else
                    Call CO_SetPromptError(oSinglePromptTempXML, ERR_NOTBETWEEN_EXPECTS_TWOVALUES)
                End If
                Call CO_SetRemoveByName("", "", sMTDID, sOperator, oSinglePromptTempXML)
                lResult = ERR_BETWEEN_EXPECTS_TWOVALUES
            Else
                aInput(0) = Trim(aInput(0))
                aInput(1) = Trim(aInput(1))
                If Len(aInput(0)) = 0 Or Len(aInput(1)) = 0 Then
                    If StrComp(sOperator, OperatorType_Metric & CStr(DssXmlFunctionBetween), vbTextCompare) = 0 Then
                    Call CO_SetPromptError(oSinglePromptTempXML, ERR_BETWEEN_EXPECTS_TWOVALUES)
                Else
                    Call CO_SetPromptError(oSinglePromptTempXML, ERR_NOTBETWEEN_EXPECTS_TWOVALUES)
                End If
                Call CO_SetRemoveByName("", "", sMTDID, sOperator, oSinglePromptTempXML)
                    lResult = ERR_BETWEEN_EXPECTS_TWOVALUES
                End If
            End If
        End If

        If Len(sInput) > 0 And Len(sMTDID) > 0 And lResult = 0 Then
            If Left(sOperator, 1) = OperatorType_Percent Then
                sInput = Replace(sInput, "%", "", 1, -1, vbBinaryCompare) & "%"
            End If
            lResult = AddSingleSelectionToHelperObjectForMetricQual(aConnectionInfo, sMTName, sMTDID, DssXmlDataTypeReal, sOperator, sInput, lExpType, oSinglePrompt)
            If lResult <> 0 Then
               Call CO_SetPromptError(oSinglePromptTempXML, ERR_TOOMANY_SELECTIONS_EXPRESSIONPROMPT)
	           Call GetOperatorID(sOperator, sOP)
		       Call CO_SetRemoveByName("", "", sMTDID, sOperator, oSinglePromptTempXML)
		       Call CO_GetbUnknownDef(oSinglePromptTempXML, bUnknownDef)
			   If bUnknownDef Then
					Call CO_SetDisplayUnknownDef(oSinglePromptTempXML, "1")
					Call oSinglePrompt.SetDefaultAsAnswer
               End If
			   lResult = ERR_TOOMANY_SELECTIONS_EXPRESSIONPROMPT
            End If
        ElseIf Len(sInput) = 0 Then
            Call CO_SetPromptError(oSinglePromptTempXML, ERR_OPERATOR_EXPECTS_VALUE)
            Call GetOperatorID(sOperator, sOP)
            Call CO_SetRemoveByName("", "", sMTDID, sOperator, oSinglePromptTempXML)
            lResult = ERR_OPERATOR_EXPECTS_VALUE
        ElseIf Len(sMTDID) = 0 Then
            Call CO_SetPromptError(oSinglePromptTempXML, ERR_NOT_QUALIFY_ZERO)
            Call GetOperatorID(sOperator, sOP)
            Call CO_SetRemoveByName("", "", sMTDID, sOperator, oSinglePromptTempXML)
            lResult = ERR_NOT_QUALIFY_ZERO
        End If
    End If

    AddSelectionsToHelperObjectForMetricQual = lResult
    Err.Clear
End Function

Function BuildElementAnswerXMLFromUserSelections(aConnectionInfo, sElementSelections, oSinglePrompt)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given element prompt, add them to oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, sElementSelections, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim asSelectionArray
    Dim lIndex
    Dim oElements
    Dim lErrNumber

    If Err.Number = NO_ERR Then
        asSelectionArray = Split(sElementSelections, Chr(28), -1, vbBinaryCompare)
        Set oElements = oSinglePrompt.ElementsObject
        For lIndex = 0 To UBound(asSelectionArray) - 1
            If StrComp(asSelectionArray(lIndex), "-none-", vbTextCompare) <> 0 Then
                lErrNumber = AddSingleSelectionToHelperObjectForElementPrompt(aConnectionInfo, asSelectionArray(lIndex), oElements)
            End If
            If lErrNumber <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "BuildElementAnswerFromUserSelections", "", "Error in call to UpdateExpressionNode_AllAttributeQual", LogLevelTrace)
            End If
        Next
    End If

    BuildElementAnswerXMLFromUserSelections = lErrNumber
    Err.Clear
End Function

Function BuildObjectAnswerXMLFromUserSelections(aConnectionInfo, sPin, sObjectSelections, oSinglePrompt)
'***********************************************************************************************************
'Purpose:   read all the selections in oRequest For a given object prompt, add them to oSinglePromptTempXML
'Inputs:    aConnectionInfo, oRequest, sPin, sObjectSelections, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim asSelectionArray
    Dim lIndex
    Dim sSelection
    Dim lObjType
    Dim sName
    Dim sObjectDID
    Dim aObject
    Dim oRootXML
    Dim oFolder

    If Err.Number = NO_ERR Then
        Set oFolder = oSinglePrompt.FolderObject
        asSelectionArray = Split(sObjectSelections, Chr(28), -1, vbBinaryCompare)
        For lIndex = 0 To UBound(asSelectionArray) - 1
            aObject = Split(asSelectionArray(lIndex), Chr(30), -1, vbBinaryCompare)
            sObjectDID = aObject(0)
            lObjType = CLng(aObject(1))
            sName = DecodeStr(aObject(2))
            If Len(sObjectDID) > 0 And StrComp(sObjectDID, "-none-", vbTextCompare) <> 0 Then
                Call oFolder.AddCopy(sObjectDID, lObjType, sName)
                If Err.Number <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "BuildElementAnswerFromUserSelections", "", "Error in call to UpdateExpressionNode_AllAttributeQual", LogLevelTrace)
                End If
            End If
        Next
    End If

    BuildObjectAnswerXMLFromUserSelections = lErrNumber
    Err.Clear
End Function


Function BuildExpressionAnswerFromUserSelections(aConnectionInfo, oRequest, sPin, lExpType, sExpressionSelections, sXsl, oSinglePromptQuestionXML, oSinglePromptTempXML, oSinglePrompt)
'***********************************************************************************************************
'Purpose:   Build answerXML from sUserSelections for Expression prompt
'Inputs:    aConnectionInfo, oRequest, sPin, lExpType, sExpressionSelections, oSinglePromptQuestionXML, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim lResult
    Dim aSelections
    Dim asSelectionArray
    Dim sAvailable
    Dim sAttribute
    Dim sInput
    Dim sOperator
    Dim sFilterOperator
    Dim sSelection
    Dim aMT
    Dim sMT
    Dim sMTName
    Dim sMTDID
    Dim aMTs
    Dim aAttribute
    Dim sATDID
    Dim sATName
    Dim sFMDID
    Dim sFMName
    Dim sDT
    Dim lIndex

    sFilterOperator = CStr(oRequest("filteroperator_" & sPin))
    lResult = Err.Number
    If lResult = NO_ERR Then
        asSelectionArray = Split(sExpressionSelections, Chr(28), -1, vbBinaryCompare)
        If Err.Number = NO_ERR Then
            Select Case oSinglePrompt.ExpressionType
            Case DssXmlFilterSingleMetricQual
                For lIndex = 0 To UBound(asSelectionArray) - 1
                    sSelection = asSelectionArray(lIndex)
                    aSelections = Split(sSelection, Chr(27), -1, vbBinaryCompare)
                    sMT = aSelections(0)
                    aMT = Split(sMT, Chr(30), -1, vbBinaryCompare)
                    sMTDID = aMT(0)
                    sMTName = aMT(1)
                    sOperator = aSelections(1)
                    sInput = aSelections(2)

                    lResult = AddSingleSelectionToHelperObjectForMetricQual(aConnectionInfo, sMTName, sMTDID, DssXmlDataTypeReal, sOperator, sInput, lExpType, oSinglePrompt)
                    If lResult <> 0 Then
                        Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "BuildExpressionNode_SingleMetricQual", "", "Error in call to CO_AddSingleMetricQualtoAnswerXML", LogLevelTrace)
                    End If
                Next

            Case DssXmlFilterAllAttributeQual, DssXmlExpressionMDXSAPVariable
                lResult = BuildExpressionAnswerFromUserSelections_HierPrompt(aConnectionInfo, sPin, lExpType, asSelectionArray, sXsl, oSinglePromptQuestionXML, oSinglePromptTempXML, oSinglePrompt, oRequest)
                If lResult <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "AddSelectionsToExpressionPrompt", "", "Error in call to BuildExpressionNode_HierPrompt", LogLevelTrace)
                End If

            Case DssXmlFilterAttributeIDQual, DssXmlFilterAttributeDESCQual
                For lIndex = 0 To UBound(asSelectionArray) - 1
                    sSelection = asSelectionArray(lIndex)
                    aSelections = Split(sSelection, Chr(27), -1, vbBinaryCompare)
                    sAttribute = aSelections(0)

					Call GetAttributeFormInfoFromString(sAttribute, sATDID, sFMDID, sATName, sFMName, sDT)

                    sInput = aSelections(2)
                    If StrComp(Left(sInput, Len(";")), ";", vbBinaryCompare) = 0 Then
                        sInput = Right(sInput, Len(sInput) - Len(";"))
                    End If
                    If StrComp(Right(sInput, Len(";")), ";", vbBinaryCompare) = 0 Then
                        sInput = Left(sInput, Len(sInput) - Len(";"))
                    End If
                    sOperator = aSelections(1)
                    If Len(sAttribute) > 0 And Len(sInput) > 0 Then
                        lResult = AddSingleSelectionToHelperObjectForAttributeQual(aConnectionInfo, sATName, sATDID, sFMName, sFMDID, CLng(sDT), sOperator, sInput, lExpType, oSinglePrompt, oSinglePromptTempXML, sXsl)
                        If lResult <> 0 Then
                            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "BuildExpressionNode_SingleMetricQual", "", "Error in call to CO_AddSingleMetricQualtoAnswerXML", LogLevelTrace)
                        End If
                    Else
                        Call CO_SetPromptError(oSinglePromptTempXML, ERR_ANSWERPROMPT)
                        lResult = ERR_ANSWERPROMPT
                    End If
                Next
            Case Else
                lResult = ERR_CUSTOM_UNKNOWN_EXPPROMPT_TYPE
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "AddSelectionsToExpressionPrompt", "", "Unknown Prompt Type", LogLevelError)
            End Select
        End If
    End If

    BuildExpressionAnswerFromUserSelections = lResult
    Err.Clear
End Function

Function BuildExpressionAnswerFromUserSelections_HierPrompt(aConnectionInfo, sPin, lExpType, asSelectionArray, sXsl, oSinglePromptQuestionXML, oSinglePromptTempXML, oSinglePrompt, oRequest)
'***********************************************************************************************************
'Purpose:   From sUserSelections, build an expression <nd> For Hierachical Prompt (either Elem Or Qual)
'Inputs:    aConnectionInfo, sPin, lExpType, asSelectionArray, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, bNeedChange
'Outputs:   oSinglePromptTempXML
'***********************************************************************************************************
    On Error Resume Next
    Dim temArray
    Dim sFMDID
    Dim sFMName
    Dim sATDID
    Dim sATName
    Dim lResult
    Dim sSelection
    Dim aSelections
    Dim aAttribute
    Dim lBeginIndex
    Dim lEndIndex
    Dim sAvailable
    Dim sAttribute
    Dim sOperator
    Dim sInput
    Dim j
    Dim sFilterOperator
    Dim sDT

    sFilterOperator = CStr(oRequest("filteroperator_" & sPin))
    lResult = Err.Number
    If lResult = NO_ERR Then
        For j = 0 To UBound(asSelectionArray) - 1
            sSelection = asSelectionArray(j)
            If InStr(1, sSelection, Chr(27), vbBinaryCompare) > 0 Then 'Selection on expression
                    aSelections = Split(sSelection, Chr(27), -1, vbBinaryCompare)
                    sAttribute = aSelections(0)

                    Call GetAttributeFormInfoFromString(sAttribute, sATDID, sFMDID, sATName, sFMName, sDT)

                    sOperator = aSelections(1)
                    sInput = aSelections(2)
                    If Len(sAttribute) > 0 And Len(sInput) > 0 Then
                        lResult = AddSingleSelectionToHelperObjectForAttributeQual(aConnectionInfo, sATName, sATDID, sFMName, sFMDID, CLng(sDT), sOperator, sInput, lExpType, oSinglePrompt, oSinglePromptTempXML, sXsl)
                        If lResult <> 0 Then
                            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "BuildExpressionNode_SingleMetricQual", "", "Error in call to CO_AddSingleMetricQualtoAnswerXML", LogLevelTrace)
                        End If
                    Else
						Call CO_SetPromptError(oSinglePromptTempXML, 899)
                        lResult = ERR_ANSWERPROMPT
                    End If
            Else
                aSelections = Split(sSelection, Chr(30), -1, vbBinaryCompare)
                aAttribute = Split(aSelections(0), ":", -1, vbBinaryCompare)
                If UBound(aAttribute) = 0 Then  'Selection on attribute
                    sATDID = aSelections(0)
                    sATName = aSelections(1)
                    lBeginIndex = j + 1
                    For lEndIndex = lBeginIndex To UBound(asSelectionArray) - 1
                        aAttribute = Split(asSelectionArray(lEndIndex), ":", -1, vbBinaryCompare)
                        If (aAttribute(0) <> sATDID) Then
                            Exit For
                        End If
                    Next
                    lEndIndex = lEndIndex - 1
                    j = lEndIndex
                    lResult = AddSelectionsToHelperObjectForElementListinHIFromUserSelection(aConnectionInfo, sPin, sATDID, sATName, asSelectionArray, lBeginIndex, lEndIndex, oSinglePrompt)
                    If lResult = ERR_ONLYADDNONE_HIPROMPT Then
                    ElseIf lResult <> 0 Then
                        Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "BuildExpressionNode_HierPrompt", "", "Error in call to BuildExpressionNode_AllAttributeQual", LogLevelTrace)
                        Err.Raise lResult
                    End If
                End If
            End If
        Next
    Else
        Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "BuildExpressionNode_HierPrompt", "", "Error setting ", LogLevelError)
    End If

    BuildExpressionAnswerFromUserSelections_HierPrompt = lResult
    Err.Clear
End Function

'===============From ProcessCulib.asp
Function CO_CleanInput(sPromptType, sInput, sCleanInput)
'***********************************************************************************************
' Purpose:  Replace ";" to "," in sInput, delete all non-data type input item for data type form
' Inputs:   sInput
' Outputs:  sDelimeter
'***********************************************************************************************
    On Error Resume Next
    Dim sTemp
    Dim aInput
    Dim lInputLength
    Dim lInputIndex

    sTemp = sInput

    aInput = Split(sTemp, ";", -1, vbTextCompare)
    lInputLength = UBound(aInput) - LBound(aInput) + 1

    'Clean input
    sCleanInput = ""
    For lInputIndex = 0 To lInputLength - 1
        If Len(aInput(lInputIndex)) > 0 Then
            sCleanInput = sCleanInput & aInput(lInputIndex) & ";"
        End If
    Next
    If Len(sCleanInput) > 0 Then
        sCleanInput = Left(sCleanInput, Len(sCleanInput) - Len(";"))
    End If

    CO_CleanInput = Err.Number
    Err.Clear
End Function

Function CO_CleanInputForDataType(sInput, sDataType, sCleanInput)
'***********************************************************************************************
' Purpose:  Replace ";" to "," in sInput, delete all non-data type input item for data type form
' Inputs:   sInput
' Outputs:  sDelimeter
'***********************************************************************************************
    On Error Resume Next
    Dim aInput
    Dim lIndex

    sCleanInput = sInput
    If CLng(sDataType) = DssXmlBaseFormNumber Then
        aInput = Split(sInput, ";", -1, vbTextCompare)
        sCleanInput = ""
        For lIndex = 0 To UBound(aInput)
            If IsNumeric(aInput(lIndex)) Then
                sCleanInput = sCleanInput & aInput(lIndex) & ";"
            End If
        Next
        If Len(sCleanInput) > 0 Then
            sCleanInput = Left(sCleanInput, Len(sCleanInput) - Len(";"))
        End If
    End If

    CO_CleanInputForDataType = Err.Number
    Err.Clear
End Function

Function LoadTextFile(aConnectionInfo, oRequest, sPin, lExpType, sXsl, oSinglePrompt, oSinglePromptTempXML)
'***********************************************************************************************
' Purpose:  Load a text file, insert IN answer
' Inputs:
' Outputs:
'***********************************************************************************************
    On Error Resume Next
    Dim sFileContents
    Dim sAvailable, oAvailable
    Dim temArray
    Dim sATDID
    Dim sFMDID
    Dim sATName
    Dim sFMName
    Dim sDT
    Dim sOperator
    Dim sInput
    Dim i
    Dim aAvailable
    Dim lTotalElements

    If CheckFileAgainstAdminPreferences(aConnectionInfo, oRequest, "nuXML_textfile_" & sPin, sErrDescription) Then

        If InStr(1, oRequest("nuXML_textfile_" & sPin), ",", vbBinaryCompare) > 0 Then
			sFileContents = Left(oRequest("nuXML_textfile_" & sPin), InStr(1, oRequest("nuXML_textfile_" & sPin), ",", vbBinaryCompare) - 1)
		Else
			sFileContents	= oRequest("nuXML_textfile_" & sPin)
		End If

		While (InStr(1, sFileContents, ";", vbBinaryCompare) > 0)
            sFileContents = Replace(sFileContents, ";", vbNewLine)
        Wend

        While (InStr(1, sFileContents, vbNewLine & vbNewLine, vbBinaryCompare) > 0)
            sFileContents = Replace(sFileContents, vbNewLine & vbNewLine, vbNewLine)
        Wend

        temArray = Split(sFileContents, vbNewLine, ReadUserOption(MAX_ELEMENTS_TO_IMPORT_OPTION), vbBinaryCompare)

        lTotalElements = 0
        lTotalElements = UBound(temArray)
        For i = 0 To lTotalElements
			If (Strcomp(ReadUserOption(KEEP_WHITESPACE_IN_PROMPTS_OPTION),"checked",vbTextCompare) <> 0)  Then
				temArray(i) = Trim(temArray(i))
			End If
        Next

		sInput = Join(temArray, ";")

        sInput = Left(sInput, InStr(1, sInput, vbNewLine, vbBinaryCompare) - 1)
        If (StrComp(Right(sInput, Len(";")), ";", vbBinaryCompare) = 0) Then
            sInput = Left(sInput, Len(sInput) - Len(";"))
        End If

        'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)

        aAvailable = SplitRequest(oRequest("available_" & sPin))

        For Each oAvailable In aAvailable
            sAvailable = CStr(oAvailable)
        Next

		Call GetAttributeFormInfoFromString(sAvailable, sATDID, sFMDID, sATName, sFMName, sDT)

		If Len(oRequest("operator_" & sPin)) > 0 Then
			sOperator = oRequest("operator_" & sPin)
			If StrComp(sOperator,OperatorType_Metric & CStr(DssXmlFunctionNotIn),vbTextCompare) <> 0 Then
				sOperator = OperatorType_Metric & CStr(DssXmlFunctionIn)
			End If
		Else
			sOperator = OperatorType_Metric & CStr(DssXmlFunctionIn)
		End If

        lErrNumber = AddSingleSelectionToHelperObjectForAttributeQual(aConnectionInfo, sATName, sATDID, sFMName, sFMDID, CLng(sDT), sOperator, sInput, lExpType, oSinglePrompt, oSinglePromptTempXML, sXsl)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptProcessCuLib.asp", "BuildExpressionNode_HierPrompt", "", "Error setting ", LogLevelError)
        End If
    Else
        Call CO_SetPromptError(oSinglePromptTempXML, ERR_TEXTFILE_NOT_VALID)
    End If

    LoadTextFile = lErrNumber
    Err.Clear
End Function


Function GetAttributeFormInfoFromString(sAvailable, sATDID, sFMDID, sATName, sFMName, sDT)
'***********************************************************************************************
' Purpose:  split string to get 5 values for an attribute form
' Inputs:
' Outputs:
'***********************************************************************************************
    On Error Resume Next
	Dim temArray
	Dim sPos

	temArray = Split(sAvailable, Chr(30), -1, vbBinaryCompare)
	sATDID = CStr(temArray(0))
	sFMDID = CStr(temArray(1))
	If UBound(temArray)=4 Then	'atdid(30)fmdid(30)atname(30)fmname(30)dt
		sATName = temArray(2)
		sFMName = temArray(3)
		sDT = temArray(4)
	Else			'atdid(30)fmdid(30)atname\fmname(30)dt
		sPos = Instr(1, temArray(2), "\", vbBinaryCompare)
		sATName = Left(temArray(2), sPos-1)
		sFMName = Right(temArray(2), len(temArray(2))-sPos )
		sDT = temArray(3)
	End If

	If Len(sDT) = 0 Then
		sDT = CStr(DssXmlDataTypeTimeStamp)
	End If

	GetAttributeFormInfoFromString = Err.number
	Err.Clear
End Function


Function ValidateSegmentDataType(sCurSegment, lConstType)
'***************************************************************************************************
'Purpose:  Validate that segment value data type is correct
'Input:     sCurSegment, lConstType
'Output:    Boolean value indicating that value has correct data type or not
'***************************************************************************************************
	Dim bValid

	bValid = False

	Select Case CLng(lConstType)
		Case DssXmlDataTypeReal, DssXmlDataTypeNumeric
		    If IsNumeric(sCurSegment) Then
				bValid = True
		    End If
		Case DssXmlDataTypeDate, DssXmlDataTypeTime, DssXmlDataTypeTimeStamp
		    'If IsDate(sCurSegment) Then
			If IsValidDate(sCurSegment) Then
				bValid = True
		    End If
		Case DssXmlDataTypeChar, DssXmlDataTypeVarChar
			bValid = True
    End Select

	ValidateSegmentDataType = bValid
End Function

%>
