<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function ValidateAllPromptAnswers(aConnectionInfo, aPromptGeneralInfo, aPromptInfo, oRequest, bRemoveTemp)
'*************************************************************************************************************
'Purpose:   validate prompt answers
'Inputs:    aConnectionInfo, aPromptGeneralInfo(PROMPT_B_ISDOC), aPromptInfo, aPromptGeneralInfo(PROMPT_O_QUESTIONSXML), aPromptGeneralInfo(PROMPT_O_ANSWERSXML), oRequest, bRemoveTemp
'Outputs:   Err.Number
'*************************************************************************************************************
    On Error Resume Next
    Dim oSinglePromptQuestionXML
    Dim oSinglePromptAnswerXML
    Dim sPType
    Dim lResult
    Dim bSthWrong
    Dim sPin
    Dim lPin
    Dim sErrCode
    Dim sError
    Dim sUsed
    Dim sClosed

    bSthWrong = False
    Call CO_RemoveCurrentPin(aPromptGeneralInfo(PROMPT_O_ANSWERSXML))

    For Each oSinglePromptQuestionXML In aPromptGeneralInfo(PROMPT_O_QUESTIONPIFS)
        sUsed = oSinglePromptQuestionXML.getAttribute("usd")
        sClosed = oSinglePromptQuestionXML.getAttribute("cl")
        If sUsed = "1" And sClosed = "0" Then
            'sPType = oSinglePromptQuestionXML.getAttribute("pt")
            sPin = oSinglePromptQuestionXML.getAttribute("pin")
            lPin = CLng(sPin)
            sPType = aPromptInfo(lPin, PROMPTINFO_S_TYPE)

            Set oSinglePromptAnswerXML = aPromptGeneralInfo(PROMPT_O_ANSWERSXML).selectSingleNode("/mi/pif[@pin='"&sPin&"']")
            If oSinglePromptAnswerXML Is Nothing And Err.Number = 0 Then
                Err.Raise ERR_CUSTOM_NO_SPECIFIC_NODE
                Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), "PromptValidateCuLib.asp", "ValidateAllPromptAnswers", "", "Error working with the XML", LogLevelError)
            Else
                Call CO_GetPromptError(oSinglePromptAnswerXML, sError)
                If sError <> "0" Then   'If there is sth wrong in process, display it
                    bSthWrong = True
                Else
                    Select Case sPType
                    Case PROMPTTYPE_CONSTANT_LONG, PROMPTTYPE_CONSTANT_STRING, PROMPTTYPE_CONSTANT_DOUBLE, PROMPTTYPE_CONSTANT_DATE
                        lResult = ValidateConstantPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML)

                    Case PROMPTTYPE_OBJECT
                        lResult = ValidateObjectPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML)

                    Case PROMPTTYPE_ELEMENT
                        lResult = ValidateElementPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML)

                    Case PROMPTTYPE_EXPRESSION
                        'if aPromptInfo(CLng(sPin), PROMPTINFO_S_SUBTYPE) = EXPRESSIONTYPE_ALLATTRIBUTEQUAL then
                            'lResult = ValidateHierachicalPromptAnswer(aConnectionInfo, lPin, aPromptInfo, oSinglePromptQuestionXML, oSinglePromptAnswerXML, oRequest)
                        'else
                            lResult = ValidateExpressionPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML, oRequest)
                        'end if

                    Case PROMPTTYPE_LEVEL
                        lResult = ValidateLevelPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML)

                    Case Else
                        Err.Raise ERR_CUSTOM_UNKNOWN_PROMPT_TYPE
                        Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), "PromptValidateCuLib.asp", "ValidateAllPromptAnswers", "", "Unknown Prompt Type", LogLevelError)
                    End Select

                    If lResult <> 0 Then
                        Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), "PromptValidateCuLib.asp", "ValidateAllPromptAnswers", "", "Error in call to " & "ValidatePromptAnswer, PType=" & sPType, LogLevelTrace)
                        bSthWrong = True
                    End If

                    Call CO_GetPromptError(oSinglePromptAnswerXML, sErrCode)
                    If CStr(sErrCode) <> "0" And Len(CStr(sErrCode)) > 0 Then
                        bSthWrong = True
                        Call CO_SetCurrentPin(oSinglePromptAnswerXML)
                    End If
                End If
            End If
        End If
    Next

    If Err.Number = 0 And bRemoveTemp And Not bSthWrong Then
        lResult = CO_CleanUpAllPromptAnswers(aPromptInfo, aPromptGeneralInfo(PROMPT_O_ANSWERSXML))
        If lResult <> 0 Then
            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), "PromptValidateCuLib.asp", "ValidateAllPromptAnswers", "", "Error in call to " & "CO_CleanUpAllPromptAnswers", LogLevelTrace)
            Err.Raise Err.Number
        End If
    End If

    If bSthWrong = True Then
        ValidateAllPromptAnswers = ERR_VALIDATION_FAILED
    Else
        ValidateAllPromptAnswers = Err.Number
    End If

    Set oSinglePromptQuestionXML = Nothing
    Set oSinglePromptAnswerXML = Nothing

    Err.Clear
End Function

Function ValidateConstantPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML)
'*************************************************************************************************************
'Purpose:   validate constant prompt answers - remove leading/ending spaces
'Inputs:    aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML
'Outputs:   oSinglePromptAnswerXML
'*************************************************************************************************************
    On Error Resume Next
    Dim oPA
    Dim sPAText
    Dim bRequired
    Dim lMin
    Dim lMax

    sPAText = oSinglePromptAnswerXML.Text

    bRequired = aPromptInfo(lPin, PROMPTINFO_B_REQUIRED)
    If bRequired And Len(sPAText) = 0 Then  'If no answer
        Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_REQUIRED_PROMPT)
    Else
        While Left(sPAText, 1) = " "
            sPAText = Right(sPAText, Len(sPAText) - 1)
        Wend
        While Right(sPAText, 1) = " "
            sPAText = Left(sPAText, Len(sPAText) - 1)
        Wend
        oSinglePromptAnswerXML.Text = sPAText
        If aPromptInfo(lPin, PROMPTINFO_S_TYPE) = PROMPTTYPE_CONSTANT_STRING Then
            lMin = aPromptInfo(lPin, PROMPTINFO_L_MIN)
            lMax = aPromptInfo(lPin, PROMPTINFO_L_MAX)
            If lMax <> -1 And Len(sPAText) > lMax Then
                Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOLONG_TEXT_CONSTANTPROMPT)
            ElseIf lMin <> -1 And Len(sPAText) < lMin Then
                Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOSHORT_TEXT_CONSTANTPROMPT)
            End If
        End If
    End If

    Set oPA = Nothing

    ValidateConstantPromptAnswer = Err.Number
    Err.Clear
End Function

Function ValidateObjectPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML)
'******************************************************************************************
'Purpose:   validate object prompt answers - min/max
'Inputs:    aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML
'Outputs:   oSinglePromptAnswerXML
'******************************************************************************************
    On Error Resume Next
    Dim lMin
    Dim lMax
    Dim lSelections
    Dim bRequired
    Dim oPAIA

    lSelections = 0
    Set oPAIA = oSinglePromptAnswerXML.selectSingleNode("./pa[@ia='1']")
    If Not (oPAIA Is Nothing) Then
        lSelections = CLng(oPAIA.selectNodes("./mi/fct/*").length)
    End If

    bRequired = aPromptInfo(lPin, PROMPTINFO_B_REQUIRED)
    If bRequired And lSelections = 0 Then   'If no answer
        Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_REQUIRED_PROMPT)
    Else
        lMin = aPromptInfo(lPin, PROMPTINFO_L_MIN)
        lMax = aPromptInfo(lPin, PROMPTINFO_L_MAX)
        If lMax <> -1 And lSelections > lMax Then
            Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOMANY_SELECTIONS_OBJECTPROMPT)
        ElseIf lMin <> -1 And lSelections < lMin Then
            Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOFEW_SELECTIONS_OBJECTPROMPT)
        End If
    End If

    Set oPAIA = Nothing

    ValidateObjectPromptAnswer = Err.Number
    Err.Clear
End Function

Function ValidateLevelPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML)
'******************************************************************************************
'Purpose:   validate level prompt answers - min/max
'Inputs:    aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML
'Outputs:   oSinglePromptAnswerXML
'******************************************************************************************
    On Error Resume Next
    Dim lMin
    Dim lMax
    Dim lSelections
    Dim bRequired
    Dim oPAIA

    lSelections = 0
    Set oPAIA = oSinglePromptAnswerXML.selectSingleNode("./pa[@ia='1']")
    If Not (oPAIA Is Nothing) Then
        lSelections = CLng(oPAIA.selectNodes("./dmy/du").length)
    End If

    'Call CO_IsRequiredPrompt(oSinglePromptQuestionXML, bRequired)
    If bRequired And lSelections = 0 Then   'If no answer
        Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_REQUIRED_PROMPT)
    Else
        lMin = aPromptInfo(lPin, PROMPTINFO_L_MIN)
        lMax = aPromptInfo(lPin, PROMPTINFO_L_MAX)
        If lMax <> -1 And lSelections > lMax Then
            Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOMANY_SELECTIONS_LEVELPROMPT)
        ElseIf lMin <> -1 And lSelections < lMin Then
            Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOFEW_SELECTIONS_LEVELPROMPT)
        End If
    End If

    Set oPAIA = Nothing

    ValidateLevelPromptAnswer = Err.Number
    Err.Clear
End Function


Function ValidateElementPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML)
'*******************************************************************************************
'Purpose:   validate element prompt answers - min/max
'Inputs:    aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML
'Outputs:   oSinglePromptAnswerXML
'*******************************************************************************************
    On Error Resume Next
    Dim lMin
    Dim lMax
    Dim lSelections

    lSelections = CLng(oSinglePromptAnswerXML.selectNodes("./e").length)

    If aPromptInfo(lPin, PROMPTINFO_B_REQUIRED) And lSelections = 0 Then    'If no answer
        Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_REQUIRED_PROMPT)
    Else
        lMin = aPromptInfo(lPin, PROMPTINFO_L_MIN)
        lMax = aPromptInfo(lPin, PROMPTINFO_L_MAX)
        If lMax <> -1 And lSelections > lMax Then
            Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOMANY_SELECTIONS_ELEMENTPROMPT)
        ElseIf lMin <> -1 And lSelections < lMin Then
            Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOFEW_SELECTIONS_ELEMENTPROMPT)
        End If
    End If

    ValidateElementPromptAnswer = Err.Number
    Err.Clear
End Function

Function ValidateExpressionPromptAnswer(aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML, oRequest)
'********************************************************************************************************
'Purpose:   validate expression prompt answers
'Inputs:    aConnectionInfo, aPromptInfo, lPin, oSinglePromptQuestionXML, oSinglePromptAnswerXML, oRequest
'Outputs:   oSinglePromptAnswerXML
'********************************************************************************************************
    On Error Resume Next
    Dim oROOTOP
    Dim oROOTNode
    Dim bRequired
    Dim sFilterOP
    Dim lSelections
    Dim lMin
    Dim lMax
    Dim sDisplayUnknownDef

    Set oROOTNode = Nothing
    Set oROOTNode = oSinglePromptAnswerXML.selectSingleNode(".//pa[@ia='1']/exp/nd")

    bRequired = aPromptInfo(lPin, PROMPTINFO_B_REQUIRED)
    If bRequired And (oROOTNode Is Nothing) Then    'If no answer
        Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_REQUIRED_PROMPT)
    ElseIf Not (oROOTNode Is Nothing) Then
        Call CO_GetDisplayUnknownDef(oSinglePromptAnswerXML, sDisplayUnknownDef)
        If StrComp(sDisplayUnknownDef, "0") = 0 Then 'don't validate (Default)
            If oROOTNode.getAttribute("et") = CStr(ND_ExpressionType_BranchQual) Then
                Set oROOTOP = oROOTNode.selectSingleNode("./op")
                If oROOTOP Is Nothing And Err.Number = 0 Then
                    Err.Raise ERR_CUSTOM_NO_SPECIFIC_NODE
                    Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), "PromptValidateCuLib.asp", "ValidateExpressionPromptAnswer", "", "Error working with the XML", LogLevelError)
                Else
                    lSelections = oROOTNode.selectNodes("./nd").length
                    Call CO_GetFilterOperator(oSinglePromptAnswerXML, sFilterOP)
                    If sFilterOP = "OR" Then
                        Call oROOTOP.setAttribute("fnt", CStr(OP_FunctionType_Or))      'And all subexpressions
                    Else
                        Call oROOTOP.setAttribute("fnt", CStr(OP_FunctionType_And))     'OR all subexpressions
                    End If
                End If
            Else
                lSelections = 1
            End If
            If Err.Number = 0 Then
                lMin = aPromptInfo(lPin, PROMPTINFO_L_MIN)
                lMax = aPromptInfo(lPin, PROMPTINFO_L_MAX)
                If lMax <> -1 And lSelections > lMax Then
                    If aPromptInfo(lPin, PROMPTINFO_S_SUBTYPE) = EXPRESSIONTYPE_ALLATTRIBUTEQUAL Then
                        Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOMANY_SELECTIONS_HIPROMPT)
                    Else
                        Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOMANY_SELECTIONS_EXPPROMPT)
                    End If
                ElseIf lMin <> -1 And lSelections < lMin Then
                    If aPromptInfo(lPin, PROMPTINFO_S_SUBTYPE) = EXPRESSIONTYPE_ALLATTRIBUTEQUAL Then
                        Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOFEW_SELECTIONS_HIPROMPT)
                    Else
                        Call CO_SetPromptError(oSinglePromptAnswerXML, ERR_TOOFEW_SELECTIONS_EXPPROMPT)
                    End If
                End If
            Else
                Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), "PromptValidateCuLib.asp", "ValidateExpressionPromptAnswer", "", "Error setting filter operator", LogLevelError)
            End If
        End If
    End If

    Set oROOTOP = Nothing
    Set oROOTNode = Nothing

    ValidateExpressionPromptAnswer = Err.Number
    Err.Clear
End Function
%>
