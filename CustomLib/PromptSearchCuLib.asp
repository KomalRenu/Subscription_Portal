<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function BuildFilterXMLforDrillinHIPrompt(aConnectionInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest, oExpression)
'***************************************************************************************************
'Purpose:   Create a FilterXML for Drill, ex Year in (2000, 2001)
'Input:     aConnectionInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest,
'Output:    oFilterXML
'***************************************************************************************************
    On Error Resume Next
    Dim sATName
    Dim sATDID
    Dim sFilterHier
    Dim oOperatorNode
    Dim oElementList
    Dim oElements
    Dim sSelection
    Dim temArray
    Dim oRootXML
    Dim sElemID
    Dim aAvailable

    Set oExpression = Nothing

	'aAvailable = Split(oRequest("available_" & sPin), ", ", -1, vbBinaryCompare)
	aAvailable = SplitRequest(oRequest("available_" & sPin))

    If UBound(aAvailable) > -1 Then
        Call CO_GetAttBeforeDrillforHIPrompt(oSinglePromptTempXML, sATName, sATDID)

        Set oExpression = Server.CreateObject("WebAPIHelper.DSSXMLExpression")

        sFilterHier = CStr(oRequest("nuXML_filterHier_" & sPin))

		If Len(sFilterHier) = 0 Or IsEmpty(sFilterHier) Then
			oExpression.RootNode.Operator = DssXmlFunctionAnd
		Else
			Call oExpression.LoadFromXML(sFilterHier)
		End If

        Set oOperatorNode = oExpression.CreateOperatorNode(DssXmlFilterListQual, DssXmlFunctionIn)
        Call oExpression.CreateShortCutNode(sATDID, DssXmlTypeAttribute, oOperatorNode)
        Set oElementList = oExpression.CreateElementListNode(sATDID, oOperatorNode)
        Set oElements = oElementList.ElementsObject

        For Each sSelection In aAvailable
            If StrComp(CStr(sSelection), "-none-", vbTextCompare) = 0 Then
                Exit For
            End If

            sElemID = Left(sSelection, InStr(1, sSelection, chr(30), vbBinaryCompare))

            If InStr(1, sElemID, chr(30), vbBinaryCompare) > 0 Then
				sElemID = Left(sElemID, Len(sElemID)-1)
            End If

            oElements.Add (sElemID)
        Next
    End If

    BuildFilterXMLforDrillinHIPrompt = Err.Number
    Err.Clear
End Function


Function BuildFilterXMLforDrillupHIPrompt(aConnectionInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, sElemID, oExpression, oRequest)
'***************************************************************************************************
'Purpose:   Create a FilterXML for Drill up, ex Year in (2000, 2001)
'Input:     aConnectionInfo, sPin, oSinglePromptQuestionXML, oSinglePromptTempXML, oRequest,
'Output:    oFilterXML
'***************************************************************************************************
    On Error Resume Next
    Dim sATName
    Dim sATDID
    Dim oOperatorNode
    Dim oElementList
    Dim oElements
    Dim sSelection
    Dim temArray
    Dim oRootXML
    Dim sFilterHier
    Dim oFilterHier
    Dim sErrDescription
    Dim lATDispID

    Set oExpression = Nothing

    Call CO_GetAttBeforeDrillforHIPrompt(oSinglePromptTempXML, sATName, sATDID)

    Set oExpression = Server.CreateObject("WebAPIHelper.DSSXMLExpression")

    sFilterHier = CStr(oRequest("nuXML_filterHier_" & sPin))

	If Len(sFilterHier) = 0 Or IsEmpty(sFilterHier) Then
		oExpression.RootNode.Operator = DssXmlFunctionAnd
	Else
		Call oExpression.LoadFromXML(sFilterHier)
	End If

	Call GetXMLDOM(aConnectionInfo, oFilterHier, sErrDescription)

	If Err.number = NO_ERR Then
		Call oFilterHier.loadXML(sFilterHier)
		If Not oFilterHier.selectSingleNode("./f/mi/exp/nd/op[@fnt='19']") Is Nothing Then
			If Not oFilterHier.selectSingleNode("./f/mi/exp/nd/nd[./nd[@et='1' $and$ @nt='2']/mi/in/oi[@did='" & sElemID & "']]") Is Nothing Then
				lATDispID = CLng(oFilterHier.selectSingleNode("./f/mi/exp/nd/nd[./nd[@et='1' $and$ @nt='2']/mi/in/oi[@did='" & sElemID & "']]/@disp_id").text)
				Call oExpression.RootNode.RemoveChild(oExpression.FindNodeByDisplayID(lATDispID))
			End If
		Else
			If Not oFilterHier.selectSingleNode("./f/mi/exp/nd[./nd[@et='1' $and$ @nt='2']/mi/in/oi[@did='" & sElemID & "']]") Is Nothing Then
				lATDispID = CLng(oFilterHier.selectSingleNode("./f/mi/exp/nd[./nd[@et='1' $and$ @nt='2']/mi/in/oi[@did='" & sElemID & "']]/@disp_id").text)
				Call oExpression.RootNode.RemoveChild(oExpression.FindNodeByDisplayID(lATDispID))
			End If
		End If
	End If

    BuildFilterXMLforDrillupHIPrompt = Err.Number
    Err.Clear
End Function



Function CO_BuildFilterXMLForSearchField(aConnectionInfo, oSession, oSinglePrompt, sSearch, oSinglePromptQuestionXML, oSinglePromptTempXML, sFlag, oFilterExpression)
'***************************************************************************************************
'Purpose:  Build oFilterXML part for search field for element prompt
'Input:     aConnectionInfo, sSearch, oSinglePromptQuestionXML, oSinglePromptTempXML, sFlag
'Output:    oFilterXML
'***************************************************************************************************
    On Error Resume Next
    Dim lResult
    Dim sLogicalStart
    Dim sLogicalEnd
    Dim sLogicalPos1
    Dim sLogicalPos2
    Dim sLogicalPos3
    Dim sLogicalPos4
    Dim sLogicalPos5
    Dim sRest
    Dim sLogical
    Dim sSegment
    Dim oCurND
    Dim oNDRoot
    Dim oEXP
    Dim oRootOP
    Dim oNewRoot
    Dim bFirst
    Dim sPosEndDoubleQuote
    Dim bErrorLog
    Dim oElementSource
    Dim sATName
    Dim sATDID
    Dim sAttributeXML
    Dim oAttributeXML

    bErrorLog = True

    If oFilterExpression Is Nothing Or IsNull(oFilterExpression) Or IsEmpty(oFilterExpression) Then
		Set oFilterExpression = Server.CreateObject("WebAPIHelper.DSSXMLExpression")
		oFilterExpression.RootNode.Operator = DssXmlFunctionOr
	End If

    Select Case sFlag
		Case SEARCHFIELD_ELEPROMPT
		    Set oElementSource = oSinglePrompt.ElementSourceObject

		    If oSinglePrompt.PromptType = DssXmlPromptElements And Len(oSinglePrompt.ElementSourceObject.Filter) = 0 Then
				Call oElementSource.ExpressionObject.RootNode.Clear()
				oElementSource.ExpressionObject.RootNode.Operator = DssXmlFunctionOr
			End If

		    sATDID = oElementSource.AttributeID
		Case SEARCHFIELD_HIPROMPT
		    Call CO_GetAttributeforHIPrompt(oSinglePromptTempXML, sATName, sATDID)
		Case SEARCHFIELD_HIPROMPT_BEFOREDRILL
		    Call CO_GetAttBeforeDrillforHIPrompt(oSinglePromptTempXML, sATName, sATDID)
    End Select

    Set oObjServer = oSession.ObjectServer
    sAttributeXML = oObjServer.FindObject(aConnectionInfo(S_TOKEN_CONNECTION), sATDID, DssXmlTypeAttribute, 127, 0, 1, 0)
    Call GetXMLDOM(aConnectionInfo, oAttributeXML, sErrDescription)

    Call oAttributeXML.loadXML(sAttributeXML)

    If Err.Number = 0 Then
        sSearch = Trim(sSearch)
        bFirst = True
        sRest = sSearch
    End If

    'segment by ,  Or  |  And
    While Len(sRest) > 0
        If StrComp(Left(sRest, 1), """", vbTextCompare) = 0 Then
            sPosEndDoubleQuote = InStr(2, sRest, """", vbBinaryCompare)
            sSegment = Left(sRest, sPosEndDoubleQuote)
            sRest = Right(sRest, Len(sRest) - sPosEndDoubleQuote)
        Else
			If InStr(1, sRest, ":", vbBinaryCompare) = 0 Then
				sLogicalPos1 = InStr(1, sRest, " ", vbBinaryCompare)
			Else
				sLogicalPos1 = 0
			End If
            sLogicalPos2 = InStr(1, sRest, ",", vbBinaryCompare)
            sLogicalPos3 = InStr(1, sRest, " or ", vbTextCompare)
            If sLogicalPos3 > 0 Then sLogicalPos3 = sLogicalPos3 + 1
            sLogicalPos4 = InStr(1, sRest, "|", vbBinaryCompare)
            sLogicalPos5 = InStr(1, sRest, " and ", vbTextCompare)
            If sLogicalPos5 > 0 Then sLogicalPos5 = sLogicalPos3 + 1

            sLogicalStart = 0
            sLogicalEnd = 0
            If sLogicalPos2 > 0 Then
                sLogicalStart = sLogicalPos2
                sLogicalEnd = sLogicalStart
                sLogical = CStr(DssXmlFunctionOr)
            ElseIf sLogicalPos3 > 0 Then
                sLogicalStart = sLogicalPos3
                sLogicalEnd = sLogicalStart + 1
                sLogical = CStr(DssXmlFunctionOr)
            ElseIf sLogicalPos4 > 0 Then
                sLogicalStart = sLogicalPos4
                sLogicalEnd = sLogicalStart
                sLogical = CStr(DssXmlFunctionAnd)
            ElseIf sLogicalPos5 > 0 Then
                sLogicalStart = sLogicalPos5
                sLogicalEnd = sLogicalStart + 2
                sLogical = CStr(DssXmlFunctionAnd)
            ElseIf sLogicalPos1 > 0 Then        'ID:00 And J*
                sLogicalStart = sLogicalPos1
                sLogicalEnd = sLogicalStart
                sLogical = CStr(DssXmlFunctionOr)
            End If

            If sLogicalStart > 0 Then
                sSegment = Left(sRest, sLogicalStart - 1)
                sRest = Right(sRest, Len(sRest) - sLogicalEnd)
            Else
                sSegment = sRest
                sRest = ""
            End If
        End If

        sSegment = Trim(sSegment)
        sRest = Trim(sRest)

        If Err.Number = 0 Then
            lResult = CO_BuildSegmentOfFilterXMLForSearchField(aConnectionInfo, sSegment, oAttributeXML, oSinglePromptTempXML, oFilterExpression)
            If lResult <> 0 Then
                Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptSearchCoLib.asp", "CO_BuildFilterXMLForSearchField", "CO_BuildSegmentOfFilterXMLForSearchField", "Error in call to CO_BuildSegmentOfFilterXMLForSearchField", LogLevelTrace)
                Err.Raise lResult
            End If
        ElseIf bErrorLog Then
            Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.Source, "PromptSearchCoLib.asp", "CO_BuildFilterXMLForSearchField", "", "Error working with XML", LogLevelError)
        End If
    Wend

    Set oCurND = Nothing
    Set oNDRoot = Nothing
    Set oEXP = Nothing
    Set oRootOP = Nothing
    Set oNewRoot = Nothing

    CO_BuildFilterXMLForSearchField = Err.Number
    Err.Clear
End Function


Function CO_BuildSegmentOfFilterXMLForSearchField(aConnectionInfo, sSegment, oAttributeXML, oSinglePromptTempXML, oFilterExpression)
'***************************************************************************************************
'Purpose:  Build a sub Expression of oFilterXML for search field for element prompt
'Input:     aConnectionInfo, sSegment, oSinglePromptQuestionXML, oSinglePromptTempXML, sFlag
'Output:    oFilterXML, oResultND
'***************************************************************************************************
    Dim sFormName
    Dim sFormPos
    Dim lResult
    Dim oCurND
    Dim oNDRoot
    Dim oNewRoot
    Dim oRootOP
    Dim oFMOI
    Dim oBFFUS
    Dim oFU
    Dim sFUrfd
    Dim sATName
    Dim oATOI
    Dim lBFCount
    Dim oResultNDParent
    Dim oIN
    Dim oNode
    Dim oSearchNode

    sFormName = ""
    sFormPos = InStr(1, sSegment, ":", vbBinaryCompare)
    If sFormPos > 0 Then
        sFormName = Trim(Left(sSegment, sFormPos - 1))
        sSegment = Trim(Right(sSegment, Len(sSegment) - sFormPos))
    End If

	If oFilterExpression.RootNode.ChildCount > 0 Then
		Set oNDRoot = oFilterExpression.CreateOperatorNode(DssXmlFilterBranchQual,DssXmlFunctionOr)
	Else
		Set oNDRoot = oFilterExpression.RootNode
	End If

    Set oSearchNode = oFilterExpression.CreateOperatorNode(DssXmlFilterBranchQual, DssXmlFunctionOr, oNDRoot)


    If Len(sFormName) > 0 Then
        lResult = CO_BuildBasicExpofFilterXMLForSearchField(aConnectionInfo, sFormName, oAttributeXML, sSegment, oSinglePromptTempXML, oFilterExpression, oSearchNode)
        If lResult <> 0 Then
            Set oBFFUS = oAttributeXML.selectNodes("./mi/fi/bfs/fu")

			For Each oFU In oBFFUS
	            sFUrfd = oFU.getAttribute("rfd")
		        Set oFMOI = oAttributeXML.selectSingleNode("./mi/in/oi[@id = '"&sFUrfd&"']" )
			    sFormName = oFMOI.getAttribute("n")

				If Err.Number = 0 Then
	                lResult = CO_BuildBasicExpofFilterXMLForSearchField(aConnectionInfo, sFormName, oAttributeXML, sSegment, oSinglePromptTempXML, oFilterExpression, oSearchNode)
		            If lResult <> 0 Then
			            Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptSearchCoLib.asp", "CO_BuildSegmentOfFilterXMLForSearchField", "", "Error in call to CO_BuildBasicExpofFilterXMLForSearchField", LogLevelTrace)
				        Err.Raise lResult
					End If
				Else
					Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.Source, "PromptSearchCoLib.asp", "CO_BuildSegmentOfFilterXMLForSearchField", "oATOI.parentNode.selectSingleNode", "Error working with XML", LogLevelError)
				End If
	        Next
        End If
    Else    'Form1 or sSegment or Form1 OP sSegment
        Set oBFFUS = oAttributeXML.selectNodes("./mi/fi/bfs/fu")

        For Each oFU In oBFFUS
            sFUrfd = oFU.getAttribute("rfd")
            Set oFMOI = oAttributeXML.selectSingleNode("./mi/in/oi[@id = '"&sFUrfd&"']" )
            sFormName = oFMOI.getAttribute("n")

            If Err.Number = 0 Then
                lResult = CO_BuildBasicExpofFilterXMLForSearchField(aConnectionInfo, sFormName, oAttributeXML, sSegment, oSinglePromptTempXML, oFilterExpression, oSearchNode)
                If lResult <> 0 Then
                    Call LogErrorXML(aConnectionInfo, CStr(lResult), CStr(Err.Description), Err.Source, "PromptSearchCoLib.asp", "CO_BuildSegmentOfFilterXMLForSearchField", "", "Error in call to CO_BuildBasicExpofFilterXMLForSearchField", LogLevelTrace)
                    Err.Raise lResult
                End If
            Else
                Call LogErrorXML(aConnectionInfo, CStr(Err.Number), CStr(Err.Description), Err.Source, "PromptSearchCoLib.asp", "CO_BuildSegmentOfFilterXMLForSearchField", "oATOI.parentNode.selectSingleNode", "Error working with XML", LogLevelError)
            End If

        Next
    End If

    CO_BuildSegmentOfFilterXMLForSearchField = Err.Number
    Err.Clear
End Function


Function CO_BuildBasicExpofFilterXMLForSearchField(aConnectionInfo, sFormName, oAttributeXML, sSegment, oSinglePromptTempXML, oFilterExpression, oNDRoot)
'***************************************************************************************************
'Purpose:  Build a sub Expression of oFilterXML for search field for element prompt
'Input:     aConnectionInfo, oFMOI, sSegment, oSinglePromptQuestionXML, oSinglePromptTempXML, sFlag
'Output:    oFilterXML, oCurExp
'***************************************************************************************************
    On Error Resume Next
    Dim sATDID
    Dim sFMDID
    Dim oFM
    Dim sDT
    Dim sDDT
    Dim sCurSegment
    Dim sMatchCase
    Dim oLHSOperatorNodeUCase
    Dim oRHSOperatorNodeUCase
    Dim oOperatorNodeLike
    Dim lOperator
    Dim lConstType

    sCurSegment = sSegment

    sATDID = oAttributeXML.selectSingleNode("./mi/in/oi[@tp='12']").getAttribute("did")
    Set oFM = oAttributeXML.selectSingleNode("./mi/in/oi[@tp='21' $and$ ./fdt/text()='" & sFormName & "']")

    If oFM Is Nothing Then
		Set oFM = oAttributeXML.selectSingleNode("./mi/in/oi[@tp='21' $and$ @n='" & sFormName & "']")
    End If

    sFMDID = oFM.getAttribute("did")

    'If Not oFM Is Nothing Then
    If Err.Number = NO_ERR Then
		If Not oFM.selectSingleNode("fdt/@dt") Is Nothing Then
			sDT = oFM.selectSingleNode("fdt/@dt").text
		Else
			sDT = "3"
		End If

		Call MapDTtoDDT(sDT, sDDT)

		lOperator = DssXmlFunctionEquals
		lConstType = sDDT

		If CO_ValidateSegmentDataType(sCurSegment, lConstType) Then
			If Int(sDDT) = DssXmlDataTypeChar Or Int(sDDT) = DssXmlDataTypeVarChar Then
				If StrComp(Left(sCurSegment, 1), """", vbBinaryCompare) = 0 Then
				    sCurSegment = Right(sCurSegment, Len(sCurSegment) - 1)
				    If StrComp(Right(sCurSegment, 1), """", vbBinaryCompare) = 0 Then
						sCurSegment = Left(sCurSegment, Len(sCurSegment) - 1)
					End If
				Else
					If InStr(1, sCurSegment, "*", vbBinaryCompare) = 0 And InStr(1, sCurSegment, "?", vbBinaryCompare) = 0 _
						And (Int(sDDT) = DssXmlDataTypeChar Or Int(sDDT) = DssXmlDataTypeVarChar) Then
						sCurSegment = "*" & sCurSegment & "*"
					End If
					lOperator = DssXmlFunctionLike
				End If
			End If

			Call CO_GetMatchCase(oSinglePromptTempXML, sMatchCase)
			If StrComp(sMatchCase, "0") = 0 And (Int(sDDT) = DssXmlDataTypeChar Or Int(sDDT) = DssXmlDataTypeVarChar) Then
				Set oOperatorNodeLike = oFilterExpression.CreateOperatorNode(DssXmlFilterSingleBaseFormQual, lOperator, oNDRoot)
				Set oLHSOperatorNodeUCase = oFilterExpression.CreateOperatorNode(DssXmlFilterSingleBaseFormQual, DssXmlFunctionUcase, oOperatorNodeLike)
				Call oFilterExpression.CreateFormShortCutNode(sATDID, sFMDID, oLHSOperatorNodeUCase)
				Call oFilterExpression.CreateConstantNode(UCase(sCurSegment), lConstType, oOperatorNodeLike)
			Else
				Set oOperatorNodeLike = oFilterExpression.CreateOperatorNode(DssXmlFilterSingleBaseFormQual, lOperator, oNDRoot)
				Call oFilterExpression.CreateFormShortCutNode(sATDID, sFMDID, oOperatorNodeLike)
				Call oFilterExpression.CreateConstantNode(sCurSegment, lConstType, oOperatorNodeLike)
			End If
		End If
	End If

    CO_BuildBasicExpofFilterXMLForSearchField = Err.Number
    Err.Clear
End Function

Function CO_ValidateSegmentDataType(sCurSegment, lConstType)
'***************************************************************************************************
'Purpose:  Validate that segment value data type is correct
'Input:     sCurSegment, lConstType
'Output:    Boolean value indicating that value has correct data type or not
'***************************************************************************************************
	Dim bValid

	bValid = False

	Select Case CLng(lConstType)
		Case DssXmlDataTypeReal
		    If IsNumeric(sCurSegment) Then
				bValid = True
		    End If
		Case DssXmlDataTypeDate, DssXmlDataTypeTime, DssXmlDataTypeTimeStamp
			If IsValidDate(sCurSegment) Then
		    'If IsDate(sCurSegment) Then
				bValid = True
		    End If
		Case DssXmlDataTypeChar, DssXmlDataTypeVarChar
			bValid = True
    End Select

	CO_ValidateSegmentDataType = bValid
End Function

%>
