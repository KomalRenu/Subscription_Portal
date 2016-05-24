<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'

Dim aFilterProperties
Redim aFilterProperties(9)
Public Const VERT_PIXELS_PER_CHAR = 10
Public Const HORI_PIXELS_PER_CHAR = 10

Private Const TABS				= 0
Private Const CARRIAGE_RETURN	= 1
Private Const CROP_FILTER		= 2
Private Const MAXIMUM_SIZE		= 3
Private Const NUMBER_OF_LINES	= 4
Private Const FILTER_CONTENTS	= 5
Private Const FOR_EXPORT		= 6
Private Const FILTER_WIDTH		= 7
Private Const FILTER_HEIGHT		= 8
Private Const REPORT_FILTER_CONTENTS = 9
'*** CORE FUNCTIONS ***'
Function GetFilterDetails(aConnectionInfo, oReportXML, aFilterProperties, sErrDescription)
'******************************************************************************
'Purpose: Show the filter details in HTML, or in plain text
'Inputs:  aConnectionInfo, oReportXML, aFilterProperties
'Outputs: aFilterProperties, sErrDescription
'******************************************************************************
	On Error Resume Next
	Dim sFilterText
	Dim sMetricLimit
	Dim lErrNumber
	Dim bEmptyReportFilter
	Dim bEmptyViewFilter

	bEmptyReportFilter = False
	bEmptyViewFilter = False

	' Report filter
	If Not oReportXML.SelectSingleNode("/mi/rit/f/mi/exp") Is Nothing Then
		If oReportXML.SelectSingleNode("/mi/rit/f/mi/exp").hasChildNodes Then
			aFilterProperties(FILTER_CONTENTS) = oReportXML.SelectSingleNode("/mi/rit/f/mi/exp").firstChild.text
		End If
	ElseIf Not oReportXML.SelectSingleNode("/mi/rit/working_set/f/mi/exp") Is Nothing Then ' filter is here if the flag: DSSXMLResultWorkingset was used
		If oReportXML.SelectSingleNode("/mi/rit/working_set/f/mi/exp").hasChildNodes Then
			aFilterProperties(FILTER_CONTENTS) = oReportXML.SelectSingleNode("/mi/rit/working_set/f/mi/exp").firstChild.text
		End If
	ElseIf Not oReportXML.SelectSingleNode("/mi/rit/rdt/@fex") Is Nothing Then
		aFilterProperties(FILTER_CONTENTS) = oReportXML.SelectSingleNode("/mi/rit/rdt/@fex").text
	End If

	lErrNumber = Err.number
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, Cstr(lErrNumber), Err.description, Err.source, "FilterDetailsCuLib.asp", "GetFilterDetails", "aFilterProperties(FILTER_CONTENTS) = oReportXML.SelectSingleNode(""/mi/rit/working_set/f/mi/exp"").firstChild.text", "Error loading XML", LogLevelError)
	End If

	If Not oReportXML.SelectSingleNode("/mi/rit/rdt/@tml") Is Nothing Then
		sMetricLimit = oReportXML.SelectSingleNode("/mi/rit/rdt/@tml").nodeValue
		lErrNumber = Err.number
		If lErrNumber <> NO_ERR Then
			sErrDescription = asDescriptors(801) 'Descriptor: Error loading XML data.
			Call LogErrorXML(aConnectionInfo, Cstr(lErrNumber), sErrdescription, Err.source, "FilterDetailsCuLib.asp", "GetFilterDetails", "sMetricLimit = oReportXML.SelectSingleNode(""/mi/rit/rdt/@tml"").nodeValue", "Error loading XML", LogLevelError)
		Else
			If (Len(sMetricLimit) > 0) And (Not IsEmpty(sMetricLimit)) Then
				aFilterProperties(FILTER_CONTENTS) = asDescriptors(178) & aFilterProperties(CARRIAGE_RETURN) & aFilterProperties(FILTER_CONTENTS) & aFilterProperties(CARRIAGE_RETURN) & asDescriptors(1121) & aFilterProperties(CARRIAGE_RETURN) 'Descriptors: Filter details: | Report limit:
				lErrNumber = ParseFilterDetails(aConnectionInfo, sMetricLimit, aFilterProperties, sErrDescription)
				aFilterProperties(FILTER_CONTENTS) = aFilterProperties(FILTER_CONTENTS) & sMetricLimit
			End If
		End If
	End If

	' View filter
	If Not oReportXML.SelectSingleNode("/mi/rit/view_report/view_filter/f/mi/exp") Is Nothing Then
		aFilterProperties(REPORT_FILTER_CONTENTS) = oReportXML.SelectSingleNode("/mi/rit/view_report/view_filter/f/mi/exp").firstChild.text
		lErrNumber = Err.number
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, Cstr(lErrNuber), Err.description, Err.source, "FilterDetailsCuLib.asp", "GetFilterDetails", "aFilterProperties(REPORT_FILTER_CONTENTS) = oReportXML.SelectSingleNode(""/mi/rit/view_report/view_filter/f/mi/exp"").firstChild.text", "Error loading XML", LogLevelError)
		End If
	End If

	'Now choose what to display
	If (Len(aFilterProperties(FILTER_CONTENTS)) = 0) Or (IsEmpty(aFilterProperties(FILTER_CONTENTS))) Then
		bEmptyReportFilter = True
		'aFilterProperties(FILTER_CONTENTS) = asDescriptors(179) 'Descriptor: The filter is empty
	Else
		lErrNumber = ParseFilterDetails(aConnectionInfo, aFilterProperties(FILTER_CONTENTS), aFilterProperties, sErrDescription)
	End If

	If (Len(aFilterProperties(REPORT_FILTER_CONTENTS)) = 0) Or (IsEmpty(aFilterProperties(REPORT_FILTER_CONTENTS))) Then
		bEmptyViewFilter = True
		'aFilterProperties(REPORT_FILTER_CONTENTS) = asDescriptors(179) 'Descriptor: The filter is empty
	Else
		lErrNumber = ParseFilterDetails(aConnectionInfo, aFilterProperties(REPORT_FILTER_CONTENTS), aFilterProperties, sErrDescription)
	End If

	If bEmptyReportFilter And bEmptyViewFilter Then
		aFilterProperties(FILTER_CONTENTS) = asDescriptors(179) 'Descriptor: The filter is empty
		aFilterProperties(REPORT_FILTER_CONTENTS) = ""
	Else
		If bEmptyReportFilter Then
			aFilterProperties(FILTER_CONTENTS) = ""'asDescriptors(179) 'Descriptor: The filter is empty
		ElseIf bEmptyViewFilter Then
			aFilterProperties(REPORT_FILTER_CONTENTS) = "" 'asDescriptors(179) 'Descriptor: The filter is empty
		End If
	End If

	Err.Clear
End Function

Function ParseFilterDetails(aConnectionInfo, sFilter, aFilterProperties, sErrDescription)
'******************************************************************************
'Purpose: Show the filter details in HTML, or in plain text
'Inputs:  aConnectionInfo, sFilter, aFilterProperties
'Outputs: sFilter, aFilterProperties, sErrDescription
'******************************************************************************
	On Error Resume Next
	Dim sFilterDetailsSize
	Dim sIndent
	Dim i
	Dim bInsideName
	Dim bBetween
	Dim bSpace
	Dim lErrNumber

	Dim Bracket
	Dim bracketlevel
	Dim stemp
	Dim sCurrent
	Dim sOldFilter

	Bracket = -1
	bracketlevel = 0
	lErrNumber = NO_ERR
	sIndent = ""
	aFilterProperties(NUMBER_OF_LINES) = 1
	bInsideName = False
	bBetween = False
	sOldFilter = sFilter
	sFilter = ""

	If StrComp(Left(sOldFilter, 1), "(", vbBinaryCompare) = 0 And StrComp(Right(sOldFilter, 1), ")", vbBinaryCompare) = 0 Then
		sOldFilter = Mid(sOldFilter, 2, len(sOldFilter) - 2)
	End If

	sOldFilter = CleanString(sOldFilter)
	Dim bNeedTab, Pars, bNeedCR
	bNeedTab = False
	bNeedCR = False
	bSpace = False
	Pars = 0
	For i = 1 To Len(sOldFilter)
		sCurrent = Mid(sOldFilter,i,1)
		If StrComp(sCurrent, "{", vbBinaryCompare) = 0 Then
			If aFilterProperties(FOR_EXPORT) And (StrComp(sCurrent, "=", vbBinaryCompare) = 0 Or StrComp(sCurrent, "+", vbBinaryCompare) = 0 Or StrComp(sCurrent, "-", vbBinaryCompare) = 0) Then
				sFilter = sFilter & " "
			End If
			bInsideName = True
			bNeedTab = False
			bSpace = False
		ElseIf StrComp(sCurrent, "}", vbBinaryCompare) = 0 Then
			bInsideName = False
			bNeedTab = False
			bSpace = False
		ElseIf StrComp(sCurrent, "(", vbBinaryCompare) = 0 Then
			If bInsideName Then
				sFilter = sFilter & "("
			ElseIf bNeedTab or i = 1 Then
				sFilter = sFilter & aFilterProperties(CARRIAGE_RETURN) & sIndent & aFilterProperties(TABS)
				sIndent = sIndent & aFilterProperties(TABS)
				bNeedTab = True
			Else
				sFilter = sFilter & "("
				Pars = Pars + 1
			End If
			bSpace = False
		ElseIf StrComp(sCurrent, ")", vbBinaryCompare) = 0 Then
			bNeedTab = False
			If bInsideName Then
				sFilter = sFilter & ")"
			ElseIf  Pars > 0 Then
				sFilter = sFilter & ")"
				Pars = Pars - 1
			Else
				sIndent = Left(sIndent,Len(sIndent) - Len(aFilterProperties(TABS)))
				sFilter = sFilter & aFilterProperties(CARRIAGE_RETURN) & sIndent
				aFilterProperties(NUMBER_OF_LINES) = aFilterProperties(NUMBER_OF_LINES) + 1
				bNeedTab = True
				bNeedCR = 1
			End If
			bSpace = False
		ElseIf StrComp(sCurrent, ",", vbBinaryCompare) = 0 Then
			bNeedTab = False
			If bInsideName Then
				sFilter = sFilter & ","
			Else
				sFilter = sFilter & ", "
			End If
			bSpace = False
		ElseIf StrComp(sCurrent, " ", vbBinaryCompare) = 0 Then
			If bNeedCR = 2 Then
				bNeedCR = 0
				bNeedTab = True
			ElseIf bNeedCR > 0 Then
				bNeedCR = bNeedCR + 1
			Else
				sFilter = sFilter & " "
			End If
			bSpace = True
		Else
			bNeedTab = False
			If bSpace Then
				sFilter = sFilter & " "
			End If
			sFilter = sFilter & Mid(sOldFilter, i, 1)
			bSpace = False
		End If

		If aFilterProperties(CROP_FILTER) And (aFilterProperties(NUMBER_OF_LINES) > aFilterProperties(MAXIMUM_SIZE)) And (aFilterProperties(MAXIMUM_SIZE) <> -1) Then
			sFilter = sFilter & "..."
			Exit For
		End If
	Next

	If Len(sFilter) > 0 Then
		Call CleanFilterDetails(sFilter, aFilterProperties)
	Else
		sFilter = asDescriptors(179) 'Descriptor: The filter is empty
	End If

	lErrNumber = Err.number
	If lErrNumber <> NO_ERR Then
		sErrDescription = Err.description
		Call LogErrorXML(aConnectionInfo, Cstr(lErrNumber), sErrdescription, Err.source, "FilterDetailsCuLib.asp", "ParseFilterDetails", "", "Error on ParseFilterDetails", LogLevelError)
	End If
	ParseFilterDetails = lErrNumber
	Err.Clear
End Function

Function CleanFilterDetails(sFilter, aFilterProperties)
'******************************************************************************
'Purpose: Get rid of extra blank spaces, tabs, or carriage returns inside the filter details text
'Inputs:  sFilter, aFilterProperties
'Outputs: sFilter, aFilterProperties
'******************************************************************************
	On Error Resume Next
	Dim i, j, iTemp
	Dim sRows
	Do While InStr(1, sFilter, aFilterProperties(TABS) & aFilterProperties(CARRIAGE_RETURN), vbBinaryCompare) > 0
		sFilter = Replace(sFilter, aFilterProperties(TABS) & aFilterProperties(CARRIAGE_RETURN), aFilterProperties(CARRIAGE_RETURN))
	Loop
	Do While InStr(1, sFilter, aFilterProperties(CARRIAGE_RETURN) & " ", vbBinaryCompare) > 0
		sFilter = Replace(sFilter, aFilterProperties(CARRIAGE_RETURN) & " ", aFilterProperties(CARRIAGE_RETURN))
	Loop
	Do While InStr(1, sFilter, " " & aFilterProperties(CARRIAGE_RETURN), vbBinaryCompare) > 0
		sFilter = Replace(sFilter, " " & aFilterProperties(CARRIAGE_RETURN), aFilterProperties(CARRIAGE_RETURN))
	Loop
	Do While InStr(1, sFilter, aFilterProperties(CARRIAGE_RETURN) & aFilterProperties(CARRIAGE_RETURN), vbBinaryCompare) > 0
		sFilter = Replace(sFilter, aFilterProperties(CARRIAGE_RETURN) & aFilterProperties(CARRIAGE_RETURN), aFilterProperties(CARRIAGE_RETURN))
		aFilterProperties(NUMBER_OF_LINES) = aFilterProperties(NUMBER_OF_LINES) - 1
	Loop
	If Len(sFilter) > 0 Then
		Do While StrComp(Left(sFilter, Len(aFilterProperties(CARRIAGE_RETURN))), aFilterProperties(CARRIAGE_RETURN), vbBinaryCompare) = 0
			sFilter = Mid(sFilter, Len(aFilterProperties(CARRIAGE_RETURN)) + 1, Len(sFilter) - Len(aFilterProperties(CARRIAGE_RETURN)))
			aFilterProperties(NUMBER_OF_LINES) = aFilterProperties(NUMBER_OF_LINES) - 1
		Loop
	End If
	If Len(sFilter) > 0 Then
		Do While StrComp(Right(sFilter, Len(aFilterProperties(CARRIAGE_RETURN))), aFilterProperties(CARRIAGE_RETURN), vbBinaryCompare) = 0
			sFilter = Left(sFilter, Len(sFilter) - Len(aFilterProperties(CARRIAGE_RETURN)))
			aFilterProperties(NUMBER_OF_LINES) = aFilterProperties(NUMBER_OF_LINES) - 1
		Loop
	End If
	aFilterProperties(NUMBER_OF_LINES) = 0
	If Len(sFilter) > 0 Then
		If InStr(1, sFilter, aFilterProperties(CARRIAGE_RETURN), vbTextCompare) > 0 Then
			sRows = Split(sFilter, aFilterProperties(CARRIAGE_RETURN))
			aFilterProperties(NUMBER_OF_LINES) = UBound(sRows) + 1
		End If
	End If
	i = 0
	j = 0
	sFilter = Replace(sFilter, "&nbsp;", Chr(160))
	If Len(sFilter) > 0 Then
		i = InStr(i + 1, sFilter, aFilterProperties(CARRIAGE_RETURN), vbBinaryCompare)
		If i = 0 Then
			i = Len(sFilter)
		End If
		Do While (i - j) > 0
			If (i - j) >= aFilterProperties(FILTER_WIDTH) Then
				If aFilterProperties(FOR_EXPORT) Then
					iTemp = j
					Do While iTemp + aFilterProperties(FILTER_WIDTH) < i
						sFilter = Mid(sFilter, 1, iTemp +  aFilterProperties(FILTER_WIDTH)) & _
						  aFilterProperties(CARRIAGE_RETURN) & Mid(sFilter, iTemp + aFilterProperties(FILTER_WIDTH) + 1)
						iTemp = iTemp + aFilterProperties(FILTER_WIDTH)
					Loop

				End If
				aFilterProperties(NUMBER_OF_LINES) = aFilterProperties(NUMBER_OF_LINES) + Int((i - j) / aFilterProperties(FILTER_WIDTH))
			End If
			j = i
			i = InStr(i + 1, sFilter, aFilterProperties(CARRIAGE_RETURN), vbBinaryCompare)
			If i = 0 Then
				i = Len(sFilter)
			End If
		Loop
	End If

	lErrNumber = Err.number
	If lErrNumber <> NO_ERR Then
		sErrDescription = Err.description
		Call LogErrorXML(aConnectionInfo, Cstr(lErrNumber), sErrdescription, Err.source, "FilterDetailsCuLib.asp", "CleanFilterDetails", "", "Error on CleanFilterDetails", LogLevelError)
	End If
	sFilter = Replace(sFilter, Chr(160), "&nbsp;")
	Err.Clear
End Function
%>
