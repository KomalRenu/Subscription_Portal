<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Private Const FOR_READING_SHOWERRORS = 1
Private Const MESSAGES_SHOWERRORS = 20

Private Const SHOWERRORS_S_ERROR = 0
Private Const SHOWERRORS_S_WARNING = 1
Private Const SHOWERRORS_S_MESSAGES = 2
Private Const SHOWERRORS_S_DAY = 3
Private Const SHOWERRORS_S_MONTH = 4
Private Const SHOWERRORS_S_YEAR = 5

Function RecieveShowErrorsRequest(oRequest, sOrderBy, sSortOrder, sWarn, sErr, sMessage, iFirstIndex, iLastIndex, sDay, sMonth, sYear, bReturn, asDisplayDate)
'*********************************************************************************************
' Purpose:	Get input parameters from form
' Inputs:	oRequest
' Outputs:	sOrderBy, sSortOrder, sWarn, sErr, sMessage, iFirstIndex, iLastIndex, sDay, sMonth, sYear, bReturn, asDisplayDate
'*********************************************************************************************
On Error Resume Next
	Dim bFirstTime
	Dim sFirst
	Dim sLast

	bFirstTime = False
	'If (Len(CStr(oRequest("First"))) > 0) Or (Len(RemoveParameterFromURL(oRequest, "toolbar")) = 0) Then
	If (Len(CStr(oRequest("First"))) > 0) Or (Len(RemoveParameterFromURL(oRequest, "showHelp")) = 0) Or (Len(RemoveParameterFromURL(oRequest, "src")) = 0) Then
		bFirstTime = True
	End If

	sOrderBy = CStr(oRequest("OrderBy"))
	sSortOrder = CStr(oRequest("SortOrder"))
	If Len(sOrderBy) = 0 Then
		sOrderBy = "Time"
		If Len(sSortOrder) = 0 Then
			sSortOrder = "DESC"
		End If
	End If


	If not bFirstTime Then
		sWarn = CStr(oRequest("cbWarning"))
		If Len(sWarn) = 0 Then sWarn = ""
		asDisplayDate(SHOWERRORS_S_WARNING) = sWarn

		sErr = CStr(oRequest("cbError"))
		If Len(sErr) = 0 Then sErr = ""
		asDisplayDate(SHOWERRORS_S_ERROR) = sErr

		sMessage = CStr(oRequest("cbMessages"))
		If Len(sMessage) = 0 Then sMessage = ""
		asDisplayDate(SHOWERRORS_S_MESSAGES) = sMessage
	Else
		sWarn = "on"
		sErr = "on"
		sMessage = "on"
		asDisplayDate(SHOWERRORS_S_WARNING) = "on"
		asDisplayDate(SHOWERRORS_S_ERROR) = "on"
		asDisplayDate(SHOWERRORS_S_MESSAGES) = "on"
	End If

	sFirst = CStr(oRequest("hFirst"))
	sLast = CStr(oRequest("hLast"))
	sPrevious = CStr(oRequest("bPrev.x"))
	sNext = CStr(oRequest("bNext.x"))

	If Len(sPrevious) > 0 Then
		iFirstIndex = sFirst - MESSAGES_SHOWERRORS
		iLastIndex = sFirst - 1
	ElseIf Len(sNext) > 0 Then
		iFirstIndex = sLast + 1
		iLastIndex = sLast + MESSAGES_SHOWERRORS
	Else
		iFirstIndex = 1
		iLastIndex = MESSAGES_SHOWERRORS
	End If

	sDay = CStr(oRequest("cbDay"))
	If Len(sDay) = 0 Then sDay = Day(Date)
	asDisplayDate(SHOWERRORS_S_DAY) = sDay

	sMonth = CStr(oRequest("cbMonth"))
	If Len(sMonth) = 0 Then sMonth = Month(Date)
	asDisplayDate(SHOWERRORS_S_MONTH) = sMonth

	sYear = CStr(oRequest("tYear"))
	If Len(sYear) = 0 Then sYear = Year(Date)
	asDisplayDate(SHOWERRORS_S_YEAR) = sYear

	bReturn = CStr(oRequest("bContinue"))
	bReturn = (Len(bReturn) > 0)

	RecieveShowErrorsRequest = err.number
	Err.Clear
End Function

Function CreateRequestForShowErrors(oRequest)
'********************************************************
'*Purpose: Based on the aSvcConfigInfo array, creates the string that can be used
'           as the parameters of the link to a page.
'*Inputs:  aSvcConfigInfo: an array with the information needed to config a service
'*Outputs: This functions returns the string directly, not an error
'********************************************************
	Dim sRequest

    sRequest = ""

	'cbMonth=4&cbDay=2&tYear=2001&cbError=on&cbWarning=on&cbMessages=on&bRefresh=Refresh&hFirst=1&hLast=4&OrderBy=Time&SortOrder=DESC

    If Len(oRequest("cbMonth"))> 0 Then sRequest = sRequest & "&cbMonth=" & oRequest("cbMonth")
    If Len(oRequest("cbDay"))> 0 Then sRequest = sRequest & "&cbDay=" & oRequest("cbDay")
    If Len(oRequest("tYear"))> 0 Then sRequest = sRequest & "&tYear=" & oRequest("tYear")
    If Len(oRequest("cbError"))> 0 Then sRequest = sRequest & "&cbError=" & oRequest("cbError")
    If Len(oRequest("cbWarning"))> 0 Then sRequest = sRequest & "&cbWarning=" & oRequest("cbWarning")
    If Len(oRequest("cbMessages"))> 0 Then sRequest = sRequest & "&cbMessages=" & oRequest("cbMessages")
    If Len(oRequest("hFirst"))> 0 Then sRequest = sRequest & "&hFirst=" & oRequest("hFirst")
    If Len(oRequest("hLast"))> 0 Then sRequest = sRequest & "&hLast=" & oRequest("hLast")
    If Len(oRequest("OrderBy"))> 0 Then sRequest = sRequest & "&OrderBy=" & oRequest("OrderBy")
    If Len(oRequest("SortOrder"))> 0 Then sRequest = sRequest & "&SortOrder=" & oRequest("SortOrder")
    If Len(oRequest("src"))> 0 Then sRequest = sRequest & "&src=" & oRequest("src")

    If Len(sRequest) > 0 Then sRequest = Mid(sRequest, 2)

    CreateRequestForShowErrors = sRequest

End Function


Sub reportError(where, error)
'*********************************************************************************************
' Purpose:	Show the error when there is one
' Inputs:	where, error
' Outputs:
'*********************************************************************************************
	Response.Write("<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """><B>" + asDescriptors(172) + " '" + where + "'</B></FONT> <BLOCKQUOTE><XMP>" + error.reason + "</XMP></BLOCKQUOTE>") 'Descriptor: Error loading
End Sub

Function isValidDate (iDay, iMonth, iYear)
'*********************************************************************************************
' Purpose:	Validate a date
' Inputs:	iDay, iMonth, iYear
' Outputs:	isValidDate
'*********************************************************************************************
	Dim bLeap
	Dim iLeap
	iLeap = 0

	If Not (iYear >= 1000 And iYear <= 3000) Then
		isValidDate = False
		Exit Function
	Else
		If iYear Mod 1000 = 0 Then
			bLeap = True
		Elseif iYear Mod 100 = 0 Then
			bLeap = False
		Elseif iYear Mod 4 = 0 Then
			bLeap = True
		Else
			bLeap = False
		End If
	End If

	If Not (iMonth >= 1 And iMonth <= 12) Then
		isValidDate = False
		Exit Function
	End If

	If iMonth = 1 Or iMonth = 3 Or _
		iMonth = 5 Or iMonth = 7 Or _
		iMonth = 8 Or iMonth = 10 Or _
		iMonth = 12 Then
		If Not (iDay >= 1 And iDay <= 31) Then
			isValidDate = False
			Exit Function
		End If
	ElseIf iMonth = 4 Or iMonth = 6 Or _
		iMonth = 9 Or iMonth = 11 Then
		If Not (iDay >= 1 And iDay <= 30) Then
			isValidDate = False
			Exit Function
		End If
	Else 'February
		If bLeap Then iLeap = 1
		If Not (iDay >= 1 And CInt(iDay) <= CInt((28+iLeap))) Then
			isValidDate = False
			Exit Function
		End If
	End If
	isValidDate	= True
End Function

Function SelectSortInfo(sOrderBy, sSortOrder, sItem, asDisplayDate, sSortArrow, sSortTitle, sURL)
'*********************************************************************************************
' Purpose:	Select sorting icon and title
' Inputs:	sOrderBy, sSortOrder, sItem, asDisplayDate
' Outputs:	sSortArrow, sSortTitle, sURL
'*********************************************************************************************
	On Error Resume Next
	If StrComp(sOrderBy, sItem) = 0 Then
		If StrComp(sSortOrder, "ASC") = 0 Then
			sSortArrow = "sort_asc.gif"
			sSortTitle = asDescriptors(108) 'Sort descending
			sURL = "OrderBy=" & sItem & "&SortOrder=DESC"
		Else
			sSortArrow = "sort_desc.gif"
			sSortTitle = asDescriptors(107) 'Sort ascending
			sURL = "OrderBy=" & sItem & "&SortOrder=ASC"
		End If
	Else
		sSortArrow = "sort_row.gif"
		If StrComp(sItem, "Time") = 0 Then
			sSortTitle = asDescriptors(108) 'Sort descending
			sURL = "OrderBy=" & sItem & "&SortOrder=DESC"
		Else
			sSortTitle = asDescriptors(107) 'Sort ascending
			sURL = "OrderBy=" & sItem & "&SortOrder=ASC"
		End If
	End If
	sURL = sURL & "&cbDay=" & asDisplayDate(SHOWERRORS_S_DAY) & "&cbMonth=" & asDisplayDate(SHOWERRORS_S_MONTH) & "&tYear=" & asDisplayDate(SHOWERRORS_S_YEAR) & "&cbError=" & asDisplayDate(SHOWERRORS_S_ERROR) & "&cbWarning=" & asDisplayDate(SHOWERRORS_S_WARNING) & "&cbMessages=" & asDisplayDate(SHOWERRORS_S_MESSAGES)
	Err.Clear
End Function

Function ShowErrors(opreXML, aFontInfo, sOrderBy, sSortOrder)
'*********************************************************************************************
' Purpose:	Show the errors table
' Inputs:	opreXML, aFontInfo, sOrderBy, sSortOrder
' Outputs:	HTML
'*********************************************************************************************
	On Error Resume Next
	Dim oRowXML
	Dim vRow
	Dim sSortArrow, sSortTitle, sURL
	'Headers
	Response.Write "<TABLE WIDTH=""98%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
	Response.Write "<TR BGCOLOR=""#6699CC"">"

	'Descriptors: Sort ascending | Name | Sort ascending
	'sOrderBy, sSortOrder
	Call SelectSortInfo(sOrderBy, sSortOrder, "Time", asDisplayDate, sSortArrow, sSortTitle, sURL)
	Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#FFFFFF"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>&#160;" & asDescriptors(174)  & "&#160;</B></FONT><A HREF=""ShowErrors.asp?" & sURL & """ TITLE=""" & sSortTitle & """><IMG SRC=""../Images/" & sSortArrow & """ WIDTH=""17"" HEIGHT=""8"" ALT=""" & sSortTitle & """ BORDER=""0"" /></A></TD>" 'Descriptor: Time
	Response.Write "<TD>&#160;&#160;</TD>"

	Call SelectSortInfo(sOrderBy, sSortOrder, "User", asDisplayDate, sSortArrow, sSortTitle, sURL)
	Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#FFFFFF"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>&#160;" & asDescriptors(688) & "&#160;</B></FONT><A HREF=""ShowErrors.asp?" & sURL &""" TITLE=""" & sSortTitle & """><IMG SRC=""../Images/" & sSortArrow & """ WIDTH=""17"" HEIGHT=""8"" ALT=""" & sSortTitle & """ BORDER=""0"" /></A></TD>" 'Descriptor: User IP
	Response.Write "<TD>&#160;&#160;</TD>"

	Call SelectSortInfo(sOrderBy, sSortOrder, "Level", asDisplayDate, sSortArrow, sSortTitle, sURL)
	Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#FFFFFF"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>&#160;" & asDescriptors(180) & "&#160;</B></FONT><A HREF=""ShowErrors.asp?" & sURL & """ TITLE=""" & sSortTitle & """><IMG SRC=""../Images/" & sSortArrow & """ WIDTH=""17"" HEIGHT=""8"" ALT=""" & sSortTitle & """ BORDER=""0"" /></A></TD>" 'Descriptor: Level
	Response.Write "<TD>&#160;&#160;</TD>"

	Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#FFFFFF"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>&#160;" & asDescriptors(123)  & "<BR />&#160;" & asDescriptors(485)  & "<BR />&#160;" & asDescriptors(175) & "</B></FONT></TD>" 'Descriptor: Error number | Error source | Error description
	Response.Write "<TD>&#160;&#160;</TD>"
	Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#FFFFFF"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>&#160;" & asDescriptors(176) & "&#160;</B></FONT></TD>" 'Descriptor: File
	Response.Write "<TD>&#160;&#160;</TD>"
	Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#FFFFFF"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>&#160;" & asDescriptors(177) & "&#160;</B></FONT></TD>" 'Descriptor: ASP function
	Response.Write "<TD>&#160;&#160;</TD>"
	Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#FFFFFF"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>&#160;" & asDescriptors(178) & "&#160;</B></FONT></TD>" 'Descriptor: API function
	Response.Write "<TD>&#160;&#160;</TD>"
	Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#FFFFFF"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>&#160;" & asDescriptors(179) & "&#160;</B></FONT></TD>" 'Descriptor: Comments
	Response.Write "<TD>&#160;&#160;</TD>"
	Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#FFFFFF"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>&#160;" & asDescriptors(685) & "&#160;</B></FONT></TD>" 'Descriptor: Origin
	Response.Write "</TR>"

	'Body
	Response.Write  "<TR BGCOLOR=""#003366""><TD HEIGHT=""1"" COLSPAN=""17""><IMG SRC=""../Images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"
	Response.Write "<TR><TD HEIGHT=""1"" COLSPAN=""17""><IMG SRC=""../Images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"
	Set oRowXML = opreXML.selectSingleNode("/ERRORS")
	If oRowXML.hasChildNodes Then
		Set vRow = oRowXML.firstChild
		If Not IsEmpty(vRow) Then
			Do While Not vRow is Nothing
				For Each vRow In oRowXML
					Response.Write "<TR>"
					Response.Write "<TD NOWRAP=""1""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>" & vRow.SelectSingleNode("time").text & "</B></FONT></TD>"
					Response.Write "<TD></TD>"
					Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & vRow.SelectSingleNode("user").text & "</FONT></TD>"
					Response.Write "<TD></TD>"

					Select Case CInt(vRow.SelectSingleNode("errLevel").text)
					Case LogLevelError
						Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(39) & "</FONT></TD>"		'Descriptor: Error
					Case LogLevelTrace
						Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(692) & "</FONT></TD>"	'Descriptor: Trace
					Case LogLevelInfo
						Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(693) & "</FONT></TD>"		'Descriptor: Information
					Case LogLevelWarning
						Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(694) & "</FONT></TD>"	'Descriptor: Warning
					End Select

					Response.Write "<TD></TD>"
					Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & vRow.SelectSingleNode("errID").text & "<BR />" & vRow.SelectSingleNode("errSrc").text & "<BR />" & escapeChars(vRow.SelectSingleNode("errDesc").text) & "</FONT></TD>"
					Response.Write "<TD></TD>"
					Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & vRow.SelectSingleNode("file").text & "</FONT></TD>"
					Response.Write "<TD></TD>"
					Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & vRow.SelectSingleNode("ASPFunc").text & "</FONT></TD>"
					Response.Write "<TD></TD>"
					Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & vRow.SelectSingleNode("APIFunc").text & "</FONT></TD>"
					Response.Write "<TD></TD>"
					Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & escapeChars(vRow.SelectSingleNode("comments").text) & "</FONT></TD>"
					Response.Write "<TD></TD>"
					Select Case CInt(vRow.SelectSingleNode("errOrigin").text)
					Case LogErrorOriginPortal
						Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(686) & "</FONT></TD>"	'Descriptor: Portal
					Case LogErrorOriginAdmin
						Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(687) & "</FONT></TD>"	'Descriptor: Admin
					End Select
					Response.Write "</TR>"
					Response.Write "<TR BGCOLOR=""#99CCFF""><TD HEIGHT=""1"" COLSPAN=""17""><IMG SRC=""../Images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"
				Next
				Set vRow = vRow.nextSibling
				If vRow is Nothing Then Exit Do
			Loop
		End If
	End If
	Response.Write "</TABLE>"
	Set vRow = Nothing
	Set oRowXML = Nothing
	Err.Clear
End Function

Function  escapeChars(sInputString)
'*************************************************************
' Purpose: escape HTML chars
' Inputs: string
' Outputs: string
'*************************************************************
On Error Resume Next
	Dim sModifiedString

	sModifiedString = Replace(sInputString, "<", "&lt;")
	sModifiedString = Replace(sModifiedString, ">", "&gt;")

	escapeChars = sModifiedString

End Function

%>