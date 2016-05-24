<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Dim sDayField
Dim sMonthField
Dim sYearField
Dim sDateFormat
%>
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<TR><TD COLSPAN="2">
		<FONT FACE="Verdana,Arial,MS Sans Serif" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#FFFFFF"><B><%Response.Write asDescriptors(288) 'Descriptor: Select a Day:%></B></FONT>
	</TD></TR>
	<%
	sDayField = "<TR><TD><FONT FACE=""Verdana,Arial,MS Sans Serif"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"" >&nbsp;&nbsp;&nbsp;" & asDescriptors(290) & "&nbsp;&nbsp;</FONT></TD>" 'Descriptor: Day
	sDayField = sDayField & "<TD><SELECT NAME=""cbDay"">"
		For i = 1 To 31
			sDayField = sDayField & "<OPTION VALUE=""" & i & """"
			If CInt(sDay) = i Then sDayField = sDayField & " SELECTED=""1"""
			sDayField = sDayField & ">" & i & "</OPTION>"
		Next
	sDayField = sDayField & "</SELECT></TD></TR>"

	sMonthField = "<TR><TD><FONT FACE=""Verdana,Arial,MS Sans Serif"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">&nbsp;&nbsp;&nbsp;" & asDescriptors(289) & "&nbsp;&nbsp;</FONT></TD>" 'Descriptor: Month
	sMonthField = sMonthField & "<TD><SELECT NAME=""cbMonth"">"
		For i = 1 To 12
			sMonthField = sMonthField & "<OPTION VALUE=""" & i & """"
			If CInt(sMonth) = i Then sMonthField = sMonthField & " SELECTED=""1"""
			sMonthField = sMonthField & ">" & i & "</OPTION>"
		Next
	sMonthField = sMonthField & "</SELECT></TD></TR>"

	sYearField = "<TR><TD><FONT FACE=""Verdana,Arial,MS Sans Serif"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">&nbsp;&nbsp;&nbsp;" & asDescriptors(291) & "&nbsp;&nbsp;</FONT></TD>" 'Descriptor: Year
	sYearField = sYearField & "<TD><INPUT TYPE=""TEXT"" NAME=""tYear"" VALUE=""" & sYear & """ SIZE=""4"" STYLE=""font-family: courier"" /></TD></TR>"

	sDateFormat = asDescriptors(327) 'Descriptor: MM/DD/YYYY
	sDateFormat = Replace(sDateFormat, "/", "")
	sDateFormat = Replace(sDateFormat, "-", "")
	sDateFormat = Replace(sDateFormat, ".", "")
	sDateFormat = Replace(sDateFormat, "MM", "**MM**")
	sDateFormat = Replace(sDateFormat, "DD", "**DD**")
	sDateFormat = Replace(sDateFormat, "YYYY", "**YYYY**")
	sDateFormat = Replace(sDateFormat, "**MM**", sMonthField)
	sDateFormat = Replace(sDateFormat, "**DD**", sDayField)
	sDateFormat = Replace(sDateFormat, "**YYYY**", sYearField)
	Response.Write sDateFormat
	%>
</TABLE>