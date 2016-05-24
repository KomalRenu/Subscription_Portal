<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
   <%If aPageInfo(N_TOOLBARS_PAGE) AND HELP_TOOLBAR Then%>
	<TR>
		<TD>
			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2">
				<TR>
					<TD WIDTH="13" VALIGN="TOP"><A HREF="<%Response.Write aPageInfo(S_NAME_PAGE)%>?<%Response.Write SwitchToolbarURL(aPageInfo(S_NAME_PAGE), oRequest, "showHelp", "0")  'ReplaceURLValue(oRequest, "showHelp", "0")%>"><IMG SRC="Images/1arrow_down.gif" WIDTH="13" HEIGHT="13" ALT="<%Response.Write asDescriptors(297) 'Descriptor: Hide help%>" BORDER="0" /></A></TD>
					<TD COLSPAN="2" VALIGN="MIDDLE"><A HREF="<%Response.Write aPageInfo(S_NAME_PAGE)%>?<%Response.Write SwitchToolbarURL(aPageInfo(S_NAME_PAGE), oRequest, "showHelp", "0")  'ReplaceURLValue(oRequest, "showHelp", "0")%>" STYLE="text-decoration:none;" TITLE="<%Response.Write asDescriptors(297) 'Descriptor: Hide help%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><B><%Response.Write asDescriptors(231) 'Descriptor: NEED HELP?%></B></FONT></A></TD>
				</TR>
				<%If (iHelpFileID <> -1) Then
					For iIndexForQuestions = 0 To Ubound(aiQuestions)%>
						<TR>
							<TD><IMG SRC="Images/1ptrans.gif" WIDTH="13" HEIGHT="1" ALT="" BORDER="0" /></TD>
							<TD VALIGN="TOP" WIDTH="3"><IMG SRC="Images/bullet_arrow_right.gif" WIDTH="3" HEIGHT="8" ALT="" BORDER="0" /></TD>
							<TD VALIGN="TOP"><A HREF="Help.asp?FAQId=<%Response.Write iHelpFileID%>&tab=3#QA<%Response.Write (iIndexForQuestions + 1)%>" TARGET="_blank"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><%Response.Write asDescriptors(aiQuestions(iIndexForQuestions))%></FONT></A></TD>
						</TR>
					<%Next
				End If%>
				<TR>
					<TD><IMG SRC="Images/1ptrans.gif" WIDTH="13" HEIGHT="1" ALT="" BORDER="0" /></TD>
					<TD VALIGN="TOP" WIDTH="3"><IMG SRC="Images/bullet_arrow_right.gif" WIDTH="3" HEIGHT="8" ALT="" BORDER="0" /></TD>
					<TD VALIGN="TOP"><A HREF="Help.asp" TARGET="_blank"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><%Response.Write asDescriptors(242) 'Descriptor: Online help%></FONT></A></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
   <%Else%>
	<TR>
		<TD>
			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2">
				<TR>
					<TD VALIGN="TOP"><A HREF="<%Response.Write aPageInfo(S_NAME_PAGE)%>?<%Response.Write SwitchToolbarURL(aPageInfo(S_NAME_PAGE), oRequest, "showHelp", "1") 'ReplaceURLValue(oRequest, "showHelp", "1")%>"><IMG SRC="Images/1arrow_right.gif" WIDTH="13" HEIGHT="13" ALT="<%Response.Write asDescriptors(298) 'Descriptor: Show help%>" BORDER="0" /></A></TD>
					<TD VALIGN="MIDDLE"><A HREF="<%Response.Write aPageInfo(S_NAME_PAGE)%>?<%Response.Write SwitchToolbarURL(aPageInfo(S_NAME_PAGE), oRequest, "showHelp", "1")  'ReplaceURLValue(oRequest, "showHelp", "1")%>" STYLE="text-decoration:none;" TITLE="<%Response.Write asDescriptors(298) 'Descriptor: Show help%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><B><%Response.Write asDescriptors(231) 'Descriptor: NEED HELP?%></B></FONT></A></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
   <%End If%>
	<TR>
		<TD><IMG SRC="Images/1ptrans.gif" WIDTH="158" HEIGHT="1" ALT="" BORDER="0" /></TD>
	</TR>