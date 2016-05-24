<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/LoginCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	Dim sUserName
	Dim sPasswordHint

	sPasswordHint = ""

	If oRequest("Cancel").Count > 0 Then
		Response.Redirect "login.asp?site=" & sChannel 'Go back to login page.
	End If

	lErr = ParseRequestForPasswordHint(oRequest, sUserName)

	If lErr = NO_ERR Then
		If Request.Form.Count > 0 Then
			'lErr = cu_GetUserHint(sUsername, sPasswordHint)
			lErr = cu_GetUserHint(Trim(CStr(oRequest("userName"))), sPasswordHint)
		End If
	End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(409)'Descriptor: Forgot your password?%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY BGCOLOR="ffffff" TOPMARGIN="0" LEFTMARGIN="0" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="login_header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin left menu -->
			<BR /><BR />
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" VALIGN="TOP">
			<!-- begin center panel -->
			<BR />
			<%
			    If lErr <> NO_ERR Then
			        Call DisplayLoginError(sErrorHeader, sErrorMessage)
			    End If
			%>
				<TABLE WIDTH="500" BORDER="0" CELLSPACING="0" CELLPADDING="0">
					<FORM ACTION="password_hint.asp" METHOD="POST">
					<INPUT TYPE="HIDDEN" NAME="site" VALUE="<%Response.Write sChannel%>" />
					<TR>
						<TD BGCOLOR="#000000" WIDTH="11" ALIGN="LEFT" VALIGN="TOP"><IMG SRC="Images/loginUpperLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#000000" WIDTH="237" ALIGN="LEFT" VALIGN="MIDDLE"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>" COLOR="#FFFFFF"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(409)'Descriptor: Forgot your password?%><%Else%><%Response.Write UCase(asDescriptors(409))'Descriptor: Forgot your password?%><%End If%></B></FONT></TD>
						<TD BGCOLOR="#000000" WIDTH="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#000000" WIDTH="200" ALIGN="LEFT" VALIGN="MIDDLE"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>" COLOR="#FFFFFF">&nbsp;<B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write ""%><%Else%><%Response.Write UCase("")%><%End If%></B></FONT></TD>
						<TD BGCOLOR="#000000" WIDTH="11" ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="Images/loginUpperRightCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD COLSPAN="5" HEIGHT="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="2" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD BGCOLOR="#CCCCCC" WIDTH="11"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>

						  <TD BGCOLOR="#CCCCCC" WIDTH="237" ALIGN="LEFT" VALIGN="TOP">

							<TABLE BORDER="0" CELLSPACING="1" CELLPADDING="5" WIDTH="100%">

								<TR><TD>
									<%If Len(sPasswordHint) = 0 Then%>
									<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
										<%Response.Write asDescriptors(369)'Descriptor: User name:%><BR />

										<INPUT TYPE="TEXT" NAME="userName" SIZE="25" MAXLENGTH="250" STYLE="font-family: courier" /><BR />
										<BR /><BR />
									</FONT>
									<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
										<TR>
											<TD><INPUT TYPE="SUBMIT" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(411)'Descriptor: Display Password hint%>" /></TD>
											<TD>&nbsp;</TD>
											<TD><INPUT CLASS="buttonClass" TYPE="SUBMIT" NAME="Cancel" VALUE="<%Response.Write asDescriptors(120)'Descriptor: Cancel%>" /></TD>
										</TR>
									</TABLE>
									<%Else%>
										<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="3" WIDTH="100%"><TR><TD BGCOLOR="#ffffcc">
											<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
												<%Response.Write asDescriptors(412)'Descriptor: Your password hint is: %><BR /><BR />
												<B><%Response.Write sPasswordHint%></B>
											</FONT>
										</TD></TR></TABLE>
										<BR />
										<A HREF="login.asp?site=<%Response.Write sChannel%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%Response.Write asDescriptors(415) 'Descriptor: Back to Login%></B></FONT></A>
									<%End If%>
								</TD></TR>
							</TABLE>
						</TD>
							<TD WIDTH="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
							<TD BGCOLOR="#BDBDBD" WIDTH="200" ALIGN="LEFT" VALIGN="TOP">
								<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="5"><TR><TD>
									<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
										<%Response.Write asDescriptors(410) 'Descriptor: Enter your User name and click 'Display password hint'.%><BR /><BR />
										<%Response.Write asDescriptors(542) 'Descriptor: If your hint does not assist you in remembering your password, contact your System Administrator so your password may be reset.%><BR />
									</FONT>
								</TD></TR></TABLE>
							</TD>
							<TD BGCOLOR="#BDBDBD" WIDTH="11"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD BGCOLOR="#CCCCCC" WIDTH="11" ALIGN="LEFT" VALIGN="BOTTOM"><IMG SRC="Images/loginLowerLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#CCCCCC" WIDTH="237"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD WIDTH="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#BDBDBD" WIDTH="200" ><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#BDBDBD" WIDTH="11" ALIGN="RIGHT" VALIGN="BOTTOM"><IMG SRC="Images/loginLowerRightCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
					</TR>
					</FORM>
				</TABLE>
			<!-- end center panel -->
		</TD>
		<TD WIDTH="1%">
			<IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
	</TR>
</TABLE>
<BR />
<!-- begin footer -->
	<!-- #include file="footer.asp" -->
<!-- end footer -->
</BODY>
</HTML>