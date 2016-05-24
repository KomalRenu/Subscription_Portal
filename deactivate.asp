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
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sOptSection
	Dim bSaved
	Dim sPassword

	sOptionsStyle = ""
	sOptSection = "3"
	bSaved = False

	If oRequest("deactivateCancel").Count > 0 Then
		Response.Redirect "options.asp"
	End If

	lErr = ParseRequestForDeactivateUser(oRequest, sPassword)

	If lErr = NO_ERR Then
	    If Request.Form.Count > 0 Then
            lErr = cu_DeactivateUser(sPassword)
            	If lErr = NO_ERR Then
	                Call Logout()
	                Response.Redirect "login.asp?status=deact"
	            End If
	    End If
	End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(460) 'Descriptor: Deactivate my account%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY BGCOLOR="ffffff" TOPMARGIN="0" LEFTMARGIN="0" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin search box -->
				<!-- #include file="searchbox.asp" -->
			<!-- end search box -->
			<!-- begin left menu -->
				<!-- #include file="toolbarUserOptions.asp" -->
			<BR />
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" VALIGN="TOP">
			<!-- begin center panel -->
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
			    <TR>
			        <TD VALIGN="CENTER">
			            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(26) & " " 'Descriptor: You are here:%> <%Response.Write asDescriptors(286) 'Descriptor: Preferences%> > <B><%Response.Write asDescriptors(460) 'Descriptor: Deactivate my account%></B></FONT>
			        </TD>
			        <TD ALIGN="RIGHT"><IMG SRC="images/desktop_preferences.gif" WIDTH="60" HEIGHT="60" BORDER="0" ALT="" /></TD>
			    </TR>
			</TABLE>
			<BR />
			<%
			If lErr <> NO_ERR Then
			    Call DisplayLoginError(sErrorHeader, sErrorMessage)
			End If
			%>
				<% If bSaved = True Then %>
					<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
						<TR>
							<TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(461) 'Descriptor: Your account has been deactivated.%></B></FONT></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
						</TR>
					</TABLE>
					<BR />
				<% End If %>
				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
				<FORM NAME="deactivateForm" METHOD="POST" ACTION="deactivate.asp">
				    <TR>
				        <TD>
				            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(462) 'Descriptor: Are you sure you want to deactivate your account?%></B></FONT>
				        </TD>
				    </TR>
				    <TR>
				        <TD>
				            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>">
				                <%Response.Write asDescriptors(463) 'Descriptor: Once you deactivate your account, you will no longer receive services or be able to log in. Your account can only be reactivated by the portal administrator.%>
				            </FONT>
				        </TD>
				    </TR>
				    <TR>
				        <TD ALIGN="CENTER">
				            <BR />
				            <TABLE BORDER="0" CELLPADDING="3" CELLSPACING="0">
				                <TR>
				                    <TD>
				                        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%Response.Write asDescriptors(370) 'Descriptor: Password:%></B></FONT>
				                    </TD>
				                    <TD>
				                        <INPUT TYPE="PASSWORD" NAME="deactPassword" SIZE="15" />
				                    </TD>
				                </TR>
				                <TR>
				                    <TD></TD>
				                    <TD>
				                        <INPUT CLASS="buttonClass" TYPE="SUBMIT" VALUE="<%Response.Write asDescriptors(460) 'Descriptor: Deactivate my account%>" /> <INPUT CLASS="buttonClass" TYPE="SUBMIT" NAME="deactivateCancel" VALUE="<%Response.Write asDescriptors(120) 'Descriptor: Cancel%>" />
				                    </TD>
				                </TR>
				            </TABLE>
				        </TD>
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