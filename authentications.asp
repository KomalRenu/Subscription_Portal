<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/OptionsCuLib.asp" -->
<!-- #include file="CustomLib/LoginCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sGetInformationSourcesForSiteXML
	Dim sUserAuthenticationObjectsXML
	Dim sOptSection
	Dim bSaved
	Dim bHasProjects

	sOptionsStyle = ""
	sOptSection = "4"
	bSaved = False
	bHasProjects = False

	If oRequest("chAuthOK").Count > 0 Then
		Response.Redirect "options.asp"
	End If

    lErr = cu_GetInformationSourcesForSite(sGetInformationSourcesForSiteXML, bHasProjects)
	lErr = cu_GetUserAuthenticationObjects(sUserAuthenticationObjectsXML)

	If lErr = NO_ERR Then
		If oRequest("authSave").Count > 0 Then
			lValidationError = validate_ChangeAuthentications(oRequest)
		    If lValidationError = NO_ERR Then
                lErr = cu_UpdateUserAuthenticationObjects(sUserAuthenticationObjectsXML, oRequest)
                If lErr = NO_ERR Then
                    bSaved = True
                ElseIf lErr = ERR_ISLOGIN_ERROR Then
                    lValidationError = ERR_ISLOGIN_ERROR
                    lErr = NO_ERR
                End If
			End If
		End If
	End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(466) 'Descriptor: Change Information Source credentials%> - MicroStrategy Narrowcast Server</TITLE>
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
			            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(26) & " " 'Descriptor: You are here:%> <%Response.Write asDescriptors(286) 'Descriptor: Preferences%> > <B><%Response.Write asDescriptors(466) 'Descriptor: Change Information Source credentials%></B></FONT>
			        </TD>
			        <TD ALIGN="RIGHT"><IMG SRC="images/desktop_preferences.gif" WIDTH="60" HEIGHT="60" BORDER="0" ALT="" /></TD>
			    </TR>
			</TABLE>
			<%
			If lErr <> NO_ERR Or lValidationError <> NO_ERR Then
				Call DisplayLoginError(sErrorHeader, sErrorMessage)
			End If
			%>
				<% If bSaved = True Then %>
					<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
					<FORM ACTION="authentications.asp" METHOD="POST" id=form1 name=form1>
						<TR>
							<TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(357) 'Descriptor: Your Information Source credentials have been changed.%></B></FONT></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="10" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
						    <TD><INPUT CLASS="buttonClass" TYPE="SUBMIT" NAME="chAuthOK" VALUE="<%Response.Write asDescriptors(543)'Descriptor: OK%>" /></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
						</TR>
					</TABLE>
					<BR />
				<%Else%>
				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
				<FORM NAME="authenticationsForm" METHOD="POST" ACTION="authentications.asp">
					<TR>
						  <TD WIDTH="300" ALIGN="LEFT" VALIGN="TOP">

							<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2" WIDTH="100%">
								<TR>
								    <TD WIDTH="1%" VALIGN="TOP"></TD>
								    <TD WIDTH="99%" VALIGN="TOP">
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								            <B><%Response.Write asDescriptors(438) 'Descriptor: Please enter your login credentials for the following Information Source(s).%></B>
								            <%Response.Write Replace(asDescriptors(434), "*", "<FONT COLOR=""#cc0000"">*</FONT>") 'Descriptor: Required information is noted with a red asterisk.%>
								        </FONT>
								    </TD>
								</TR>
								<%If lValidationError <> NO_ERR Then%>
								    <TR>
                                        <TD></TD>
                                        <TD>
                                            <FONT COLOR="#cc0000" FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
                                                <B>
                                                    <%
                                                    If (lValidationError And ERR_ISLOGIN_BLANK) Then
                                                        Response.Write asDescriptors(445) 'Descriptor: Please enter the required user name(s) and password(s) below.
                                                    End If
                                                    If (lValidationError And ERR_ISLOGIN_ERROR) Then
                                                        Response.Write asDescriptors(531) 'Descriptor: Either the User name or Password was incorrect for one or more Information Sources.  Please try again.
                                                    End If
                                                    %>
                                                </B>
                                            </FONT>
                                        </TD>
								    </TR>
								<%End If
                                If bSaved Then
									Call RenderInformationSourceLogins(sGetInformationSourcesForSiteXML, "")
								Else
									Call RenderInformationSourceLogins(sGetInformationSourcesForSiteXML, sUserAuthenticationObjectsXML, false)
								End If%>
				                <TR>
				                    <TD></TD>
				                    <TD>
				                        <%If bHasProjects Then%><INPUT NAME="authSave" CLASS="buttonClass" TYPE="SUBMIT" VALUE="<%Response.Write asDescriptors(467) 'Descriptor: Change credentials%>" /><%End If%> <INPUT CLASS="buttonClass" TYPE="SUBMIT" NAME="authCancel" VALUE="<%Response.Write asDescriptors(120) 'Descriptor: Cancel%>" />
				                    </TD>
				                </TR>
							</TABLE>
						</TR>
					</TD>
				</FORM>
				</TABLE>
				<%End If%>

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