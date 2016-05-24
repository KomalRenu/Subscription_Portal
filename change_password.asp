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

	Dim sOldPassword
	Dim sNewPassword
	Dim sConfirmNewPassword
	Dim sHint
	Dim sOptSection
	Dim bSaved
	Dim sGetUserPropertiesXML
	Dim sLocaleID
	Dim sDefAddID

	lValidationError = NO_ERR
	bSaved = False
	sOptionsStyle = ""
	sOptSection = "2"

	If oRequest("chPwdCancel").Count > 0 Then
		Response.Redirect "options.asp"
	End If

	lErr = ParseRequestForChangePassword(oRequest, sOldPassword, sNewPassword, sConfirmNewPassword, sHint)

	If lErr = NO_ERR Then
		If Request.Form.Count > 0 Then
			lValidationError = validate_ChangeUserPassword(sOldPassword, sNewPassword, sConfirmNewPassword, sHint)
			If lValidationError = NO_ERR Then
			    lErr = cu_UpdateUserPassword(sOldPassword, sNewPassword)
			    If lErr = NO_ERR Then
			        lErr = cu_GetUserProperties(sGetUserPropertiesXML)
			        If lErr = NO_ERR Then
			            lErr = GetVariablesFromXML_ChangePassword(sGetUserPropertiesXML, sLocaleID, sDefAddID)
			            If lErr = NO_ERR Then
			                lErr = cu_UpdateUserProperties(sHint, sLocaleID, sDefAddID)
			            End If
			        End If
			        If lErr = NO_ERR Then
    					bSaved = True
			        End If
			    End If
			End If
		End If
	End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(143)'Descriptor: Change my password%> - MicroStrategy Narrowcast Server</TITLE>
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
			            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(26) & " " 'Descriptor: You are here:%> <%Response.Write asDescriptors(286) 'Descriptor: Preferences%> > <B><%Response.Write asDescriptors(143) 'Descriptor: Change my password%></B></FONT>
			        </TD>
			        <TD ALIGN="RIGHT"><IMG SRC="images/desktop_preferences.gif" WIDTH="60" HEIGHT="60" BORDER="0" ALT="" /></TD>
			    </TR>
			</TABLE>
			<%
			If lErr <> NO_ERR Then
				Call DisplayLoginError(sErrorHeader, sErrorMessage)
			ElseIf lValidationError <> NO_ERR Then
			    Call DisplayLoginError(sErrorHeader, sErrorMessage)
			End If
			%>
				<% If bSaved = True Then %>
					<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
					<FORM ACTION="change_password.asp" METHOD="POST">
						<TR>
							<TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(181) 'Descriptor: Your password has been changed.%></B></FONT></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="10" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
						    <TD><INPUT CLASS="buttonClass" TYPE="SUBMIT" NAME="chPwdCancel" VALUE="<%Response.Write asDescriptors(543)'Descriptor: OK%>" /></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
						</TR>
					</FORM>
					</TABLE>
					<BR />
				<% Else %>
				<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
				<FORM ACTION="change_password.asp" METHOD="POST">
					<INPUT TYPE="HIDDEN" NAME="userName" VALUE="<%Response.Write GetUsername()%>" />
					<TR>
						  <TD WIDTH="300" ALIGN="LEFT" VALIGN="TOP">
							<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2" WIDTH="100%">
								<TR>
								    <TD WIDTH="1%" VALIGN="TOP"></TD>
								    <TD WIDTH="99%" VALIGN="TOP">
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								            <B><%Response.Write asDescriptors(14) 'Descriptor: Please type in your old password and new password.%></B>
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
                                                    If (lValidationError And ERR_LOGIN_BLANKS) Then
                                                        Response.Write asDescriptors(444) & "<BR />" 'Descriptor: Either the old or new password was blank.  Please enter them again.
                                                    End If
                                                    If (lValidationError And ERR_CONFIRM_PASSWORD) Then
                                                        Response.write asDescriptors(392) & "<BR />" 'Descriptor: The Password and Confirm password did not match.
                                                    End If
                                                    %>
                                                </B>
                                            </FONT>
                                        </TD>
								    </TR>
								<%End If%>
								<TR>
								    <TD></TD>
								    <TD>
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
										<%Response.Write asDescriptors(369)'Descriptor: User name:%>&nbsp;&nbsp;<B><%Response.Write " " & GetUsername()%></B><BR /><BR />
                                        <%Response.Write asDescriptors(183)'Descriptor: Old password:%><BR />
										<INPUT TYPE="PASSWORD" NAME="oldPwd" SIZE="25" MAXLENGTH="250" STYLE="font-family: courier" /><BR />
										<%Response.Write asDescriptors(184)'Descriptor: New password:%><BR />
										<INPUT TYPE="PASSWORD" NAME="newPwd" SIZE="25" MAXLENGTH="250" STYLE="font-family: courier" /><BR />
										<%Response.Write asDescriptors(397) 'Descriptor: Confirm new password:%><BR />
										<INPUT TYPE="PASSWORD" NAME="confirmNewPwd" SIZE="25" MAXLENGTH="250" STYLE="font-family: courier" />
										</FONT>
								    </TD>
								</TR>
								<TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT=""></TD></TR>
								<TR>
								    <TD VALIGN="TOP"></TD>
								    <TD>
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								            <B><%Response.Write asDescriptors(436) 'Descriptor: Please enter a hint that will help you remember your password.%></B>
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
                                                    If (lValidationError And ERR_HINT_BLANK) Then
                                                        Response.Write asDescriptors(441) 'Descriptor: The password hint was blank.  Please enter it again.
                                                    End If
                                                    %>
                                                </B>
                                            </FONT>
                                        </TD>
								    </TR>
								<%End If%>
								<TR>
								    <TD></TD>
								    <TD>
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
                                            <%Response.Write asDescriptors(395) 'Descriptor: Password hint:%><BR />
										    <INPUT TYPE="TEXT" NAME="Hint" SIZE="25" MAXLENGTH="250" STYLE="font-family: courier" />
										</FONT>
								    </TD>
								</TR>
								<TR>
								    <TD></TD>
								    <TD>
                                        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
									    	<TR>
									    		<TD><INPUT TYPE="SUBMIT" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(182) 'Descriptor: Change password%>" /></TD>
									    		<TD>&nbsp;</TD>
									    		<TD><INPUT CLASS="buttonClass" TYPE="SUBMIT" NAME="chPwdCancel" VALUE="<%Response.Write asDescriptors(120)'Descriptor: Cancel%>" /></TD>
									    	</TR>
									    </TABLE>
								    </TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</FORM>
				</TABLE>
				<% End If %>
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