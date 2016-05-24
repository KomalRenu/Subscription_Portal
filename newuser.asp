<%'**cu_CreateUserCopyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/LoginCuLib.asp" -->
<!-- #include file="CustomLib/AddressesCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%

	If StrComp(CStr(Application("Allow_New_users")), "1", vbBinaryCompare) <> 0 Then
	    Response.Redirect "login.asp?site=" & sChannel 'Go back to login page.
	End If

	Dim sSessionID

	Dim sUserName
	Dim sPassword
	Dim sConfirmPassword
	Dim sHint
	Dim sDefEmail
	Dim sDefAddID
	Dim sLocaleID
	Dim sLanguageID
	Dim sSavePwd
	Dim bHasLocales
	Dim bHasProjects
	'???
	Dim bEncryptedFlag
	bEncryptedFlag = False
	'???
	Dim sGetLocalesForSiteXML
	Dim sGetInformationSourcesForSiteXML
	Dim nStepCount
	Dim IServerUser
	Dim NTUser
	nStepCount = 1

	bHasProjects = False

	If oRequest("Cancel").Count > 0 Then
		Response.Redirect "login.asp?site=" & sChannel 'Go back to login page.
	End If

	lErr = ParseRequestForNewUser(oRequest, sUserName, sPassword, sConfirmPassword, sHint, sDefEmail, sLocaleID, sLanguageID, sSavePwd)

	If lErr = NO_ERR Then

		lErr = cu_GetLocalesForSiteByCurrentLocale(sGetLocalesForSiteXML)
		If lErr = NO_ERR Then
		    bHasLocales = SiteHasLocales(sGetLocalesForSiteXML)
		    lErr = cu_GetInformationSourcesForSite(sGetInformationSourcesForSiteXML, bHasProjects)
		End If
	End If

	If lErr = NO_ERR Then
	    If Request.Form.Count > 0 Then
	    	lValidationError = validate_CreateUser(oRequest, sUserName, sPassword, sConfirmPassword, sHint)
	    	If lValidationError = NO_ERR Then
	    	    lErr = cu_CreateUser(sUserName, sPassword, sHint, sLocaleID, sLanguageID, sDefAddID)
	            If lErr = NO_ERR And lValidationError = NO_ERR Then
	            	lErr = cu_CreateSession(sUserName, sPassword, bEncryptedFlag, sSessionID)
	            	If lErr = 0 Then
	            	    Call SetSessionID("", sSessionID)
	            	End If
	            End If
	    	End If

	    	If lErr = NO_ERR And lValidationError = NO_ERR Then
	    	    lErr = cu_AddPortalAddress(sUserName)
	    	    If lErr = NO_ERR Then
	    	        If Len(sDefEmail) > 0 Then
	    	            lErr = AddDefaultAddress(sUserName, sDefAddID, sDefEmail)
	    	        End If
	    	    End If
	    	End If

	    	If lErr = NO_ERR And lValidationError = NO_ERR Then
	    	    lErr = SetSessionInfo(sChannel, sSessionID, sUserName, sSavePwd, sLocaleID, sLanguageID)
	    	    If lErr = NO_ERR Then
	    	        lErr = cu_SaveUserAuthenticationObjects(oRequest)
	    	        If lErr = ERR_ISLOGIN_ERROR Then
	    	            lValidationError = ERR_ISLOGIN_ERROR
	    	            lErr = NO_ERR
	    	        ElseIf lErr = API_ERR_LOGIN_PASSWORD_INVALID Then
	    	            lValidationError = ERR_ISLOGIN_ERROR
	    	            lErr = NO_ERR
	    	        End If
	    	    End If
	    	End If

	        If lErr = NO_ERR And lValidationError = NO_ERR Then
	            'Response.Redirect "default.asp?account=new"

				'Set the cookies used for Iserver Authentication to be false
				IServerUser =	false
				NTUser = "no"
				Call SetIServerCookies("","",IServerUser,NTUser)

	            Response.Redirect "newUserDetails.asp?getUD=1&account=new"
	        End If

	        If (lErr <> NO_ERR) Or (lValidationError <> NO_ERR) Then
	            Call cu_DeleteUser()
               	Call cu_CloseSession()
	            Call Logout()
	        End If
	    End If
	End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(372) 'Descriptor: Create New User%> - MicroStrategy Narrowcast Server</TITLE>
	<%If StrComp(GetJavaScriptSetting(), "1", vbBinaryCompare) = 0 Then%>
		  <SCRIPT LANGUAGE="JavaScript"><!--
			function SetFocus(){
				var oUserName = null;
				oUserName = document.NewUserForm.userName;
				oUserName.focus()
			}
		  //--></SCRIPT>
	<%End If%>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0" <%If StrComp(GetJavaScriptSetting(), "1", vbBinaryCompare) = 0 Then%>onLoad="SetFocus()" <%End If%> >
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
			    ElseIf lValidationError <> NO_ERR Then
			        Call DisplayLoginError(sErrorHeader, sErrorMessage)
			    End If
			%>
				<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
				<FORM ACTION="newuser.asp" METHOD="POST" NAME="NewUserForm">
				<INPUT TYPE="HIDDEN" NAME="site" VALUE="<%Response.Write sChannel%>" />
					<TR>
						<TD BGCOLOR="#000000" WIDTH="11" ALIGN="LEFT" VALIGN="TOP"><IMG SRC="Images/loginUpperLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#000000" WIDTH="300" ALIGN="LEFT" VALIGN="MIDDLE"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>" COLOR="#FFFFFF"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(372) 'Descriptor: Create a new account%><%Else%><%Response.Write UCase(asDescriptors(372)) 'Descriptor: Create a new account%><%End If%></B></FONT></TD>
						<TD BGCOLOR="#000000" WIDTH="11" ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="Images/loginUpperRightCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD COLSPAN="3" HEIGHT="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="2" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD BGCOLOR="#CCCCCC" WIDTH="11"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>

						  <TD BGCOLOR="#CCCCCC" WIDTH="300" ALIGN="LEFT" VALIGN="TOP">

							<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2" WIDTH="100%">
								<TR>
								    <TD WIDTH="1%"></TD>
								    <TD WIDTH="99%">
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								            <% Response.write asDescriptors(433) & " " & Replace(asDescriptors(434), "*", "<FONT COLOR=""#cc0000"">*</FONT>") 'Descriptor: To sign up for an account, please enter the following information below. Required information is noted with a red asterisk (*).%>
								        </FONT>
								    </TD>
								</TR>
								<TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT=""></TD></TR>
								<TR>
								    <TD VALIGN="TOP"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_LARGE_FONT)%>"><B><%=nStepCount%><%nStepCount = nStepCount + 1%></B></FONT></TD>
								    <TD>
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								            <B><%Response.Write asDescriptors(435) 'Descriptor: Please enter a user name and password for your new account.  Make sure to enter your password twice.%></B>
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
                                                            Response.Write asDescriptors(386) & "<BR />" 'Descriptor: Either the user name or password was blank.  Please enter them again.
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
										<%Response.Write asDescriptors(369)'Descriptor: User name:%><FONT COLOR="#cc0000">*</FONT><BR />

										<INPUT TYPE="TEXT" NAME="userName" SIZE="25" MAXLENGTH="250" VALUE="<%If Len(sUserName) > 0 Then Response.Write Server.HTMLEncode(sUserName) End If%>" STYLE="font-family: courier" /><BR />
										<%Response.Write asDescriptors(370)'Descriptor: Password:%><FONT COLOR="#cc0000">*</FONT><BR />
										<INPUT TYPE="PASSWORD" NAME="Pwd" SIZE="25" MAXLENGTH="250" STYLE="font-family: courier" /><BR />
										<%Response.Write asDescriptors(394) 'Descriptor: Confirm password:%><FONT COLOR="#cc0000">*</FONT><BR />
										<INPUT TYPE="PASSWORD" NAME="confirmPwd" SIZE="25" MAXLENGTH="250" STYLE="font-family: courier" />
										</FONT>
								    </TD>
								</TR>
								<TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT=""></TD></TR>
								<TR>
								    <TD VALIGN="TOP"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_LARGE_FONT)%>"><B><%=nStepCount%><%nStepCount = nStepCount + 1%></B></FONT></TD>
								    <TD>
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								            <B><%Response.Write asDescriptors(436) & " " 'Descriptor: Enter a hint as closely associated to your password as possible.%> <%Response.Write asDescriptors(780) 'Descriptor: If you forget your password, this hint will be displayed and upon seeing this hint again you should be reminded of your password.%></B>
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
                                            <%Response.Write asDescriptors(395) 'Descriptor: Password hint:%><FONT COLOR="#cc0000">*</FONT><BR />
										    <INPUT TYPE="TEXT" NAME="Hint" SIZE="25" MAXLENGTH="250" VALUE="<%If Len(sHint) > 0 Then Response.Write Server.HTMLEncode(sHint) End If%>" STYLE="font-family: courier" />
										</FONT>
								    </TD>
								</TR>

								<%If Len(Application("Default_Device_Name")) > 0 Then %>
								<TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT=""></TD></TR>
								<TR>
								    <TD VALIGN="TOP"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_LARGE_FONT)%>"><B><%=nStepCount%><%nStepCount = nStepCount + 1%></B></FONT></TD>
								    <TD VALIGN="TOP">
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								            <B><%Response.Write asDescriptors(633) 'Descriptor: Please enter a default address.%></B>
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
                                                        If (lValidationError And ERR_DEFAULT_ADDRESS_INVALID) Then
                                                            Select Case CStr(Application("Device_Validation"))
                                                                Case S_DEVICE_VALIDATION_EMAIL
                                                                    Response.Write asDescriptors(419) 'Descriptor: Please enter an address in the form of: user@server.com
                                                                Case S_DEVICE_VALIDATION_NUMBER
                                                                    Response.Write asDescriptors(614) 'Descriptor: Please enter a value for the address in the following form: any numbers and the following characters - ( )
                                                                Case S_DEVICE_VALIDATION_NONE
                                                                    Response.Write asDescriptors(635) 'Descriptor: Please enter an address in the following form: any text or numeric characters
                                                                Case Else
                                                            End Select
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
                                            <%
                                                Select Case CStr(Application("Device_Validation"))
                                                    Case S_DEVICE_VALIDATION_EMAIL
                                                        Response.Write asDescriptors(504) 'Descriptor: Format: xxxx@xxxxxx.xxx
                                                    Case S_DEVICE_VALIDATION_NUMBER
                                                        Response.Write asDescriptors(613) 'Descriptor: Format: any numbers and the following characters - ( )
                                                    Case S_DEVICE_VALIDATION_NONE
                                                        Response.Write asDescriptors(634) 'Descriptor: Format: any text or numeric characters
                                                    Case Else
                                                End Select
                                            %>
                                            <BR />
										    <INPUT TYPE="TEXT" NAME="defEmail" SIZE="25" MAXLENGTH="250" VALUE="<%If Len(sDefEmail) > 0 Then Response.Write sDefEmail End If%>" STYLE="font-family: courier" />
										    <BR /><%Response.Write "(" & asDescriptors(510) & ": " & Application("Default_Device_Name") & ")" 'Descriptor: Style%>
										</FONT>
								    </TD>
								</TR>
								<%End If%>

								<%If (bHasLocales) Then %>
								<TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT=""></TD></TR>
								<TR>
								    <TD VALIGN="TOP"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_LARGE_FONT)%>"><B><%=nStepCount%><%nStepCount = nStepCount + 1%></B></FONT></TD>
								    <TD VALIGN="TOP">
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								            <B><%Response.Write asDescriptors(464) 'Descriptor: Please choose a language.%></B>
								        </FONT>
								    </TD>
								</TR>
								<TR>
								    <TD></TD>
								    <TD>
                                        <% Call RenderLocaleChoices(sGetLocalesForSiteXML, "", "", GetSiteLocale()) %>
								    </TD>
								</TR>
								<%Else %>
								<TR><TD COLSPAN="2"><INPUT TYPE="HIDDEN" NAME="Locale" VALUE="<%=SYSTEM_LOCALE_ID & ";" & CStr(ENGLISH_US)%>"/></TD></TR>
								<%End If%>

								<% If (bHasProjects) Then %>
								    <TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT=""></TD></TR>
	    	                        <TR>
	    	                            <TD VALIGN="TOP"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_LARGE_FONT)%>"><B><%=nStepCount%><%nStepCount = nStepCount + 1%></B></FONT></TD>
	    	                            <TD VALIGN="TOP">
	    	                                <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.write aFontInfo(N_SMALL_FONT)%>">
	    	                                    <B><%Response.Write asDescriptors(438) ''Descriptor: Please enter your login credentials for the following Information Source(s).%></B>
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
								    <%End If%>
                                    <%Call RenderInformationSourceLogins(sGetInformationSourcesForSiteXML, "", true)%>
                                <% End If %>

								<TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT=""></TD></TR>
								<TR>
								    <TD VALIGN="TOP"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_LARGE_FONT)%>"><B><%=nStepCount%><%nStepCount = nStepCount + 1%></B></FONT></TD>
								    <TD VALIGN="TOP">
								        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								            <B><%Response.Write asDescriptors(437) 'Descriptor: You're done! Click the "Create a new account" button below.%></B>
								        </FONT>
								    </TD>
								</TR>
                                <TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT=""></TD></TR>
                                <!--
                                We need to remove the save password feature
                                <TR>
                                    <TD></TD>
                                    <TD>
										<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
											<TR>
												<TD><INPUT TYPE="CHECKBOX" NAME="SavePwd" VALUE="1" <%If StrComp(sSavePwd, "1", vbBinaryCompare) = 0 Then Response.Write "CHECKED" End If%> /></TD>
												<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(12)'Descriptor: Save my password%></FONT></TD>
											</TR>
										</TABLE>
                                    </TD>
                                </TR>
                                -->
                                <TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT=""></TD></TR>
								<TR>
								    <TD></TD>
								    <TD>
                                        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
									    	<TR>
									    		<TD><INPUT TYPE="SUBMIT" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(372) 'Descriptor: Create a new account%>" /></TD>
									    		<TD>&nbsp;</TD>
									    		<TD><INPUT CLASS="buttonClass" TYPE="SUBMIT" NAME="Cancel" VALUE="<%Response.Write asDescriptors(120)'Descriptor: Cancel%>" /></TD>
									    	</TR>
									    </TABLE>
								    </TD>
								</TR>
							</TABLE>
						</TD>
						<TD BGCOLOR="#CCCCCC" WIDTH="11"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD BGCOLOR="#CCCCCC" WIDTH="11" ALIGN="LEFT" VALIGN="BOTTOM"><IMG SRC="Images/loginLowerLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#CCCCCC" WIDTH="300"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#CCCCCC" WIDTH="11" ALIGN="RIGHT" VALIGN="BOTTOM"><IMG SRC="Images/loginLowerRightCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
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