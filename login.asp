<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>

<!-- #include file="CommonDeclarations.asp" -->
<!-- #include file="CustomLib/LoginCuLib.asp" -->
<!-- #include file="CustomLib/AddressesCuLib.asp" -->
<!-- #include file="CoreLib/ISLoginCoLib.asp" -->
<%
    'Check if this site has already been configured, if not send to welcome.asp:
    lErr = checkSiteConfiguration()
    If lErr <> NO_ERR Then
        Response.Redirect "welcome.asp"
    End If

	'Check if user is logged in.  If so, send user to home page.
	If Len(LoggedInStatus()) > 0 Then
		Response.Redirect DefaultPage()
	End If

	Dim sGetUserAddressesXML
	Dim bHasPortalAddress
	Dim bEncryptedFlag
	Dim bLoginFromSource
	Dim sUserName
	Dim sPassword
	Dim sSavePwd
	Dim sSessionID
	Dim sStatus
	Dim sNTUser
	Dim sNCSUserName
	Dim sUserID
	Dim sAuthMode

	Dim LoginMode
	Dim IServerName
	Dim IServerPort
	Dim IServerUser
	Dim sISSessionID
	Dim strErrDesc
	Dim UserInfo()
	dim sessionObj
	dim bLocalOk
	Dim bNewUser

	LoginMode = Application.Value("Login_Mode")

	'For consistency with old sites created with 7.1GA
	If LoginMode = "" Then
		LoginMode = "NC_NORMAL"
	End IF

	IServerName = GetIserverName()
	IServerPort = GetIserverPort()

	bHasPortalAddress = False
	bEncryptedFlag = False
	bLoginFromSource = False
	bLocalOk = False

	lErr = ParseRequestForLogin(oRequest, sUserName, sPassword, sSavePwd, sStatus, sNTUser)

	If lErr = NO_ERR Then
	    If Len(sStatus) = 0 Then
	        If StrComp(GetSavePasswordSetting(), "1", vbBinaryCompare) = 0 Then
	            If Len(sUserName) = 0 Then
	                sUserName = GetUsername()
                    aSourceInfo(0) = USER_COOKIE
                    aSourceInfo(1) = USER_PASSWORD
                    sPassword = ReadFromSource(aConnectionInfo, SOURCE_COOKIES, aSourceInfo)
                    bEncryptedFlag = True
                    sSavePwd = "1"
                    bLoginFromSource = True
                End If
	        End If
	    End If

		If (Request.Form.Count > 0) Or (bLoginFromSource = True) Or (StrComp(sNTUser, "yes", vbBinaryCompare) = 0) Then

			If (Left(LoginMode,2) = "NC" And StrComp(Left(sUserName,6), "local\", vbTextCompare) = 0 ) Then
				bLocalOk = true
			End If

			'Code to be executed for normal NC logins.
			If (LoginMode = "NC_NORMAL" OR bLocalOk ) Then

				If StrComp(Left(sUserName,6), "local\", vbTextCompare) = 0 Then
					sUserName = Mid(sUserName,7)
				End If

				IServerUser =	false
				Call SetIServerCookies("","",IServerUser, "no")

				lErr = cu_CreateSession(sUserName, sPassword, bEncryptedFlag, sSessionID)
				If lErr = NO_ERR Then
				    Call SetSessionID(sSavePwd, sSessionID)
				    'Error handling?
				    lErr = cu_GetUserAddresses(sGetUserAddressesXML)
				    If lErr = NO_ERR Then
				        lErr = CheckForPortalAddress(sGetUserAddressesXML, sUserName)
				    End If
				End If

				If lErr = NO_ERR Then
				    lErr = SetSessionInfo(sChannel, sSessionID, sUserName, sSavePwd, "", "")
				    If lErr = NO_ERR Then
				        Response.Redirect DefaultPage()
				    End If
				End If

				If lErr <> NO_ERR And sSessionID <> "" Then
				   	Call cu_CloseSession()
				    Call Logout()
				End If
			'Code to be executed for mandatory IS logins.
			ElseIf (LoginMode = "IS_NORMAL" OR LoginMode = "NT_NORMAL" OR LoginMode = "IS_NT_NORMAL" ) Then
				'Try loggin in to ISERVER
				lErr = GetSessionObj(sessionObj, strErrDesc)

				If lErr = NO_ERR Then
					If StrComp(sNTUser, "yes", vbBinaryCompare) = 0 Then
						sAuthMode = "2"
					Else
						sAuthMode = "1"
					End If
					Session("AuthMode") = sAuthMode

					lErr = getUserSession(sessionObj,IServerName,IServerPort, sUserName,sPassword,sAuthMode,sISSessionID,strErrDesc)
				End If

				If lErr = NO_ERR Then
					IServerUser = true
					lErr = GetUserInfo(sessionObj,sISSessionID,UserInfo,strErrDesc)
					Call SetIServerCookies(sUserName, sPassword, IServerUser, sNTUser)
					sNCSUserName = UserInfo(1) & "(" & UserInfo(0) & ")"
					sUserID = UserInfo(0)
					Call closeSession(sessionObj, sISSessionID)
					Session("CastorUserID")=sUserID
				End If

				If lErr = NO_ERR And sISSessionID <> "" Then
					lErr = cu_CreateSession(sNCSUserName, sUserID, bEncryptedFlag, sSessionID)

					'If user does not exist in hydra then create new hydra user if allowed

					If (lErr = ERR_LOGIN_ERROR And StrComp(CStr(Application("Allow_New_users")), "1", vbBinaryCompare) = 0) Then
						lErr = ProcessCreateNewUser(sNCSUserName, sUserName, sUserID, sPassword, sAuthMode, sSessionID)
						bNewUser = true
					Else
						If lErr = NO_ERR And sAuthMode <> "2" Then
							lErr = UpdateAuthenticationObject(sSessionID, sUserName, sUserID, sPassword, sAuthMode)
						End If
					End If

					If lErr = NO_ERR Then
					    Call SetSessionID(sSavePwd, sSessionID)
					    'Error handling?
					    lErr = cu_GetUserAddresses(sGetUserAddressesXML)
					    If lErr = NO_ERR Then
					        lErr = CheckForPortalAddress(sGetUserAddressesXML, sNCSUserName)
					    End If
					End If

					If lErr = NO_ERR Then
					    lErr = SetSessionInfo(sChannel, sSessionID, sNCSUserName, sSavePwd, "", "")
					    If lErr = NO_ERR Then
					    	If bNewUser Then
					    		Response.Redirect "newUserDetails.asp?account=iserver"
					    	Else
					        	Response.Redirect DefaultPage()
					    	End If
					    End If

					End If

					If lErr <> NO_ERR And sSessionID <> "" Then
					   	Call cu_CloseSession()
					    Call Logout()
					End If
				End If
			'Code to be executed for either IS logins or NC logins if IS login not successfull.
			ElseIf (LoginMode = "NC_IS_NORMAL" OR LoginMode = "NC_NT_NORMAL" OR LoginMode = "NC_IS_NT_NORMAL") Then
				'Try loggin in to ISERVER
				lErr = GetSessionObj(sessionObj, strErrDesc)

				If lErr = NO_ERR Then
					If StrComp(sNTUser, "yes", vbBinaryCompare) = 0 Then
						sAuthMode = "2"
					ElseIF (LoginMode <> "NC_NT_NORMAL") Then
						sAuthMode = "1"
					End IF
					Session("AuthMode") = sAuthMode

					lErr = getUserSession(sessionObj,IServerName,IServerPort, sUserName,sPassword,sAuthMode,sISSessionID,strErrDesc)
				End If

				'If user clicked on NT link and failed validation, do not even try to create a NC session
				If StrComp(sNTUser, "yes", vbBinaryCompare) = 0 And lErr <> NO_ERR Then
					'Do Nothing
				Else
					If lErr = NO_ERR And sISSessionID <> "" Then
						IServerUser = true
						lErr = GetUserInfo(sessionObj,sISSessionID,UserInfo,strErrDesc)
						Call SetIServerCookies(sUserName, sPassword, IServerUser, sNTUser)
						sNCSUserName = UserInfo(1) & "(" & UserInfo(0) & ")"
						sUserID = UserInfo(0)
						Call closeSession(sessionObj, sISSessionID)
						Session("CastorUserID")=sUserID
						lErr = cu_CreateSession(sNCSUserName, sUserID, bEncryptedFlag, sSessionID)
					Else
						IServerUser = false
						Call SetIServerCookies(sUserName, sPassword, IServerUser, sNTUser)
						lErr = cu_CreateSession(sUserName, sPassword, bEncryptedFlag, sSessionID)
						sNCSUserName = sUserName
					End If


					'If user does not exist in hydra then create new hydra user if allowed
					If (lErr = ERR_LOGIN_ERROR And IServerUser And StrComp(CStr(Application("Allow_New_users")), "1", vbBinaryCompare) = 0) Then
						lErr = ProcessCreateNewUser(sNCSUserName, sUserName, sUserID, sPassword, sAuthMode, sSessionID)
						bNewUser = true
					Else
						If lErr = NO_ERR And IServerUser and sAuthMode <> "2" Then
							lErr = UpdateAuthenticationObject(sSessionID, sUserName, sUserID, sPassword, sAuthMode)
						End If
					End If

					If lErr = NO_ERR Then
					    Call SetSessionID(sSavePwd, sSessionID)
					    'Error handling?
					    lErr = cu_GetUserAddresses(sGetUserAddressesXML)
					    If lErr = NO_ERR Then
					        lErr = CheckForPortalAddress(sGetUserAddressesXML, sNCSUserName)
					    End If
					End If

					If lErr = NO_ERR Then
					    lErr = SetSessionInfo(sChannel, sSessionID, sNCSUserName, sSavePwd, "", "")
					    If lErr = NO_ERR Then
					    	If bNewUser Then
					    		Response.Redirect "newUserDetails.asp?account=iserver"
					    	Else
					        	Response.Redirect DefaultPage()
					    	End If
					    End If
					End If

					If lErr <> NO_ERR And sSessionID <> "" Then
					   	Call cu_CloseSession()
					    Call Logout()
					End If
				End IF
			End If

		End If

	End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(15)'Descriptor: Login%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="login_header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin left menu -->
			    <!-- #include file="_toolbar_Login.asp" -->
			<BR />
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" VALIGN="TOP">
			<!-- begin center panel -->
			<BR />
			<%
			    If lErr <> NO_ERR Then
			        Call DisplayLoginError(sErrorHeader, sErrorMessage)
			    ElseIf StrComp(sStatus, "logout", vbBinaryCompare) = 0 Then
			        Call DisplayLoginError(asDescriptors(152), asDescriptors(421)) 'Descriptors: You have been logged out., Thank you for using MicroStrategy Narrowcast Server.
			    ElseIf StrComp(sStatus, "timeout", vbBinaryCompare) = 0 Then
			        Call DisplayLoginError (asDescriptors(152), asDescriptors(459)) 'Descriptors: You have been logged out., User session exceeded the maximum idle time. Please login again.
			    ElseIf StrComp(sStatus, "deact", vbBinaryCompare) = 0 Then
			        Call DisplayLoginError(asDescriptors(461), asDescriptors(421)) 'Descriptor: Your account has been deactivated., Thank you for using MicroStrategy Narrowcast Server.
			    End If
			%>
				<TABLE WIDTH="400" BORDER="0" CELLSPACING="0" CELLPADDING="0">
					<FORM ACTION="login.asp" METHOD="POST">
					<INPUT TYPE="HIDDEN" NAME="site" VALUE="<%Response.Write sChannel%>" />
					<TR>
						<TD BGCOLOR="#000000" WIDTH="11" ALIGN="LEFT" VALIGN="TOP"><IMG SRC="Images/loginUpperLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#000000" WIDTH="237" ALIGN="LEFT" VALIGN="MIDDLE"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>" COLOR="#FFFFFF"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(15)'Descriptor: Login%><%Else%><%Response.Write UCase(asDescriptors(15))'Descriptor: Login%><%End If%></B></FONT></TD>
						<TD BGCOLOR="#000000" WIDTH="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<%If StrComp(CStr(Application("Allow_New_users")), "1", vbBinaryCompare) = 0 Then%>
    						<TD BGCOLOR="#000000" WIDTH="139" ALIGN="LEFT" VALIGN="MIDDLE"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>" COLOR="#FFFFFF">&nbsp;<B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(368)'Descriptor: New Users%><%Else%><%Response.Write UCase(asDescriptors(368))'Descriptor: New Users%><%End If%></B></FONT></TD>
                        <%Else%>
                            <TD BGCOLOR="#000000" WIDTH="139" ALIGN="LEFT" VALIGN="MIDDLE"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
                        <%End If%>
						<TD BGCOLOR="#000000" WIDTH="11" ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="Images/loginUpperRightCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD COLSPAN="5" HEIGHT="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="2" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD BGCOLOR="#CCCCCC" WIDTH="11"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>

						  <TD BGCOLOR="#CCCCCC" WIDTH="237" ALIGN="LEFT" VALIGN="TOP">

							<TABLE BORDER="0" CELLSPACING="1" CELLPADDING="5">

								<TR><TD>
								<% If LoginMode <> "NT_NORMAL" Then %>
									<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
										<%Response.Write asDescriptors(369)'Descriptor: User name:%><BR />

										<INPUT TYPE="TEXT" NAME="userName" SIZE="25" MAXLENGTH="250" STYLE="font-family: courier"  /><BR />
										<%Response.Write asDescriptors(370)'Descriptor: Password:%><BR />
										<INPUT TYPE="PASSWORD" NAME="Pwd" SIZE="25" MAXLENGTH="250" STYLE="font-family: courier" />
										<BR /><BR />
									</FONT>
									<!--
									Comment Out the Save Password Dialog
									<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
										<TR>
											<TD><INPUT TYPE="CHECKBOX" NAME="SavePwd" VALUE="1" /></TD>
											<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(12)'Descriptor: Save my password%></FONT></TD>

										</TR>
									</TABLE>
									-->
									<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%">
										<TR>
											<TD ALIGN="RIGHT" COLSPAN="3"><INPUT TYPE="SUBMIT" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(15)'Descriptor: Login%>" /></TD>
										</TR>
										<%If LoginMode = "NC_NORMAL" OR LoginMode = "NC_IS_NORMAL" OR LoginMode = "NC_NT_NORMAL" OR LoginMode = "NC_IS_NT_NORMAL" Then %>
											<TR>
												<TD COLSPAN="3"><BR /><A HREF="password_hint.asp?site=<%Response.Write sChannel%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%Response.Write asDescriptors(409)'Descriptor: Forgot your password?%></B></FONT></A></TD>
											</TR>
										<%End If%>
									</TABLE>
								<% Else %>
									<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
										<%Response.Write asDescriptors(903) 'Descriptor:You can try to log in to this site using your NT credentials. Just click the link below."%> <BR></FONT><BR />
										<A HREF="login.asp?NTUser=yes"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><B><%Response.Write asDescriptors(922)'Descriptor: Login as NT User%></B></FONT></A>
									</FONT>
								<%End If%>
								</TD></TR>
							</TABLE>
						</TD>
							<TD WIDTH="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
							<TD BGCOLOR="#BDBDBD" WIDTH="139" ALIGN="LEFT" VALIGN="TOP">
							    <%If StrComp(CStr(Application("Allow_New_users")), "1", vbBinaryCompare) = 0 Then%>
								    <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="5">
										<TR><TD>
								    	<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								    		<%Response.Write asDescriptors(281)'Descriptor: Do not have an account?%><BR />
								    		<%Response.Write asDescriptors(371)'Descriptor: You can create a new account. Just click below.%><BR />
								    	</FONT>
								    	<A HREF="newuser.asp?site=<%Response.Write sChannel%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><B><%Response.Write asDescriptors(372)'Descriptor: Create a new account%></B></FONT></A>
										</TD></TR>
								    </TABLE>
								<%Else%>
								    <IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" />
								<%End If%>

								<% If LoginMode = "NT_NORMAL" OR LoginMode = "IS_NT_NORMAL" OR LoginMode = "NC_NT_NORMAL" OR LoginMode = "NC_IS_NT_NORMAL" Then%>
									<Table><TR><TD>
										<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
											<%Response.Write asDescriptors(903) 'Descriptor:You can try to log in to this site using your NT credentials. Just click the link below."%> <BR></FONT><BR />
									    	<A HREF="login.asp?NTUser=yes"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><B><%Response.Write asDescriptors(922)%></B></FONT></A>
										</FONT>
									</TD></TR>
									</table>
								<% End If %>


							</TD>
							<TD BGCOLOR="#BDBDBD" WIDTH="11"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD BGCOLOR="#CCCCCC" WIDTH="11" ALIGN="LEFT" VALIGN="BOTTOM"><IMG SRC="Images/loginLowerLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#CCCCCC" WIDTH="237"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD WIDTH="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#BDBDBD" WIDTH="139"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
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