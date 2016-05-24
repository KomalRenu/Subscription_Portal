<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
%>
<!-- #include file="CustomLib/ProcessUserDetailsCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->

<%
	On Error Resume Next
	Dim sAXML
	Dim sXML
	Dim tempXML
	Dim ansXML

	Dim tempDict
	Dim tDict

	Dim result
	Dim qoID
	Dim ISID
	Dim prefID

	Dim ISDef
	Dim qDef
	Dim oDecoder
	Dim i
	Dim bSaved
	Dim sOptSection

	Dim sAccount

	bSaved = False
	sOptSection = "5"
	If not Request is nothing then
		If Request("storeUDAns") <> "" Then
			If Request.Form.Count > 0 then

				qoID = Request("qoID")
				ISID = Request("ISID")
				qDef = Request("qDef")

				'tempXML = "<UserDetail><First_Name Name='First Name' Default='-1'/><Last_Name Name='Last Name' Default='-1'/><Middle_Initial Name='Middle Initial' Default='-1'/><Suffix Name='Suffix' Default='-1'/><Salutation Name='Salutation' Default='-1'/><Zip_Code Name='Zip Code' Default='-1'/><Title Name='Title' Default='-1'><Option>Mr.</Option><Option>Mrs.</Option><Option>Ms.</Option></Title></UserDetail>"
				Set tDict = Server.CreateObject("Scripting.Dictionary")

				For i=1 to Request.Form.Count
					'Response.write Request.Form.Item(i)
					Dim key
					Dim item
					key = CStr(Request.Form.key(i))
					item = CStr(Request.Form.Item(key))
					call tDict.add(key,item)
				Next

				ansXML = generateAnswer(tDict,qDef)

				If Err.number <> 0 then
					lErr = Err.Number
				Else
					If Request("edit") <> "" then
						prefID = Request("prefID")
						If prefID <> "" then
							lErr = SavePreferenceAnswer(qoID,ISID,prefID,ansXML)
						end if
					else
						lErr = SaveProfileAndPreferenceAnswer(qoID,ISID,ansXML)
					end if
				End if

				If lErr = 0 Then
					bSaved = True
					sAccount = Request("redirect")
					If sAccount = "new" Then
						Response.Redirect "default.asp?account=new"
					ElseIf sAccount = "iserver" Then
						Response.Redirect "home.asp"
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
	<TITLE><%Response.Write "Define User Details"'Descriptor: Change my password%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
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
			<%
			If lErr <> NO_ERR Then
				Call DisplayLoginError(sErrorHeader, sErrorMessage)
			ElseIf lValidationError <> NO_ERR Then
			    Call DisplayLoginError(sErrorHeader, sErrorMessage)
			End If
			%>

				<BR />
				<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
					<TR>
						<TD BGCOLOR="#000000" WIDTH="11" ALIGN="LEFT" VALIGN="TOP"><IMG SRC="Images/loginUpperLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#000000" WIDTH="200" ALIGN="LEFT" VALIGN="MIDDLE"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>" COLOR="#FFFFFF"><B><%Response.Write asDescriptors(372) 'Descriptor: Create a new account%></B></FONT></TD>
						<TD BGCOLOR="#000000" WIDTH="11" ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="Images/loginUpperRightCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD COLSPAN="3" HEIGHT="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="2" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD BGCOLOR="#CCCCCC" WIDTH="11"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#CCCCCC" WIDTH="200" ALIGN="LEFT" VALIGN="TOP">


				<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2" WIDTH="100%">
				<TR><TD>

				<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
					<%
								Response.Write "<BR>"
								qoID = "56D425DD94B211D4BE6600C04F0E93B7"
								ISID = "56D4246394B211D4BE6600C04F0E93B7"
								call GetISDefinition(ISID,sXML)
								If qoID <> "" and sXML <> "" Then
									'Create a dictionary object for holding the answers
									lErr = GetUserDetailPreferenceDefinition(qoID,prefID, sAXML)
									If lErr = 0 Then
										Response.Write "<FORM name=" & """" & "postForm" & """" & " method="  & """" & "POST" & """" & " action=" & """" & "newUserDetails.asp" & """" & ">"
										Response.write transFormDefinition(qoID,ISID,sXML,prefID,sAXML)
										If Request("account") <> "" Then Response.Write "<Input" & " name=" & """" & "redirect" & """" &  " type=" & """" & "hidden" & """" & " value=" & """" & Request("account") & """" & ">"
										Response.Write "</FORM>"
									End If
								End If
					%>
				</FONT>
				</TD>
				</TR>
				</TABLE>
						</TD>
						<TD BGCOLOR="#CCCCCC" WIDTH="11"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
					</TR>
					<TR>
						<TD BGCOLOR="#CCCCCC" WIDTH="11" ALIGN="LEFT" VALIGN="BOTTOM"><IMG SRC="Images/loginLowerLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#CCCCCC" WIDTH="200"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD BGCOLOR="#CCCCCC" WIDTH="11" ALIGN="RIGHT" VALIGN="BOTTOM"><IMG SRC="Images/loginLowerRightCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
					</TR>
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
