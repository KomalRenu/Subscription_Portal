<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
%>
<!-- #include file="CustomLib/ProcessUserDetailsCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<!-- #include file="CustomLib/LoginCuLib.asp" -->

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
				End If

			End If
		End If
	End If

%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(971) 'Descriptor: Define User Details %> - MicroStrategy Narrowcast Server</TITLE>
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
			            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(26) & " " 'Descriptor: You are here:%> <%Response.Write asDescriptors(286) 'Descriptor: Preferences%> > <B><%Response.Write asDescriptors(971) 'Descriptor: Define User Details %></B></FONT>
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
							<TR>
								<TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
							</TR>
							<TR>
								<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
							</TR>
							<TR>
								<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(403) 'Descriptor: Your password has been changed.%></B></FONT></TD>
							</TR>
						</TABLE>
					<BR />
				<%	End If%>

				<TABLE BGCOLOR="#CCCCCC" BORDER="0" WIDTH="75%" CELLSPACING="0" CELLPADDING="0">
				    <TR>
					    <TD ALIGN="LEFT" VALIGN="TOP"><IMG SRC="Images/loginUpperLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD WIDTH="100%"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(970) 'Descriptor: Please provide answers for the following user details fields %></B></FONT></TD>
						<TD ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="Images/loginUpperRightCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
                    </TR>
				</TABLE>

				<TABLE BGCOLOR="#CCCCCC" BORDER="0" WIDTH="75%" CELLSPACING="0" CELLPADDING="1">
				    <TR><TD COLSPAN="3">
					        <TABLE BGCOLOR="#FFFFFF" BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="3"><TR>
							    <TD WIDTH="1%">&nbsp;&nbsp;</TD>
								<TD><BR />
                                    <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
                                    <TR><TD colspan="3">


					<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
					<%
								qoID = "56D425DD94B211D4BE6600C04F0E93B7"
								ISID = "56D4246394B211D4BE6600C04F0E93B7"
								'sXML = "<UserDetail><First_Name Name='First Name' Default='-1'/><Last_Name Name='Last Name' Default='-1'/><Middle_Initial Name='Middle Initial' Default='-1'/><Suffix Name='Suffix' Default='-1'/><Salutation Name='Salutation' Default='-1'/><Zip_Code Name='Zip Code' Default='-1'/><Title Name='Title' Default='-1'><Option>Mr.</Option><Option>Mrs.</Option><Option>Ms.</Option></Title></UserDetail>"
								call GetISDefinition(ISID,sXML)
								If qoID <> "" and sXML <> "" Then
									'Create a dictionary object for holding the answers
									lErr = GetUserDetailPreferenceDefinition(qoID,prefID, sAXML)
									If lErr = 0 Then
										Response.Write "<FORM name=" & """" & "postForm" & """" & " method="  & """" & "POST" & """" & " action=" & """" & "ProcessUserDetails.asp" & """" & ">"
										Response.write transFormDefinition(qoID,ISID,sXML,prefID,sAXML)
										Response.Write "</FORM>"
									End If
								End If
					%>


                                    </TABLE><BR />
								</TD>
                            </TR></TABLE>
					</TD></TR>
					<TR><TD COLSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
					<TR><TD COLSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
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
