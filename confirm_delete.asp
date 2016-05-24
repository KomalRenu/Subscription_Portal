<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Dim sQueryString
	Dim sFormPage
	Dim sCancelButton
	Dim selectedTab
	Dim sDeleteType			'1=addresses, 2=subscriptions
	Dim sServiceID
	Dim sFolderID
	Dim i
	Dim sServiceInfo
	Dim iStart
	Dim temArray

	lErr = ParseRequestForDeleteConfirm(oRequest, sFormPage, sCancelButton, sDeleteType, sServiceID, sFolderID)

	If (Not (oRequest("delSubsGUID").Count > 0)) And StrComp(sDeleteType, "2", vbBinaryCompare) = 0 Then
	    Response.Redirect "subscriptions.asp?serviceID=" & sServiceID & "&folderID=" & sFolderID
	End If
%>
<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE>MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin search box -->
				<!-- #include file="searchbox.asp" -->
			<!-- end search box -->
			<!-- begin left menu -->
			<BR /><BR />
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" VALIGN="TOP">
			<!-- begin center panel -->
			<%
			If lErr <> NO_ERR Then
				If StrComp(CStr(oRequest("deleteType")), "1", vbBinaryCompare) = 0 Then
			        Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(381), "addresses.asp") 'Descriptor: Back to Addresses
				ElseIf StrComp(CStr(oRequest("deleteType")), "2", vbBinaryCompare) = 0 Then
					Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(383), "services.asp") 'Descriptor: Back to Services
				Else
					Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(380), "home.asp") 'Descriptor: Back to Home
				End If
			Else
			%>
					<TABLE BORDER="0" CELLPADDING="3" CELLSPACING="0">
						<FORM METHOD="POST" ACTION="<%Response.Write sFormPage%>">
						<% Call PutHiddenInputsForConfirmDelete(oRequest) %>
						<TR>
							<TD>
								<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%If StrComp(sDeleteType, "1", vbBinaryCompare) = 0 Then %><%Response.Write asDescriptors(398) 'Descriptor: Confirm address deletion%><%ElseIf StrComp(sDeleteType, "2", vbBinaryCompare) = 0 Then%><%Response.Write asDescriptors(399) 'Descriptor: Confirm subscription deletion%><%End If%></B></FONT><BR /><BR />
								<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><%If StrComp(sDeleteType, "1", vbBinaryCompare) = 0 Then %><%Response.Write asDescriptors(400) 'Descriptor: Are you sure you want to delete the address?%><%ElseIf StrComp(sDeleteType, "2", vbBinaryCompare) = 0 Then%><%Response.Write asDescriptors(401) 'Descriptor: Are you sure you want to delete the subscription(s)?%><%End If%></FONT>
							</TD>
						</TR>
						<%If StrComp(sDeleteType, "2", vbBinaryCompare) = 0 Then%>
						<TR>
						    <TD>
						        <BR />
						        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>">
						        <%
						            For i = 1 to oRequest("delSubsGUID").Count
						                'sServiceInfo = CStr(oRequest("delSubsGUID")(i))
						                'iStart = InStr(1, sServiceInfo, ";")
						                'iStart = InStr(iStart+1, sServiceInfo, ";")
						                temArray = Split(CStr(oRequest("delSubsGUID")(i)), ";", -1, vbBinaryCompare)
										Response.Write "<LI>" & Server.HTMLEncode(CStr(temArray(3))) & "</LI>"
						            Next
						        %>
						        </FONT>
						    </TD>
						</TR>
						<%End If%>
						<TR>
							<TD ALIGN="RIGHT">
								<BR />
								<INPUT TYPE="SUBMIT" VALUE="<%Response.Write asDescriptors(119) 'Descriptor: Yes%>" CLASS="buttonClass" /> <INPUT TYPE="SUBMIT" NAME="<%Response.Write sCancelButton%>" VALUE="<%Response.Write asDescriptors(118) 'Descriptor: No%>" CLASS="buttonClass" />
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