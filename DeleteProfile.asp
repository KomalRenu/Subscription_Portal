<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/DeleteProfileCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sCacheXML
	Dim sSubGUID
	Dim sQOID
	Dim sPrefID
	Dim sISID

	sServicesLinkColor = "#cc0000"
	aPageInfo(S_NAME_PAGE) = "DeleteProfile.asp"

	lErr = ParseRequestForDeleteProfile(oRequest, sSubGUID, sQOID, sPrefID)

	If Len(oRequest("1")) > 0 Then
		Response.Redirect "prompt.asp?subGUID=" & sSubGUID & "&qoid=" & sQOID
	End If

	If lErr = NO_ERR Then
	    lErr = ReadCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
	    If lErr = NO_ERR Then
			lErr = ParseInfoFromCache(sCacheXML, sQOID, sISID)
			If lErr = NO_ERR Then
				lErr = cu_DeleteProfile(sPrefID, sQOID, sISID)
				If lErr = HYDRA_APIERROR_DELETE_PROFILE Then

				ElseIf lErr = NO_ERR Then
				    lErr = UpdateCacheXML_DeleteProfile(sCacheXML, sQOID, sPrefID)
					If lErr = NO_ERR Then
					    lErr = WriteCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
						If lErr = NO_ERR Then
							Response.Redirect "prompt.asp?subGUID=" & sSubGUID & "&qoid=" & sQOID
						End If
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
	<TITLE></TITLE>
</HEAD>
<BODY BGCOLOR="ffffff" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
<!-- #include file="header_multi.asp" -->
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
	<TR>
		<TD WIDTH="1%" valign="TOP">
			<!-- begin search box -->
				<!-- #include file="searchbox.asp" -->
			<!-- end search box -->
			<!-- begin left menu -->
			<BR>

			<BR>
			<!-- end left menu -->
			<img src="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" valign="TOP">
			<!-- begin center panel -->
			<BR>
			<%If lErr = HYDRA_APIERROR_DELETE_PROFILE Then
				Call DisplayLoginError(sErrorHeader, sErrorMessage)
				Response.write "<A HREF=""prompt.asp?subGUID=" & sSubGUID & "&qoid=" & sQOID & """>" & asDescriptors(149) & "</A>"	'Descriptor: Back
			Else
				Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(383), "prompt.asp") 'Descriptor: Back to Services
			End If %>
			<!-- end center panel -->
		</TD>
		<TD WIDTH="1%">
			<img src="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
	</TR>
</TABLE>
<BR>
<!-- begin footer -->
	<!-- #include file="footer.asp" -->
<!-- end footer -->
</BODY>
</HTML>