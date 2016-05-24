<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CommonDeclarations.asp" -->
<!-- #include file="CustomLib/ReportsCuLib.asp" -->
<%
    Dim aDocInfo(3)
    Dim sContentXML

    'Check if user is logged in.  If not, send user to login page.
    If Len(LoggedInStatus()) = 0 Then
        Response.Redirect "login.asp"
    End If

	  'Start with no errors
	  lErr = NO_ERR

    'Parse the request info:
    If lErr = NO_ERR Then
        aDocInfo(DOC_SUBS_ID) = Trim(CStr(oRequest("subsId")))
    End If

    'Request the content of a document, if any is requested:
    If lErr = NO_ERR Then
        'If no subscription, show an Error:
        If Len(aDocInfo(DOC_SUBS_ID)) = 0 Then
            lErr = URL_MISSING_PARAMETER
        Else
            lErr = GetSubscriptionContent(aDocInfo, sContentXML)
            If lErr = NO_ERR Then
                'Get the DocBody from the XML:
                lErr = GetDocBody(aDocInfo, sContentXML)
            ElseIf lErr = ERR_DOC_BODY_NOT_FOUND Then
                lErr = NO_ERR
                aDocInfo(DOC_BODY) = asDescriptors(359) 'Descriptor: Information not available
            End If
        End If
    End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(360)'Descriptor: Reports%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY BGCOLOR="FFFFFF" TOPMARGIN="0" LEFTMARGIN="10" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<%
    'If everything was ok, show the body and finish here
	If lErr = NO_ERR Then
		Response.Write aDocInfo(DOC_BODY)
		Erase aDocInfo
		Response.End
	End If

	'If not, show an error message.
%>
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
			<BR />
			<%
			    If lErr <> NO_ERR Then
			        Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(250) & " reports.asp", "reports.asp") 'Descriptor: Return to:
			    End If
			%>
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
<%
    Erase aDocInfo
%>