<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	On Error Resume Next
%>
<!-- #include file="../CustomLib/CommonLib.asp" -->
<%
Dim lStatus
Dim aConnectionInfo(13)  'MAX_CONNECTION_INFO

    'We need to call this function:
    Call SetApplicationlVariables()

    'Check if this site has already been configured, if not send to welcome.asp:
    lStatus = checkSiteConfiguration()
    If lStatus = CONFIG_OK Then
        Response.Redirect "adminOverview.asp?section=3"
    ElseIf (lStatus And CONFIG_MISSING_ENGINE) > 0 Then
        Response.Redirect "welcome.asp"
    ElseIf (lStatus And CONFIG_MISSING_MD) > 0 Then
        Response.Redirect "select_md.asp"
    ElseIf (lStatus And CONFIG_MISSING_SITE) > 0 Then
        Response.Redirect "select_site.asp"
    ElseIf (lStatus And CONFIG_MISSING_AUREP) > 0 Then
        Response.Redirect "select_aurep.asp"
    ElseIf (lStatus And CONFIG_MISSING_SBREP) > 0 Then
        Response.Redirect "select_sbrep.asp"
    Else
        Response.Redirect "welcome.asp"
    End If

%>
