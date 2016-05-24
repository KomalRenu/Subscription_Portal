<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="../CoreLib/ISCoLib.asp" -->
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<%
Dim sOrigin
Dim nCount
Dim oItem
Dim aInformationSources()
Dim i
Dim lStatus

    'Check for actions cancelled:
    If oRequest("back") <> "" Then
		Erase aInformationSources
        Response.Redirect("devices_config.asp")
    End If

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "is_config.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_IS & " " & asDescriptors(568) 'Descriptor:Information Sources
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3

    lStatus = checkSiteConfiguration()

    'Assume that the rest of the arguments would be IS:
    Redim aInformationSources(oRequest.Count - 1, 1)

    i = 0
    For Each oItem in oRequest
        If Left(oItem, 3) = "is." Then
            aInformationSources(i, 0) = Mid(oItem, 4)
            aInformationSources(i, 1) = oRequest(oItem)
            i= i + 1
        End If
    Next

    'If no IS, cancel:
    If i = 0 Then
		Erase aInformationSources
		Set oItem = Nothing

        Response.Redirect("preferences.asp")
    End If

    If lErr = NO_ERR Then
        lErr = setInfSources(aInformationSources)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "modify_is.asp", "", "", "Error calling setInfSources", LogLevelTrace)
    End If

    If lErr = NO_ERR Then
        Call ResetApplicationVariables()
        Erase aInformationSources
		Set oItem = Nothing

        Call Response.Redirect("preferences.asp")
    End If


%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<!-- #include file="../NSStyleSheet.asp" -->

</HEAD>
<BODY BGCOLOR="FFFFFF" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
  <TR>
    <TD COLSPAN="6" HEIGHT="1%">
      <!-- begin header -->
        <!-- #include file="admin_header.asp" -->
      <!-- end header -->
    </TD>
  </TR>
  <TR>
    <TD WIDTH="1%" valign="TOP">
      <!-- begin toolbar -->
        <!-- #include file="_toolbar_site_preferences.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(568) , "is_config.asp") 'Descriptor: Return to: 'Descriptor:Information Sources %>
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="1%">
        <!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
	Erase aInformationSources
	Set oItem = Nothing
%>