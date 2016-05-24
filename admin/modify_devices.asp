<%'** Copyright © 1996-2012  MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<%
Dim sDevice
Dim sName
Dim sId
Dim sFolderId
Dim aSiteProperties()
Redim aSiteProperties(MAX_SITE_PROP)
Dim lStatus

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "devices_config.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SITE_DEVICES & " " & asDescriptors(579) 'Descriptor:Site Devices
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3

    lStatus = checkSiteConfiguration()

    'Check for actions cancelled:
    If oRequest("cancel") <> "" Then
		Erase aSiteProperties
        Response.Redirect("devices_config.asp")
    End If

    'Read rest of request variables:
    sDevice = oRequest("device")
    sName = oRequest("n")
    sId = oRequest("id")
    sFolderId = oRequest("fid")

    'Get the previous values:
    If lErr = NO_ERR Then
        lErr = getSiteProperties(aSiteProperties)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "modify_devices.asp", "", "", "Error calling getSiteProperties", LogLevelTrace)
    End If

    If lErr = NO_ERR Then

        If sDevice = "portal" Then
            aSiteProperties(SITE_PROP_PORTAL_DEV_NAME) = sName
            aSiteProperties(SITE_PROP_PORTAL_DEV_ID) = sId
            aSiteProperties(SITE_PROP_PORTAL_FOLDER_ID) = sFolderId
        Else
            aSiteProperties(SITE_PROP_DEFAULT_DEV_NAME) = sName
            aSiteProperties(SITE_PROP_DEFAULT_DEV_ID) = sId
            aSiteProperties(SITE_PROP_DEFAULT_FOLDER_ID) = sFolderId
        End If

        lErr = setSiteProperties(aSiteProperties, FLAG_PROP_GROUP_DEVICES)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "modify_devices.asp", "", "", "Error calling setSiteProperties", LogLevelTrace)

    End If

    If lErr = NO_ERR Then
        Call ResetApplicationVariables()
        Erase aSiteProperties

        Call Response.Redirect("devices_config.asp")
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
      <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(579) , "devices_config.asp") 'Descriptor: Return to: 'Descriptor:Site Preferences %>
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
	Erase aSiteProperties
%>