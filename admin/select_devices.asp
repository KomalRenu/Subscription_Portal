<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
    Option Explicit
    Response.CacheControl = "no-cache"
    Response.AddHeader "Pragma", "no-cache"
    Response.Expires = -1
    On Error Resume Next
%>
<!-- #include file="../CustomLib/SiteConfigCuLib.asp" -->
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<%
Dim sDevice
Dim sName
Dim sId
Dim sFolderId
Dim sFolderContentXML
Dim lStatus


    'Check for actions cancelled:
    If oRequest("action") = "Cancel" Then
        Response.Redirect("devices_config.asp")
    End If

    'Read rest of request variables:
    sDevice = oRequest("device")
    sName = oRequest("n")
    sId = oRequest("id")
    sFolderId = oRequest("fid")

    If sDevice = "portal" Then
        aPageInfo(S_TITLE_PAGE) = STEP_SELECT_DEVICE & " " & asDescriptors(594) 'Descriptor:Select Portal Device
    Else
        aPageInfo(S_TITLE_PAGE) = STEP_SELECT_DEVICE & " " & asDescriptors(595) 'Descriptor:Select Default Device
    End If

    'Get the Channels list request from the request object:
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = "device=" & sDevice & "&id=" & sId & "&n=" & sName

    lStatus = checkSiteConfiguration()

    'Get the previous values:
    If lErr = NO_ERR Then
        lErr = co_getFolderXML(sFolderId, ROOT_DEVICE_FOLDER_TYPE, sFolderContentXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "select_devices.asp", "", "", "Error calling co_getFolderXML", LogLevelTrace)
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
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(579) , "devices_config.asp") 'Descriptor: Return to: 'Descriptor:Site Preferences %>
      <%Else%>
        <BR />
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
          <%Call Response.Write(asDescriptors(675)) 'Click on the appropriate device to change the selection.%>
        </FONT>
        <BR />
        <BR />

        <%Call RenderDevicesFolderPath(sFolderId, sId, sDevice, sFolderContentXML) %>
        <BR>

        <BR>
        <%Call RenderList_Devices(sFolderId, sId, sDevice, sFolderContentXML)%>

      <%End If %>
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="1%">
        <!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>