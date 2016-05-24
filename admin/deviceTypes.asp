<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/DeviceTypesCuLib.asp" -->
<!-- #include file="../CommonDeclarations.asp" -->
<%
Dim sDeviceTypesXML
Dim sAction
Dim sDeviceTypeID
Dim sRenameDeviceTypeID
Dim sRenameDeviceTypeName
Dim lStatus

Dim aDeviceTypeInfo

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "deviceTypes.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_DEVICE_TYPES & " " & asDescriptors(493) 'Descriptor: Device Types
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3

    lStatus = checkSiteConfiguration()

    lErr = ParseRequestForDeviceType(oRequest, aDeviceTypeInfo)

    If lErr = NO_ERR Then
        lErr = ReadDeviceTypesXML(sDeviceTypesXML)
    End If

    If lErr = NO_ERR Then
        If Len(oRequest("RenameDT")) > 0 Then
            lErr = RenameDeviceType(aDeviceTypeInfo(DEV_TYPE_ID), aDeviceTypeInfo(DEV_TYPE_NAME), sDeviceTypesXML)
            If lErr = NO_ERR Then
                lErr = WriteDeviceTypesXML(sDeviceTypesXML)
            Else
                aDeviceTypeInfo(DEV_TYPE_ACTION) = "rename"
            End If
        End If
    End If


%>

<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>
<%If aDeviceTypeInfo(DEV_TYPE_ACTION) = "rename" Then %>
<SCRIPT LANGUAGE=javascript>
<!--

  function validateForm() {
  var sMsg

    sMsg = "";
    if ((FormDevName.DTName.value == "") || isBlank(FormDevName.DTName.value)) {
      <%Call Response.Write("sMsg += ""<LI>" & "Please provide a name for the device type" & """;") 'Descriptor:Please provide a name for the device type%>
    }

    if (sMsg != "") {
      sMsg = sMsg + "<P>";
      if(document.all){
         document.all("validation").innerHTML = sMsg;
         document.all("validation").style.display = "block";
      }
      return false;
    }
  }
//-->
</SCRIPT>

<!-- #include file="validationJS.asp" --><!-- #include file="../NSStyleSheet.asp" -->
<%End If%>
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
      <%If (lErr <> NO_ERR) And (lErr <> ERR_INVALID_NAME) Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
      <%Else%>
      <BR />
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
        <%
        Response.Write(asDescriptors(670)) 'The device type is a group of devices for the user to select how they will receive the content. When a new site definition is created four standard device types will be created and configured automatically.
        %>
      </FONT>
      <BR />
      <BR />
      <%If lErr = ERR_INVALID_NAME Then%>
          <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV class="validation" id="validation"><LI>Please provide a different name, "<%=Server.HTMLEncode(aDeviceTypeInfo(DEV_TYPE_NAME))%>" already exists.<P></DIV></FONT>
      <%Else%>
          <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>
      <%End If%>

      <% Call RenderExistingDeviceTypes(sDeviceTypesXML, aDeviceTypeInfo(DEV_TYPE_ACTION), aDeviceTypeInfo(DEV_TYPE_ID)) %>

      <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
        <TR>
          <TD COLSPAN="2">
            <BR />
          </TD>
        </TR>

        <TR>
          <TD ALIGN="left" NOWRAP WIDTH="1%">
            <FORM ACTION="channels.asp"><INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;</FORM>
          </TD>
          <TD ALIGN="left" NOWRAP WIDTH="98%">
            <FORM ACTION="devices_config.asp"><INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;</FORM>
          </TD>
        </TR>

      </TABLE>

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