<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/DeviceTypesCuLib.asp" -->
<%
    Dim sDeviceTypesXML
    Dim aDeviceTypeInfo()

    Dim lStatus

    If oRequest("back") <> "" Then
        Response.Redirect "deviceTypes.asp"
    End If

    lErr = ParseRequestForDeviceType(oRequest, aDeviceTypeInfo)

    If lErr = NO_ERR Then
        lErr = ReadDeviceTypesXML(sDeviceTypesXML)
    End If

    If lErr = NO_ERR Then
        If Request.Form.Count > 0 Then
            If aDeviceTypeInfo(DEV_TYPE_ACTION) = "new" Then
                If aDeviceTypeInfo(DEV_TYPE_NAME) <> "" Then
                    lErr = AddNewDeviceType(aDeviceTypeInfo, sDeviceTypesXML)
                    If lErr = NO_ERR Then
                        lErr = WriteDeviceTypesXML(sDeviceTypesXML)
                        If lErr = NO_ERR Then
                            Response.Redirect "deviceTypeFolders.asp?dtID=" & aDeviceTypeInfo(DEV_TYPE_ID)
                        End If
                    End If
                End If
            ElseIf aDeviceTypeInfo(DEV_TYPE_ACTION) = "edit" Then
                lErr = EditDeviceType(aDeviceTypeInfo, sDeviceTypesXML)
                If lErr = NO_ERR Then
                    lErr = WriteDeviceTypesXML(sDeviceTypesXML)
                    If lErr = NO_ERR Then
                        Response.Redirect "deviceTypeFolders.asp?dtID=" & aDeviceTypeInfo(DEV_TYPE_ID)
                    End If
                End If
            End If
        End If
    End If

    If lErr = NO_ERR Then
        If Len(aDeviceTypeInfo(DEV_TYPE_ID)) > 0 Then
            lErr = GetDeviceTypeInfo(aDeviceTypeInfo, sDeviceTypesXML)
        End If
    End If

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "editDeviceType.asp"
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SITE_MANAGEMENT
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForDeviceType(aDeviceTypeInfo)
    lStatus = checkSiteConfiguration()

    If aDeviceTypeInfo(DEV_TYPE_ACTION) = "new" Then
        aPageInfo(S_TITLE_PAGE) = STEP_DEVICE_TYPES_EDIT & " " & asDescriptors(515) 'Descriptor: New Definition
    Else
        aPageInfo(S_TITLE_PAGE) = STEP_DEVICE_TYPES_EDIT & " " & asDescriptors(516) & " " & Server.HTMLEncode(aDeviceTypeInfo(DEV_TYPE_NAME)) 'Descriptor: Definition:
    End IF

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<% If aDeviceTypeInfo(DEV_TYPE_ACTION) = "new" Then %>
<SCRIPT LANGUAGE=javascript>
<!--

  function validateForm() {
  var sMsg

    sMsg = "";
    if ((FormDevType.DTName.value == "") || isBlank(FormDevType.DTName.value)) {
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
<!-- #include file="validationJS.asp" -->
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
      <%If (lErr <> 0) And (lErr <> ERR_INVALID_NAME) Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(493), "deviceTypes.asp") 'Descriptor:Device Types%>
      <%Else%>

        <FORM METHOD=POST ACTION="editDeviceType.asp" NAME="FormDevType">
          <INPUT TYPE="HIDDEN" NAME="action" VALUE="<%=aDeviceTypeInfo(DEV_TYPE_ACTION)%>" />
          <INPUT TYPE="HIDDEN" NAME="dtID" VALUE="<%=aDeviceTypeInfo(DEV_TYPE_ID)%>" />

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD>
              <BR />
            </TD>
          </TR>
          <TR>
            <TD>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
				<%Response.Write asDescriptors(671) 'Descriptor:For each new device type definition, you will need to specify the device type name, the large and small icon URL, the address format and which fields will be used when displaying and editing.%>
		      <P></FONT>
		    </TD>
	      </TR>
	    </TABLE>

        <%If lErr = ERR_INVALID_NAME Then%>
            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV class="validation" id="validation"><LI>Please provide a different name, "<%=Server.HTMLEncode(aDeviceTypeInfo(DEV_TYPE_NAME))%>" already exists.<P></DIV></FONT>
        <%Else%>
            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>
        <%End If%>


        <% Call RenderDeviceTypeEditor(aDeviceTypeInfo) %>

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD>
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="left" NOWRAP>
              <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
              <INPUT name=next <% If aDeviceTypeInfo(DEV_TYPE_ACTION) = "new" Then Response.Write(" onClick=""return validateForm();"" ") %> type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
            </TD>
          </TR>

        </TABLE>

        </FORM>

      <%End If%>

    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="1%">
        <!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>