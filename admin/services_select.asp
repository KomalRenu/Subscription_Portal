<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/ServicesConfigCuLib.asp" -->
<%
Dim sFolderId

Dim sType
Dim sFolderContentXML
Dim sFolderLink
Dim sId

Dim lStatus


    'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_SELECT & " " & asDescriptors(781)'"Select Service"
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES

    lStatus = checkSiteConfiguration()

    'Read rest of request variables:
    sFolderId = oRequest("fid")
    sId       = oRequest("id")
    sType     = TYPE_SERVICE

    If Len(sFolderId) = 0 Then sFolderId = oRequest("sfid")

    Redim aSvcConfigInfo(MAX_SVCCFG_INFO)
    aSvcConfigInfo(SVCCFG_SVC_ID) = sId
    aSvcConfigInfo(SVCCFG_SVC_PARENT_ID) = sFolderId

    'Get the previous values:
    If lErr = NO_ERR Then
        lErr = co_getFolderXML(sFolderId, ROOT_APP_FOLDER_TYPE, sFolderContentXML)
        If lErr <> NO_ERR Then aSvcConfigInfo(SVCCFG_SVC_PARENT_ID) = ""
    End If

    sFolderLink = "services_select.asp?id="
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForSvcConfig(aSvcConfigInfo)

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
        <!-- #include file="_toolbar_services.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & "Services Overview", "services_overview.asp") 'Descriptor:Services Overview%>
      <%Else%>
        <BR />
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
          <%Call Response.Write(asDescriptors(788)) 'Please select the service you wish to configure.%><BR />
          <BR />
        </FONT>

        <FORM ACTION="services_static.asp">
        <INPUT TYPE="HIDDEN" NAME="fid" VALUE="<%=sFolderId%>" />
        <!-- begin folder content -->
          <!-- #include file="folder_content_widget.asp" -->
        <!-- end folder content -->

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD COLSPAN="2">
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="left" NOWRAP WIDTH="1%">
              <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
            </TD>
            <TD ALIGN="left" NOWRAP WIDTH="98%">
              <%If oItems.length > 0 Then%>
              <INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
              <%End If%>
            </TD>
          </TR>
        </TABLE>
        </FORM>
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