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
Dim lStatus

    'Check for actions cancelled:
    If oRequest("back") <> "" Then
        Response.Redirect("services_select.asp?fid=" & oRequest("fid") & "&id=" & oRequest("id"))
    End If

    If oRequest("selectQO") <> "" Then
        Response.Redirect("services_select_qo.asp")
    End If
    
    
    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "modify_services_static.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_STATIC & " " & asDescriptors(286) 'Descriptor:Preferences
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3
    
    lStatus = checkSiteConfiguration()
    
    If lErr = NO_ERR Then
        Call Response.Redirect("services_select_qo.asp")
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
      <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(286) , "preferences.asp") 'Descriptor: Return to: 'Descriptor:Preferences %>
    </TD>
    
    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="1%">
        <!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>