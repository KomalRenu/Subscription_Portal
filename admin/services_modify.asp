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
Dim sOldAnswer
Dim sAnswer
Dim aSiteProperties()
Redim aSiteProperties(MAX_SITE_PROP)

    'Check for actions cancelled:
    If oRequest("back") <> "" Then
		Erase aSiteProperties
        Response.Redirect("adminOverview.asp?section=" & SECTION_SERVICES)
    End If
    
    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "services_config.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_PREFERENCES & " " & asDescriptors(286) 'Descriptor:Preferences
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    
    lStatus = checkSiteConfiguration()
    
    sAnswer = oRequest("ans")
    sOldAnswer = oRequest("old")
    
    'Save values if necessary:
    If StrComp(sAnswer, sOldAnswer) <> 0 Then
        aSiteProperties(SITE_PROP_DEFAULT_ANSWER) = sAnswer
        lErr = setSiteProperties(aSiteProperties, FLAG_PROP_GROUP_SERVICES)
    End If

    If lErr = NO_ERR Then
    
        Call ResetApplicationVariables()
        Erase aSiteProperties
        
        'Check for actions cancelled:
        If oRequest("conf") <> "" Then
            Response.Redirect("services_overview.asp")
        End If
    
        Call Response.Redirect("finish.asp")
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
    <TD VALIGN="TOP">
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
            <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(362) , "services_config.asp") 'Descriptor: Return to: 'Descriptor:Services %>
          </TD>
          
          <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

          <TD WIDTH="1%" VALIGN="TOP">
              <!-- #include file="help_widget.asp" -->
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
	Erase aSiteProperties
%>