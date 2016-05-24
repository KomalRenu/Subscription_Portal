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
Dim lStatus

Dim lCount
Dim i
Dim sId
Dim sRequestForSvcConfig
Dim aPromptList

    If lErr = NO_ERR Then
        lErr = ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)
    End If

    'Check for actions cancelled:
    If Len(oRequest("back")) > 0 Then
		sRequestForSvcConfig = CreateRequestForSvcConfig(aSvcConfigInfo)
        Response.Redirect("services_subsset.asp?" & sRequestForSvcConfig)
    End If

    'if this is the mapping of the question objects, the the Details for questions:
    If lErr = NO_ERR Then
        If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then
            lErr = GetPromptsForQuestion(aSvcConfigInfo, aPromptList)
        End If    
    End If
        
        
    If lErr = NO_ERR Then
        lErr = SaveQuestionConfig(aSvcConfigInfo, aPromptList)
    End If
    
    If lErr = NO_ERR Then
		sRequestForSvcConfig = CreateRequestForSvcConfig(aSvcConfigInfo)
        Response.Redirect("services_select_map.asp?" & sRequestForSvcConfig)
    End If
    
    'We will show an error:
    aPageInfo(S_NAME_PAGE) = "services_select_qo.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_PREFERENCES & " " & asDescriptors(286) 'Descriptor:Preferences
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForSvcConfig(aSvcConfigInfo)
    
    lStatus = checkSiteConfiguration()
    
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
            <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & "Overview" , "services_overview.asp") 'Descriptor: Return to: 'Descriptor:Overview %>
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