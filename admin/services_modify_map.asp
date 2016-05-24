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

Dim aMapInfo
Dim aTablesInfo


    If lErr = NO_ERR Then
        lErr = ParseRequestForMapInfo(oRequest, aSvcConfigInfo, aMapInfo)
    End If

    'Check for actions cancelled:
    If Len(oRequest("back")) > 0 Then
		sRequestForSvcConfig = CreateRequestForSvcConfig(aSvcConfigInfo)
        Response.Redirect("services_select_qo.asp?" & sRequestForSvcConfig)
    End If

    'Validate Prompts:        
    If lErr = NO_ERR Then
        lErr = SaveQuestionConfig(aSvcConfigInfo, aMapInfo(MAP_QO_IS), aMapInfo(MAP_QO_PROMPT_COUNT))
    End If
    
        
    If lErr = NO_ERR Then
        lErr = SaveMapConfig(aSvcConfigInfo, aMapInfo)
    End If
    
    If lErr = NO_ERR Then
		sRequestForSvcConfig = CreateRequestForSvcConfig(aSvcConfigInfo)
        Response.Redirect("services_subsset.asp?" & sRequestForSvcConfig)
    End If
    
    'We will show an error:
    aPageInfo(S_NAME_PAGE) = "services_select_map.asp"
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