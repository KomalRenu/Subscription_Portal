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

Dim aTablesInfo
Dim aMapInfo

Dim lCount
Dim i
Dim sId


    If lErr = NO_ERR Then
        lErr = ParseRequestForMap(oRequest, aSvcConfigInfo, aMapInfo, aTablesInfo)
    End If

    'Check for actions cancelled:
    If Len(oRequest("back")) > 0 Then
		Erase aTablesInfo
        Response.Redirect("services_map_tables.asp?mid=" & Server.URLEncode(aMapInfo(MAP_ID)) & "&dba=" & Server.URLEncode(aMapInfo(MAP_DBALIAS)) & "&mf=" & Server.URLEncode(aMapInfo(MAP_FILTER)) & "&tbls=" & Server.URLEncode(aMapInfo(MAP_TABLES)) & "&" & CreateRequestForSvcConfig(aSvcConfigInfo))
    End If
        
    If lErr = NO_ERR Then
        lErr = SaveStorageMapping(aSvcConfigInfo, aMapInfo, aTablesInfo)
    End If
    
    If lErr = NO_ERR Then
        If Len(aSvcConfigInfo(SVCCFG_QO_ID)) > 0 Then
            Response.Redirect("services_select_map.asp?mid=" & aMapInfo(MAP_ID) & "&" & CreateRequestForSvcConfig(aSvcConfigInfo))
        Else
            lErr = SaveMapConfig(aSvcConfigInfo, aMapInfo)
            
            If lErr = NO_ERR Then
                Response.Redirect("services_subsset_save.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo))
            End If
        End If
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
<%
	Erase aTablesInfo
%>