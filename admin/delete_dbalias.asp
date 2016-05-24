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
<!-- #include file="../CustomLib/SiteConfigCuLib.asp" -->
<!-- #include file="../CustomLib/DeviceTypesCuLib.asp" -->
<!-- #include file="../CustomLib/ServicesConfigCuLib.asp" -->

<%
Dim lStatus
Dim aDBAliasInfo
Dim lRepositoryType
Dim bConfirm
Dim sRedirectPage
Dim sConfirmed
Dim aMapInfo

Dim sName

	'Get settings based upon repository type
    Call ParseRequestForDBAlias(oRequest, aDBAliasInfo, lRepositoryType)

    If Len(aDBAliasInfo(DBALIAS_DECODED_NAME)) > 0 Then
		sName = aDBAliasInfo(DBALIAS_DECODED_NAME)
	Else
		sName = aDBAliasInfo(DBALIAS_NAME)
    End If

    'Based upon DB type, set variables
    Select Case lRepositoryType
    Case REPOSITORY_MD:
        aPageInfo(S_TITLE_PAGE) = STEP_SELECT_MD_DBALIAS & " " & asDescriptors(830) 'Descriptor:Delete a Database Alias
        aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_ENGINE_CONFIG
        sRedirectPage = "select_md.asp"
    Case REPOSITORY_AUREP:
        aPageInfo(S_TITLE_PAGE) = STEP_SITE_AUREP_DBALIAS & " " & asDescriptors(830) 'Descriptor:Delete a Database Alias
        aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT
        sRedirectPage = "select_aurep.asp"
    Case REPOSITORY_SBREP:
        aPageInfo(S_TITLE_PAGE) = STEP_SITE_SBREP_DBALIAS & " " & asDescriptors(830) 'Descriptor:Delete a Database Alias
        aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT
        sRedirectPage = "select_sbrep.asp"
    Case REPOSITORY_WAREHOUSE:
        aPageInfo(S_TITLE_PAGE) = STEP_SITE_SBREP_DBALIAS & " " & asDescriptors(830) 'Descriptor:Delete a Database Alias
        aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES

        If lErr = NO_ERR Then
            lErr = ParseRequestForMapInfo(oRequest, aSvcConfigInfo, aMapInfo)
        End If

        sRedirectPage = "services_map_tables.asp?mid=" & Server.URLEncode(aMapInfo(MAP_ID)) & "&mn=" & Server.URLEncode(aMapInfo(MAP_NAME)) & "&mf=" & Server.URLEncode(aMapInfo(MAP_FILTER)) & "&" & CreateRequestForSvcConfig(aSvcConfigInfo)

    End Select

	aPageInfo(N_ALIAS_PAGE) = lRepositoryType

    'Check for actions back:
    If (Len(oRequest("cancel")) > 0) Then
        Response.Redirect(sRedirectPage)
    End If

    lStatus = checkSiteConfiguration()

    sConfirmed = oRequest("confirm")

    'If no given name so far for the dbalias:
    If lErr = NO_ERR Then

        'If confirmed, delete the dbalias:
        If sConfirmed = "yes" Then

            lErr = DeleteDBAlias(aDBAliasInfo)

            'If everything went fine, redirect to the original page again:
            If lErr = NO_ERR Then
                Call Response.Redirect(sRedirectPage)
            End If

        End If

    End If

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<!-- #include file="../NSStyleSheet.asp" -->

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
		<%If lRepositoryType = REPOSITORY_MD Then%>
			<!-- #include file="_toolbar_engine_config.asp" -->
        <%ElseIf lRepositoryType = REPOSITORY_AUREP Then%>
			<!-- #include file="_toolbar_portal_management.asp" -->
		<%ElseIf lRepositoryType = REPOSITORY_SBREP Then%>
			<!-- #include file="_toolbar_portal_management.asp" -->
				<%ElseIf lRepositoryType = REPOSITORY_WAREHOUSE Then%>
			<!-- #include file="_toolbar_services.asp" -->
		<%End If%>
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(6) & " " & asDescriptors(18) , "select_site.asp") 'Descriptor: Return to:'Descriptor:Site Definition %>
      <%Else%>
      <BR />

      <TABLE BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
        <FORM ACTION="delete_dbalias.asp" METHOD="POST">
         <%If lRepositoryType = 4 Then RenderSvcConfigInputs(aSvcConfigInfo) %>
          <TR>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
                <B><%Call Response.write(asDescriptors(322)) 'Descriptor:Warning!%></B><BR />
                <%Call Response.write(asDescriptors(839)) 'Descriptor:If you delete a database connection, objects using it (i.e. Object Repository, Storage Mappings, etc.) will become invalid.%><BR />
                <BR />
                <%Call Response.Write(Replace(asDescriptors(840), "#", "<B>" & Server.HTMLEncode(sName) & "</B>" )) 'Descriptor:Are you sure you want to delete '#' ?%>
                <BR/>
                <BR />

              </FONT>
            </TD>
          </TR>
          <TR>
            <TD ALIGN=CENTER>
              <BR />
              <INPUT name=tAliasName type=HIDDEN value="<%=Server.HTMLEncode(aDBAliasInfo(DBALIAS_NAME))%>"></INPUT>
              <INPUT name=rep        type=HIDDEN value="<%=lRepositoryType%>"></INPUT>
              <INPUT name=confirm    type=HIDDEN value="yes"   ></INPUT>
              <INPUT name=ok         type=submit class="buttonClass" value="<%Response.Write(asDescriptors(543)) 'Descriptor:Ok%>"></INPUT> &nbsp;
              <INPUT name=cancel     type=submit class="buttonClass" value="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT>
            </TD>
          </TR>
        </FORM>
      </TABLE>

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