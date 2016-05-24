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
<!-- #include file="../CustomLib/ServicesConfigCuLib.asp" -->

<!-- #include file="editDBAlias_widget.asp" -->
<%
Dim lStatus
Dim aDBAliasInfo
Dim lRepositoryType
Dim bConfirm
Dim sRedirectPage
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
        aPageInfo(S_TITLE_PAGE) = STEP_SELECT_MD_DBALIAS & " " & asDescriptors(891) & " : " & sName
        aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_ENGINE_CONFIG
        sRedirectPage = "select_md.asp"
    Case REPOSITORY_AUREP:
        aPageInfo(S_TITLE_PAGE) = STEP_SITE_AUREP_DBALIAS & " " & asDescriptors(891) & " : " & sName
        aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT
        sRedirectPage = "select_aurep.asp"
    Case REPOSITORY_SBREP:
        aPageInfo(S_TITLE_PAGE) = STEP_SITE_SBREP_DBALIAS & " " & asDescriptors(891) & " : " & sName
        aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT
        sRedirectPage = "select_sbrep.asp"
    Case REPOSITORY_WAREHOUSE:
        aPageInfo(S_TITLE_PAGE) = STEP_SITE_SBREP_DBALIAS & " " & asDescriptors(891) & " : " & sName
        aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES

        If lErr = NO_ERR Then
            lErr = ParseRequestForMapInfo(oRequest, aSvcConfigInfo, aMapInfo)
        End If

        sRedirectPage = "services_map_tables.asp?mid=" & Server.URLEncode(aMapInfo(MAP_ID)) & "&mn=" & Server.URLEncode(aMapInfo(MAP_NAME)) & "&mf=" & Server.URLEncode(aMapInfo(MAP_FILTER)) & "&" & CreateRequestForSvcConfig(aSvcConfigInfo)

    End Select

    aPageInfo(N_ALIAS_PAGE) = lRepositoryType

    'Check for actions back:
    If (Len(oRequest("back")) > 0) Then
        Response.Redirect(sRedirectPage)
    End If

    lStatus = checkSiteConfiguration()

    'Proccess form submitted
    If oRequest("submit") <> 0 Then
        lErr = ProccessEditDBAlias(aDBAliasInfo, lRepositoryType)

        If lErr = NO_ERR Then
            Response.Redirect(sRedirectPage)
        Else
            Select Case lErr
            Case ERR_WRONG_DBALIAS_DEFINITION
                sErrorMessage = Replace(asDescriptors(826), "#", asDescriptors(828)) 'Descriptos:There was an error while attempting to validate the database connection (#). Please verify your information and try again. /wrong database connection definition
            Case ERR_UNABLE_CONNECT_TRANSACTOR
                sErrorMessage = Replace(asDescriptors(826), "#", asDescriptors(827)) 'Descriptos:There was an error while attempting to validate the database connection (#). Please verify your information and try again. /could not establish connection to the subscription engine
            Case ERR_WRONG_DBALIAS_NAME
                sErrorMessage = Replace(asDescriptors(826), "#", asDescriptors(829)) 'Descriptos:There was an error while attempting to validate the database connection (#). Please verify your information and try again. /wrong datasource name
            Case Else
                sErrorMessage = Replace(asDescriptors(826), "#", lErr) 'Descriptos:There was an error while attempting to validate the database connection (#). Please verify your information and try again.
            End Select
        End If

   End If

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>
  <SCRIPT LANGUAGE=javascript>

  function validateForm() {
  var sMsg

    sMsg = "";
    if (FormEditDBAlias.tAliasName.value == "" || isBlank(FormEditDBAlias.tAliasName.value)) {
      <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(704) & """;") 'Descriptor:Please provide a name to the Alias Definition %>
    } else {
      if (checkIsAlphaNumeric(FormEditDBAlias.tAliasName.value) == false) <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(882) & """;") 'Descriptor:Database connection names cannot contain special characters. Please enter a name which only uses a-z, 0-9 or underscore. %>
    }
    if (FormEditDBAlias.tServerName.value == "" || isBlank(FormEditDBAlias.tServerName.value)) <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(703) & """;") 'Descriptor:Please provide a name to the Server Definition %>
    if (FormEditDBAlias.tDBName.value == "" || isBlank(FormEditDBAlias.tDBName.value)) <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(705) & """;")'Descriptor:Please provide a name to the Database Definition %>
    if (FormEditDBAlias.tConfirmPassword.value != FormEditDBAlias.tPassword.value) <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(276) & """;") 'Descriptor:Your second new password does not match the first. Please try again.%>

    if (sMsg != "") {
      if(document.all){
         document.all("validation").innerHTML = sMsg;
         document.all("validation").style.display = "block";
      }
      return false;
    }
  }

</SCRIPT>

<!-- #include file="validationJS.asp" -->
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
      <%If lErr <> NO_ERR And aDBAliasInfo(DBALIAS_CONFIRM) = False Then
	  	Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(621), "editBAlias.asp") 'Descriptor: Return to:'Descriptor:Subscription Engine Location"
	  Else%>
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="<%If lErr = NO_ERR Then Response.write "display:none;"%>" class="validation" id="validation"><LI><%=sErrorMessage%></DIV></FONT>
      <BR />

      <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
        <FORM NAME="FormEditDBAlias" ACTION="editDBAlias.asp" METHOD="POST">
        <%If lRepositoryType = 4 Then
             RenderSvcConfigInputs(aSvcConfigInfo) %>
        <INPUT TYPE="HIDDEN" NAME="mid" VALUE="<%=aMapInfo(MAP_ID)%>" />
        <INPUT TYPE="HIDDEN" NAME="dba" VALUE="<%=aMapInfo(MAP_DBALIAS)%>" />
        <INPUT TYPE="HIDDEN" NAME="mf" VALUE="<%=aMapInfo(MAP_FILTER)%>" />
        <%End If%>
          <TR>
            <TD>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" COLOR="#000000">
				<%Call Response.Write(asDescriptors(818)) 'Enter the Database Alias information:               %>
              </FONT>
              <BR />
              <BR />
            </TD>
          </TR>

          <TR>
            <TD>
              <!--Start DBAlias list: -->
              <%
				lErr = displayEditDBAliasWidget(aDBAliasInfo, asDescriptors)
              %>
              <!--End DBAlias list -->
            </TD>
          </TR>
          <TR>
            <TD>
              <BR />
            </TD>
          </TR>
          <TR>
            <TD ALIGN="left" NOWRAP>
              <INPUT name=back   type=submit class="buttonClass" value="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT> &nbsp;
  			  <INPUT name=next   onClick="return validateForm();" type=submit class="buttonClass" value="<%Response.Write(asDescriptors(890)) 'Update database connection%>"></INPUT> &nbsp; <%'Need Descriptor:Add DB Alias%>
  			  <INPUT name=rep  type=hidden value=<%=lRepositoryType%> />
            </TD>
          </TR>
          <%End If%>
        </FORM>
      </TABLE>
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="1%">
        <!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
    Erase aDBAliasInfo
%>
