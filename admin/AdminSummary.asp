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
<%
	Dim lStatus
	Dim lAdminSection
	Dim sRedirectPage
	Dim	sPreviousPage
	Dim bContinue

    lStatus = checkSiteConfiguration()

    'Getting section from URL
    lAdminSection = Clng(oRequest("section"))

    'Set the PageInfo to be used by the navigator bar and the header.
    lErr = getSummarySettings(aPageInfo, lAdminSection, sRedirectPage, sPreviousPage, sErrorMessage)

    If oRequest("back") <> "" Then
        Call Response.Redirect(sPreviousPage)
    End If


    If oRequest("next") <> "" Then
        Call Response.Redirect(sRedirectPage)
    End If

    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = "section=" & lAdminSection

    bContinue = True

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
        <%If lAdminSection = SECTION_ENGINE_CONFIG Then%>
			 <!-- #include file="_toolbar_engine_config.asp" -->
		<%ElseIf lAdminSection = SECTION_PORTAL_MANAGEMENT Then%>
			 <!-- #include file="_toolbar_portal_management.asp" -->
		<%ElseIf lAdminSection = SECTION_SITE_MANAGEMENT Then%>
			 <!-- #include file="_toolbar_site_preferences.asp" -->
		<%ElseIf lAdminSection = SECTION_SERVICES Then%>
			 <!-- #include file="_toolbar_site_preferences.asp" -->
		<%End If%>
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> NO_ERR Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
      <%Else%>
        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD COLSPAN="2">
              <BR />
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>">
				<%If lAdminSection = SECTION_ENGINE_CONFIG Then%>
					<!-- #include file="summaryEngine_widget.asp" -->
				<%ElseIf lAdminSection = SECTION_PORTAL_MANAGEMENT Then%>
					<!-- #include file="summaryPortal_widget.asp" -->
				<%ElseIf lAdminSection = SECTION_SITE_MANAGEMENT Then%>
					<!-- #include file="summarySite_widget.asp" -->
				<%ElseIf lAdminSection = SECTION_SERVICES Then%>
					<!-- #include file="summaryEngine_widget.asp" -->
				<%End If%>
		      </FONT>
		      <HR noShade SIZE=1>
            </TD>
          </TR>

          <TR>
            <FORM ACTION="AdminSummary.asp" id=form1 name=form1>
              <TD ALIGN="left" NOWRAP WIDTH="1%">
                <BR/><INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
                <INPUT name=section type=hidden value="<%=lAdminSection%>"></INPUT>
              </TD>
              <TD ALIGN="left" NOWRAP WIDTH="98%">
                <BR/><%If bContinue Then%><INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;<%End If%>
              </TD>
            </FORM>
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