<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="../CustomLib/SiteConfigCuLib.asp" -->
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<%
Dim lStatus
    Dim sSiteId

    'Check for actions cancelled:
    If oRequest("cancel") <> "" Then
        Response.Redirect("site_config.asp")
    End If

    'Back
    If oRequest("back") <> "" Then
        Call Response.Redirect("select_portal.asp")
    End If

    aPageInfo(S_NAME_PAGE) = "select_site.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_SITE & " " & asDescriptors(623) 'Descriptor:Site Definition
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SITE_MANAGEMENT

    If lErr = NO_ERR Then

        sSiteId = oRequest("sid")

        'If a new site was selected, get it's properties and call setSite:
        If sSiteId <> Application.Value("SITE_ID") Then

            If lErr = NO_ERR Then
                lErr = SetSite(sSiteId)
            End If

            If lErr = NO_ERR Then
                Call GenerateSiteDynamicSQL(sSiteId)
            End If

        End If
    End If

    lStatus = checkSiteConfiguration()

    'If everything was fine, redirect to next page, which is the one that needs to get configured.
    If lErr = NO_ERR Then

        Call ResetApplicationVariables()

        If (lStatus And CONFIG_MISSING_SITE) <> 0 Then
            Call Response.Redirect("site_name.asp")

        ElseIf (lStatus And CONFIG_MISSING_AUREP) <> 0 Then
            Call Response.Redirect("select_aurep.asp")

        ElseIf (lStatus And CONFIG_MISSING_SBREP) <> 0 Then
            Call Response.Redirect("select_sbrep.asp")
        Else
            Call Response.Redirect("adminSummary.asp?section=2")
        End If

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
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="1%">
        <!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
