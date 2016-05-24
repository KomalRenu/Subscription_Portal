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

Dim aSiteProperties()
Redim aSiteProperties(MAX_SITE_PROP)

Dim sDBA
Dim sPrefix
Dim nChkSBR

    'Cancel
    If oRequest("cancel") <> "" Then
		Erase aSiteProperties
        Response.Redirect("select_site.asp")
    End If

    'Back
    If oRequest("back") <> "" Then
		Erase aSiteProperties
        Call Response.Redirect("site_name.asp")
    End If

    aPageInfo(S_NAME_PAGE) = "select_aurep.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SITE_AUREP & " " & asDescriptors(575) 'Descriptor:Project Repository
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT

    lStatus = checkSiteConfiguration()

    sDBA    = oRequest("dba")
    sPrefix = oRequest("pre")
    nChkSBR = oRequest("chksbr")

    If lErr = NO_ERR Then

        lErr = getSiteProperties(aSiteProperties)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "modify_aurep.asp", "", "", "Error calling getSiteProperties", LogLevelTrace)

    End If

    If lErr = NO_ERR Then
        aSiteProperties(SITE_PROP_AUREP) = sDBA
        aSiteProperties(SITE_PROP_AUREP_PREFIX) = sPrefix

        If nChkSBR = "yes" Then
            aSiteProperties(SITE_PROP_SBREP) = sDBA
            aSiteProperties(SITE_PROP_SBREP_PREFIX) = sPrefix
        End If

        lErr = setSiteProperties(aSiteProperties, FLAG_PROP_GROUP_CONN)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "modify_aurep.asp", "", "", "Error calling setSiteProperties" , LogLevelTrace)

    End If


    If lErr = NO_ERR Then
        lErr = SetSite(aSiteProperties)
    End If

    If lErr = NO_ERR Then
        lErr = ResetSubscriptionEngine()
    End If

    If lErr = NO_ERR Then
        Call ResetApplicationVariables()
        Erase aSiteProperties

        If nChkSBR = "yes" Then
			Call Response.Redirect("adminSummary.asp?section=2")
        Else
			Call Response.Redirect("select_sbrep.asp")
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
        <!-- #include file="_toolbar_portal_management.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(575), "select_aurep.asp") 'Descriptor: Return to: 'Project Repository %>
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
	Erase aSiteProperties
%>