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
<%
Dim lStatus

Dim aSiteProperties()
Redim aSiteProperties(MAX_SITE_PROP)
Dim sSiteId
Dim sName
Dim sDesc

    'Cancel
    If oRequest("cancel") <> "" Then
		Erase aSiteProperties
        Response.Redirect("site_config.asp")
    End If

    'Back
    If oRequest("back") <> "" Then
		Erase aSiteProperties

        If oRequest("sid") = "" Then
            Call Response.Redirect("select_md.asp")
        Else
            Call Response.Redirect("select_site.asp")
        End If
    End If

    aPageInfo(S_NAME_PAGE) = "site_name.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SITE_NAME & " " & asDescriptors(482) 'Descriptor:Name and Description
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT

    lStatus = checkSiteConfiguration()

    sSiteId = oRequest("sid")
    sName   = oRequest("n")
    sDesc   = oRequest("des")

    'Check if we either need to create or edit a site definition:
    If (sSiteId = "new") Or (sSiteId = "") Then
        'For new site, get default values:
        lErr = getDefaultSiteProperties(aSiteProperties)

        If lErr = NO_ERR Then
            aSiteProperties(SITE_PROP_NAME) = sName
            aSiteProperties(SITE_PROP_DESC) = sDesc

            lErr = CreateSite(aSiteProperties)

            If lErr = NO_ERR Then
                lErr = CreateDefaultChannels(aSiteProperties(SITE_PROP_ID))
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "modify_name.asp", "", "", "Error calling createDefaultChannels" , LogLevelTrace)
            End If

            If lErr = NO_ERR Then
                lErr = createDefaultDeviceTypes(aSiteProperties(SITE_PROP_ID))
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "modify_name.asp", "", "", "Error calling createDefaultChannels" , LogLevelTrace)
            End If

            'Reset the engine, after the site was created so Info gets set.
            If lErr = NO_ERR Then
                lErr = ResetSubscriptionEngine()
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "modify_name.asp", "", "", "Error calling resetSubscriptionEngine", LogLevelTrace)
            End If

        End If

    Else

        'For an existing site, call getSiteProperties:
        aSiteProperties(SITE_PROP_ID) = sSiteId
        lErr = getSiteProperties(aSiteProperties)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "modify_name.asp", "", "", "Error calling getSiteProperties", LogLevelTrace)

        If lErr = NO_ERR Then
            If (aSiteProperties(SITE_PROP_NAME) = sName) Or _
               (aSiteProperties(SITE_PROP_DESC) = sDesc) Then

                aSiteProperties(SITE_PROP_NAME) = sName
                aSiteProperties(SITE_PROP_DESC) = sDesc

                lErr = setSiteProperties(aSiteProperties, FLAG_PROP_GROUP_NAME)
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "modify_name.asp", "", "", "Error calling createSite" , LogLevelTrace)
            End If
        End If

    End If

    'Set the site to the new
    If lErr = NO_ERR Then
        lErr = SetSite(aSiteProperties(SITE_PROP_ID))
    End If


    If lErr = NO_ERR Then
        Call ResetApplicationVariables()
		Call Response.Redirect("select_aurep.asp")
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
<%
	Erase aSiteProperties

%>