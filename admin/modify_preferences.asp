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
Dim aSiteProperties()
Redim aSiteProperties(MAX_SITE_PROP)
Dim lStatus
Dim sLocaleData
Dim iPos
Dim sLoginMode

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "preferences.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_PREFERENCES & " " & asDescriptors(286) 'Descriptor:Preferences
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SITE_MANAGEMENT

    lStatus = checkSiteConfiguration()

    'Check for actions cancelled:
    If oRequest("back") <> "" Then
		Erase aSiteProperties
        Response.Redirect("is_config.asp")
    End If

    'Read rest of request variables:
    aSiteProperties(SITE_PROP_NEW_USERS) = oRequest("usrs")
    sLocaleData = CStr(oRequest("locale"))
    iPos = Instr(1, sLocaleData, ";")
	aSiteProperties(SITE_PROP_NEW_LOCALE) = Left(sLocaleData, iPos - 1)
	aSiteProperties(SITE_PROP_GUI_LANG) = Right(sLocaleData, Len(sLocaleData)-iPos)

    'aSiteProperties(SITE_PROP_NEW_LOCALE) = oRequest("locale")

    aSiteProperties(SITE_PROP_NEW_EXPIRE) = oRequest("exp")
    'aSiteProperties(SITE_PROP_GUI_LANG) = oRequest("lang")
    aSiteProperties(SITE_PROP_USE_DHTML) = oRequest("dhtml")
    aSiteProperties(SITE_PROP_TMP_DIR) = oRequest("dir")
    aSiteProperties(SITE_PROP_PROMPT_CACHE) = oRequest("cache")
    aSiteProperties(SITE_PROP_SUMMARY_PAGE) = oRequest("summary")
    aSiteProperties(SITE_PROP_EMAIL) = oRequest("email")
    aSiteProperties(SITE_PROP_PHONE) = oRequest("phone")
    aSiteProperties(SITE_PROP_STREAM_ATTACHMENTS) = "1"
    aSiteProperties(SITE_PROP_TIMEZONE) = oRequest("timezone")

    'Based on the checked items, determine the login mode for the site
	If (oRequest("LoginMode") = "000")  Then
		sLoginMode = "NC_NORMAL"
	End If
	If (oRequest("LoginMode") = "001")  Then
		sLoginMode = "NC_NORMAL"
	End If
	If (oRequest("LoginMode") = "010")  Then
		sLoginMode = "IS_NORMAL"
	End If
	If (oRequest("LoginMode") = "011")  Then
		sLoginMode = "NC_IS_NORMAL"
	End If
	If (oRequest("LoginMode") = "100")  Then
		sLoginMode = "NT_NORMAL"
	End If
	If (oRequest("LoginMode") = "101")  Then
		sLoginMode = "NC_NT_NORMAL"
	End If
	If (oRequest("LoginMode") = "110")  Then
		sLoginMode = "IS_NT_NORMAL"
	End If
	If (oRequest("LoginMode") = "111")  Then
		sLoginMode = "NC_IS_NT_NORMAL"
	End If

    aSiteProperties(SITE_LOGIN_MODE) = sLoginMode
    aSiteProperties(SITE_AUTHENTICATION_SERVER_NAME) = oRequest("is_server_name")
    aSiteProperties(SITE_AUTHENTICATION_SERVER_PORT) = oRequest("is_server_port")

    aSiteProperties(SITE_ELEMENT_PROMPT_BLOCK_COUNT) = oRequest("element_count")
    aSiteProperties(SITE_OBJECT_PROMPT_BLOCK_COUNT) = oRequest("object_count")
	If len(oRequest("match_case")) > 0 then
		aSiteProperties(SITE_PROMPT_MATCH_CASE) = oRequest("match_case")
	Else
		aSiteProperties(SITE_PROMPT_MATCH_CASE) = 0
	End if

    'Set values for expiration:
    If aSiteProperties(SITE_PROP_NEW_EXPIRE) = "1" Then
        aSiteProperties(SITE_PROP_EXPIRE_VALUE) = oRequest("expDate")
    ElseIf aSiteProperties(SITE_PROP_NEW_EXPIRE) = "2" Then
        aSiteProperties(SITE_PROP_EXPIRE_VALUE) = oRequest("expCount")
    Else
        aSiteProperties(SITE_PROP_NEW_EXPIRE) = "0"
        aSiteProperties(SITE_PROP_EXPIRE_VALUE) = ""
    End If

    'We need to store the TMP_DIR with a backslash at the end:
    If Len(aSiteProperties(SITE_PROP_TMP_DIR)) > 0 Then
        If Mid(aSiteProperties(SITE_PROP_TMP_DIR), Len(aSiteProperties(SITE_PROP_TMP_DIR))) <> "\" Then
            aSiteProperties(SITE_PROP_TMP_DIR) = aSiteProperties(SITE_PROP_TMP_DIR) & "\"
        End If
    End If


    'Create the new site, if succesfull, we must continue with the Next Page"
    lErr = setSiteProperties(aSiteProperties, FLAG_PROP_GROUP_OTHER)
    If lErr = NO_ERR Then
        Call ResetApplicationVariables()
		Erase aSiteProperties

        Call Response.Redirect("adminSummary.asp?section=3")
    Else
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "modify_preferences.asp", "", "", "Error calling getSiteProperties", LogLevelTrace)
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
      <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(286) , "preferences.asp") 'Descriptor: Return to: 'Descriptor:Preferences %>
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