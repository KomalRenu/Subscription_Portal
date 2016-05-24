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

Dim sId
Dim sName
Dim sType

Dim sConfirmed

Dim aSiteProperties
Dim sDefault

    'Check for actions cancelled:
    If oRequest("cancel") <> "" Then
        Call Response.Redirect("select_site.asp")
    End If

    aPageInfo(S_NAME_PAGE) = "select_site.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_SITE & " " & asDescriptors(252) 'Descriptor:Delete Confirmation
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT

    lStatus = checkSiteConfiguration()


    'Read request variables:
    sId = oRequest("id")
    sType = oRequest("tp")
    sName = oRequest("n")

    sConfirmed = oRequest("confirm")

    'If no given name so far for the site:
    If lErr = NO_ERR Then

        'If confirmed, delete the site:
        If sConfirmed = "yes" Then

            lErr = DeleteSite(sId)

            'If everything went fine, redirect to the select_site page again:
            If lErr = NO_ERR Then
                Call ResetApplicationVariables()
                Call Response.Redirect("select_site.asp")
            End If

        End If

    End If

	Redim aSiteProperties(MAX_SITE_PROP)
	aSiteProperties(SITE_PROP_ID) = sId
	Call getSiteProperties(aSiteProperties)
	sDefault = aSiteProperties(SITE_IS_DEFAULT)

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
        <!-- #include file="_toolbar_portal_management.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623) , "select_site.asp") 'Descriptor: Return to:'Descriptor:Site Definition %>
      <%Else%>
      <BR />
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>

      <TABLE BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
        <FORM ACTION="delete_site.asp" METHOD="POST">
          <TR>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
                <%If sDefault="1" then
                	Response.write(asDescriptors(923)) 'Descriptor:You are deleting the default portal site. Deleting this site will disable deliveries.....
                 Else
                	Response.write(asDescriptors(702)) 'Descriptor:Are you sure you want to delete this site%> <B><%=Server.HTMLEncode(sName)%></B><%
              	 End If
              	%>
              </FONT>
            </TD>
          </TR>
          <TR>
            <TD ALIGN=CENTER>
              <BR />
              <INPUT name=id      type=HIDDEN value="<%=sId%>"></INPUT>
              <INPUT name=tp      type=HIDDEN value="<%=sType%>"></INPUT>
              <INPUT name=confirm type=HIDDEN value="yes"   ></INPUT>
              <INPUT name=ok      type=submit class="buttonClass" value="<%Response.Write(asDescriptors(543)) 'Descriptor:Ok%>"></INPUT> &nbsp;
              <INPUT name=cancel  type=submit class="buttonClass" value="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT>
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