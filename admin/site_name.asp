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

Dim sSiteId
Dim sName
Dim sOldName
Dim sDesc
Dim bConfirm

    'Back
    If oRequest("back") <> "" Then
		Erase aSiteProperties

        Call Response.Redirect("select_site.asp?skip=yes")
    End If


    aPageInfo(S_TITLE_PAGE) = STEP_SITE_NAME & " " & asDescriptors(482) 'Descriptor:Name and Description
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT

    sSiteId = oRequest("sid")
    sName   = oRequest("n")
    sOldName = oRequest("o")
    sDesc   = oRequest("des")
    bConfirm = Len(oRequest("confirm")) > 0

    'If the user has already input, we need to check that the new name is valid:
    If bConfirm Then

        'Check if the new name is valid:
        If sOldName <> sName Then
            If GetNewSiteName(sName) <> sName Then
                lErr = ERR_INVALID_NAME
            End If
        End If

        'If it is, continue to modify the site:
        If lErr = NO_ERR Then
            Call Response.Redirect("modify_name.asp?sid="  & sSiteId & "&n=" & Server.URLEncode(sName) & "&des=" & Server.URLEncode(sDesc))
        End If

    Else

        'If no ID sent in the request, assume we're editing current site:
        If Len(sSiteId) = 0 Then sSiteId = Application.Value("SITE_ID")

        'For new sites, use some default values:
        If (sSiteId = "new") Or (sSiteId = "") Then
            sName = GetNewSiteName(asDescriptors(592)) 'Descriptor:Default Hydra Site
            sDesc = asDescriptors(593) 'Descriptor:Site with default values
            sOldName = ""

            lStatus = lStatus Or CONFIG_MISSING_SITE Or CONFIG_MISSING_AUREP Or CONFIG_MISSING_SBREP

        Else
            'For existing, get the value from the MD:
            aSiteProperties(SITE_PROP_ID) = sSiteId

            lErr = getSiteProperties(aSiteProperties)

            If lErr = NO_ERR Then
                sName = aSiteProperties(SITE_PROP_NAME)
                sDesc = aSiteProperties(SITE_PROP_DESC)
                sOldName = sName
            End If

            'If editing a different site, call set site
            If lErr = NO_ERR Then
                If sSiteId <> Application.Value("SITE_ID") Then
                    lErr = SetSite(aSiteProperties(SITE_PROP_ID))
                End If
            End If

            lStatus = checkSiteConfiguration()

        End If

    End If

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<SCRIPT LANGUAGE=javascript>
<!--

  function validateForm() {
  var sMsg

    sMsg = "";
    if ((FormName.n.value == "") || isBlank(FormName.n.value)) {
      <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(596) & """") 'Descriptor:Please provide a name to the Site Definition %>
    }

    if (sMsg != "") {
      sMsg = sMsg + "<P>";
      if(document.all){
         document.all("validation").innerHTML = sMsg;
         document.all("validation").style.display = "block";
      }
      return false;
    }
  }
//-->
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
        <!-- #include file="_toolbar_portal_management.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If (lErr <> NO_ERR) And (lErr <> ERR_INVALID_NAME) Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
      <%Else%>
      <BR />
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV <%If lErr = NO_ERR Then Response.Write("STYLE=""display:none;""")%> class="validation" id="validation"><%If lErr = ERR_INVALID_NAME Then Response.Write("<LI>" & "There exists already a site with name """ & sName & """, please select a different name." & "<P>")%></DIV></FONT>
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
        <%Response.Write(asDescriptors(636)) 'Descriptor:Each site definition is distinguished by its unique name and description.%>&nbsp;
        <%Response.Write(asDescriptors(637)) 'Descriptor:It is best to provided a thorough description of the site definition, so any designer will know at a glance what the portal contains.%>&nbsp;
		<P>
        <%Response.Write(asDescriptors(566)) 'Descriptor:Enter an unique name and description of this portal.%>
      </FONT>
      <BR /><BR />
      <TABLE WIDTH=80% BORDER=0 CELLSPACING=0 CELLPADDING=0>
        <FORM NAME="FormName" ACTION="site_name.asp" METHOD="POST">

          <TR>
             <TD NOWRAP="1"><B><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>"> <%Response.Write(asDescriptors(306)) 'Descriptor:Name%> </FONT></B></TD>
             <TD>&nbsp;&nbsp;</TD>
             <TD ><B><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"> <%Response.Write(asDescriptors(22)) 'Descriptor:Description%> </FONT></B></TD>
           </TR>

           <TR>
             <TD NOWRAP="1" VALIGN=TOP><INPUT NAME=n   VALUE="<%=Server.HTMLEncode(sName)%>"  CLASS="textBoxclass" SIZE=35></INPUT></TD>
             <TD>&nbsp;&nbsp;</TD>
             <TD VALIGN=TOP ><TEXTAREA cols=55 name=des rows=5 CLASS="textBoxclass"><%=Server.HTMLEncode(sDesc)%></TEXTAREA></TD>
           </TR>

          <TR>
            <TD COLSPAN=3>
              <BR />
            </TD>
          </TR>

          <TR>
            <TD COLSPAN="3" ALIGN="left" NOWRAP>
              <INPUT name=sid     type=HIDDEN value="<%=sSiteId%>"></INPUT>
              <INPUT name=o       type=HIDDEN value="<%=Server.HTMLEncode(sOldName)%>"></INPUT>
              <INPUT name=confirm type=HIDDEN value="yes"></INPUT>
              <INPUT name=back    type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
              <INPUT name=next    type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>" onClick="return validateForm();"></INPUT> &nbsp;
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
<%
	Erase aSiteProperties
%>