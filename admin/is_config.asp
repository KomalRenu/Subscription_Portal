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
<!-- #include file="../CustomLib/ISCuLib.asp" -->
<%
Dim aInformationSources
Dim lStatus

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "is_config.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_IS & " " & asDescriptors(568) 'Descriptor:Information Sources
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3

    lStatus = checkSiteConfiguration()

    'Get IS List:
    If lErr = NO_ERR Then
        lErr = getInfSources(aInformationSources)
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
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
      <%Else%>
      <BR />
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
        <%Response.Write(asDescriptors(676) & " ")  'Descriptor:For each information source you will need to specify whether the user will need to enter their authentication information, such as a user name and password.%>
      </FONT>
      <BR />
      <BR />
      <TABLE  WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0>
        <FORM ACTION="modify_is.asp">
        <TR BGCOLOR="#6699CC">
          <TD ROWSPAN=3 VALIGN=TOP><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR=#ffffff><%Response.Write(asDescriptors(568)) 'Descriptor:Information Source%></FONT></B></TD>
          <TD ROWSPAN=3 ALIGN=CENTER VALIGN=TOP><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR=#ffffff><%Response.Write(asDescriptors(586)) 'Descriptor:Use project credentials%></FONT></B></TD>
          <TD ROWSPAN=3 ALIGN=CENTER VALIGN=TOP><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR=#ffffff>&nbsp;&nbsp;&nbsp;</FONT></B></TD>
          <TD COLSPAN=2 ALIGN=CENTER><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR=#ffffff><%Response.Write(asDescriptors(909)) 'Descriptor:Use user credentials%></FONT></B></TD>
        </TR>

        <TR>
          <TD COLSPAN=2><IMG SRC="../images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT=""></TD>
        </TR>

        <TR BGCOLOR="#6699CC">
          <TD ALIGN=CENTER><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR=#ffffff><%Response.Write(asDescriptors(256)) 'Descriptor:Required%></FONT></B></TD>
          <TD ALIGN=CENTER><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR=#ffffff><%Response.Write(asDescriptors(257)) 'Descriptor:Optional%></FONT></B></TD>
        </TR>

        <% RenderISList(aInformationSources) %>

        <TR>
          <TD COLSPAN="4">
            <BR />
          </TD>
        </TR>

        <TR>
          <TD COLSPAN="4" ALIGN="left" NOWRAP>
            <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
            <INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
          </TD>
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
<%
	Erase aInformationSources
%>