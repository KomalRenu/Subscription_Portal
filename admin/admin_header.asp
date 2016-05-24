<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<TABLE WIDTH="100%" BORDER="0" BGCOLOR="#cc0000" CELLSPACING="0" CELLPADDING="0">
  <TR>
    <TD WIDTH="1%" ALIGN="CENTER" VALIGN="MIDDLE" ROWSPAN="3">
      <%If (lStatus = CONFIG_OK) Then Response.write("<B><A HREF=""../"" TARGET=""portal"">")%><IMG SRC="../images/home.gif" WIDTH="22" HEIGHT="21" BORDER="0" ALT="" /><BR />
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" <%
      If (lStatus <> CONFIG_OK) Then
          Response.write(" COLOR=""#c2c2c2"" ")
      Else
          Response.write(" COLOR=""#ffffff"" ")
      End If%> ><%=GerPortalName(GetVirtualDirectoryName())%><%If (lStatus = CONFIG_OK) Then Response.write("</A></B>")%><BR>
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#ffffff"><%Call Response.write(asDescriptors(248)) 'Descriptor:Administrators Page%></FONT>
    </TD>
    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" BORDER="0" ALT="" /></TD>
    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="25" HEIGHT="1" BORDER="0" ALT="" /></TD>
    <TD WIDTH="96%" ALIGN="LEFT">
      <TABLE BORDER="0">
        <TR>
          <%Call  RenderAdminSection(asDescriptors(615), "welcome.asp", aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_ENGINE_CONFIG, True) 'Descriptor:Engine Configuration%>
          <%Call  RenderAdminSection(asDescriptors(616), "adminOverview.asp?section=2", aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_MD) = 0)) 'Descriptor:Site Management %>
          <%Call  RenderAdminSection(asDescriptors(580), "adminOverview.asp?section=3", aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SITE_MANAGEMENT, lStatus = CONFIG_OK) 'Descriptor:Site Management%>
          <%Call  RenderAdminSection(asDescriptors(721), "adminOverview.asp?section=4", aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES, lStatus = CONFIG_OK) 'Descriptor:Services Configuration%>
        </TR>
      </TABLE>
    </TD>
    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="27" BORDER="0" ALT="" /></TD>
    <TD WIDTH="1%" ALIGN="RIGHT" NOWRAP>
      <%If (aPageInfo(N_TOOLBARS_PAGE) And HELP_TOOLBAR) > 0 Then%>
        <A HREF="<%=aPageInfo(S_NAME_PAGE)%>?showHelp=0&<%=aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)%>">
      <%Else%>
        <A HREF="<%=aPageInfo(S_NAME_PAGE)%>?showHelp=1&<%=aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)%>">
      <%End If%>
      <IMG SRC="../images/questionMark.gif" WIDTH="17" HEIGHT="17" BORDER="0" ALT="Help" /></A>&nbsp;&nbsp;
    </TD>
  </TR>
  <TR>
    <TD WIDTH="1%" BGCOLOR="#cc0000" VALIGN="TOP"><IMG SRC="../images/insidecorner.gif" WIDTH="21" HEIGHT="21" BORDER="0" ALT="" /></TD>
    <TD WIDTH="1%" BGCOLOR="#FFFFFF"><IMG SRC="../images/1ptrans.gif" WIDTH="25" HEIGHT="1" BORDER="0" ALT="" /></TD>
    <TD WIDTH="97%" BGCOLOR="#FFFFFF" ROWSPAN="2" COLSPAN="3" VALIGN="MIDDLE">
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_LARGE_FONT)%>"><B>
        <%=aPageInfo(S_TITLE_PAGE)%>
      </B></FONT>
    </TD>
  </TR>
  <TR>
    <TD COLSPAN="2" BGCOLOR="#FFFFFF"><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="9" BORDER="0" ALT="" /></TD>
  </TR>
  <TR>
    <TD WIDTH="1%" BGCOLOR="#FFFFFF" ALIGN="CENTER" VALIGN="TOP"><IMG SRC="../images/1ptrans.gif" WIDTH="157" HEIGHT="1" BORDER="0" ALT="" /></TD>
    <TD COLSPAN="5" BGCOLOR="#FFFFFF"><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT="" /></TD>
  </TR>
</TABLE>
