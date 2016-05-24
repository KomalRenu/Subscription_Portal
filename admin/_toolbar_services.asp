<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<TABLE WIDTH="157" BORDER="0" CELLSPACING="0" CELLPADDING="0">
  <TR>
    <TD WIDTH="1%" BGCOLOR="#999999"><IMG SRC="../images/1ptrans.gif" WIDTH="14" HEIGHT="30" ALT="" BORDER="0" /></TD>
    <TD WIDTH="99%" BGCOLOR="#999999"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#FFFFFF"><B><%Call Response.Write(asDescriptors(721)) 'Services Configuration%></B><BR /><%=Server.HTMLEncode(Application.Value("SITE_NAME"))%></FONT></TD>
  </TR>
</TABLE>

<TABLE WIDTH="157" BORDER="0" CELLSPACING="0" CELLPADDING="0" HEIGHT="100%" BGCOLOR="#666666" >
<TR>
  <TD WIDTH="14" ><IMG SRC="../images/1ptrans.gif" WIDTH="14" HEIGHT="1" ALT="" BORDER="0" /></TD>
  <TD VALIGN="TOP">
    <%
      Call renderServicesConfigToolbar()
    %>
    <BR />
    <BR />
  </TD>
  <TD WIDTH="5" ><IMG SRC="../images/1ptrans.gif" WIDTH="5" HEIGHT="1" ALT="" BORDER="0" /></TD>
</TR>
</TABLE>

