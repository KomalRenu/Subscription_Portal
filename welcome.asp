<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<% On Error Resume Next %>
<!-- #include file="CommonDeclarations.asp" -->
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(283) 'Descriptor: Welcome%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
  <TABLE BORDER="0" WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
    <TR><TD HEIGHT="50"><IMG SRC="images/9_logo.gif" WIDTH="387" HEIGHT="28" ALT="" BORDER="0"></TD>
    <TR>
      <TD COLSPAN="4" BGCOLOR="#000000" BACKGROUND="images/Line_000000.gif"><IMG SRC="images/1ptrans.gif" WIDTH="320" HEIGHT="5" ALT="" BORDER="0" /></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" HEIGHT="25" NOWRAP BGCOLOR="#000000">&nbsp;</TD>
    </TR>
  </TABLE>

  <TABLE BORDER=0 CELLPADDING=20 CELLSPACING=0 WIDTH="100%">
    <TR>
      <TD VALIGN="top" WIDTH="1%" ><IMG SRC="images/1ptrans.gif" WIDTH="75" ALT="" BORDER="" /></TD>
      <TD VALIGN="top" WIDTH="99%">
        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_LARGE_FONT)%>">
          <B><%Call Response.Write(asDescriptors(874)) 'Welcome to Microstrategy Narrowcast Server Subscription Portal!%></B>
        </FONT>
        <P> </P>
        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>">
        <P><%Call Response.Write(asDescriptors(875)) 'MicroStrategy Narrowcast Server Portal allows users to easily subscribe to Narrowcast Server services through a Web browser.%> </P>
        <P><%Call Response.Write(asDescriptors(872)) 'Integration with the delivery engine enables users to specify service personalization, schedules and the device where they would like to receive the information. %></P>
        <P><%Call Response.Write(Replace(asDescriptors(873), "#", "<A HREF=""admin/"" target=""ncsadmin"">" & asDescriptors(285) & "</A>")) 'This site is not ready to be accessed by end-users. Please ask the administrator to configure this site. If you are the administrator, please click here.%> </P>

        <HR noShade SIZE=1>
        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD COLSPAN=3 NOWRAP>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#666666"            >Copyright © 1998-2012 <FONT COLOR="#cc0000">Micro<I>Strategy</I></FONT> Incorporated. <%Response.Write asDescriptors(7) 'Descriptor: All rights reserved. Confidential.%></FONT><BR>
            </TD>
          </TR>
          <TR>
            <TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#666666"><B><% Response.write asDescriptors(364)'Descriptor: MicroStrategy Hydra is a product of %></B></FONT></TD>
            <TD><IMG alt="MicroStrategy, Inc." border=0 height=25 src="images/125msi.gif" width=125 ></TD>
            <TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#666666"><B>(NASDAQ: MSTR)</B></FONT></TD>
          </TR>
        </TABLE>
      </TD>
    </TR>
  </TABLE>
</BODY>
</HTML>