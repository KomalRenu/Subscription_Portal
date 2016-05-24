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
<%
Dim lStatus

    'Get the Channels list request from the request object:
    aPageInfo(S_TITLE_PAGE) = STEP_FINISH & " " & asDescriptors(442) 'Descriptor:Finish
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES

    lStatus = checkSiteConfiguration()

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
        <!-- #include file="_toolbar_services.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <BR />
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">

        <%Call Response.Write(asDescriptors(683)) 'You have successfully created and configured this portal.  All that remains is for you to test the site the end user will view.%><BR />
        <BR/>
        <%Call Response.Write(asDescriptors(684)) 'Click the Finish button to launch the browser and log in to the portal.%><BR />
        <BR/>
        <BR/>
      </FONT>

      <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
        <TR>
          <TD COLSPAN="2">
            <BR />
          </TD>
        </TR>

        <TR>
          <TD ALIGN="left" NOWRAP WIDTH="1%">
            <FORM ACTION="services_overview.asp"><INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;</FORM>
          </TD>
          <TD ALIGN="left" NOWRAP WIDTH="98%">
            <FORM ACTION="../default.asp" TARGET="portal"><INPUT type=submit class="buttonClass" value="<%Response.Write(asDescriptors(442)) 'Descriptor:Finish%>"></INPUT> &nbsp;</FORM>
          </TD>
        </TR>

      </TABLE>
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="1%">
        <!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>