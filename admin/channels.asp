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
<!-- #include file="../CustomLib/ChannelsCuLib.asp" -->
<%
Dim sChannelsXML
Dim lStatus

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "channels.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_CHANNELS & " " & asDescriptors(471) 'Descriptor: Channels
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3

    lStatus = checkSiteConfiguration()

    If oRequest("back") <> ""  And (Strcomp(oRequest("page"),"channels") = 0) Then
        Call Response.Redirect("adminOverview.asp?section=3")
    End If

    If oRequest("next") <> "" Then
        Call Response.Redirect("deviceTypes.asp")
    End If

    'This is added to get sites to render tabs.  This will eventually be a transaction
    If lErr = NO_ERR Then
        lErr = cu_GetChannels(sChannelsXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "channels.asp", "", "", "Error calling cu_GetChannels", LogLevelTrace)
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
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT)%>">
          <%Call Response.Write(asDescriptors(820)) 'Descriptor: A channel is a visual element within the subscription portal.  Is presents the project information in a way so navigation is easier for the end user.%>&nbsp;
          <%Call Response.Write(asDescriptors(821)) 'Descriptor: Each channel is mapped to one folder of services in the Object Repository.%><BR/>
        <BR />
          <%Call Response.Write(asDescriptors(719)) 'Descriptor: & In this page you can edit or create new channels for this site:%><BR/>
        </FONT>
        <BR/>
        <%Call RenderChannelsList(sChannelsXML)%>

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD COLSPAN="2">
              <BR />
            </TD>
          </TR>

          <TR>
            <FORM ACTION="channels.asp" id=form1 name=form1>
              <TD ALIGN="left" NOWRAP WIDTH="1%">
                <BR /><INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
                <INPUT name=page type=hidden value="channels"></input>
              </TD>
              <TD ALIGN="left" NOWRAP WIDTH="98%">
                <BR /><INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
              </TD>
            </FORM>
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