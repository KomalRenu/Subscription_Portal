<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/ServicesConfigCuLib.asp" -->
<%
Dim lStatus
Dim lCount
Dim aServicesInfo
Dim i

    Redim aSvcConfigInfo(MAX_SVCCFG_INFO)

    'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_OVERVIEW & " " & asDescriptors(362) '"Services"
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES

    lStatus = checkSiteConfiguration()

    If lErr = NO_ERR Then
        lErr = getConfiguredServices(aServicesInfo)
    End If

    If lErr = NO_ERR Then
        If Not IsEmpty(aServicesInfo) Then lCount = UBound(aServicesInfo) + 1
    End If

    If lCount = 0 Then
        If Len(oRequest("next")) > 0 Then
            Response.Redirect("services_select.asp")
        Else
            Response.Redirect("adminOverview.asp?section=" & SECTION_SERVICES)
        End If
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
      <%If lErr <> NO_ERR Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
      <%Else%>
        <BR />
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
          <%Response.Write(asDescriptors(740)) 'You have configured the following service(s). You may click on a service name to edit its configuration, or click on its folder name to configure other services in that folder.%>
          <BR />
        </FONT>
        <BR />
          <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
            <TR BGCOLOR="#6699CC">
              <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="18" /></TD>
              <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#FFFFFF"><%Response.Write(asDescriptors(366)) 'Service%></FONT></B></TD>
              <TD>&nbsp;&nbsp;</TD>
              <TD><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#FFFFFF"><%Response.Write(asDescriptors(142)) 'Parent folder:%></FONT></B></TD>
              <TD>&nbsp;&nbsp;</TD>
              <TD>&nbsp;&nbsp;</TD>
            </TR>

            <TR>
              <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#000000"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
            </TR>

              <% For i = 0 to lCount - 1 %>
                <TR>
                  <TD WIDTH="1%"></TD>
                  <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#000000"><%=aServicesInfo(i, 1)%></FONT></TD>
                  <TD></TD>
                  <TD><A HREF="services_select.asp?id=<%=aServicesInfo(i, 0)%>&sfid=<%=aServicesInfo(i, 2)%>"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#000000"><%=aServicesInfo(i, 3)%></FONT></A></TD>
                  <TD></TD>
                  <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#000000"><A HREF="services_static.asp?id=<%=aServicesInfo(i, 0)%>&n=<%=Server.URLEncode(aServicesInfo(i, 1))%>&sfid=<%=aServicesInfo(i, 2)%>"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#000000"><%Response.Write(asDescriptors(353)) 'edit%></FONT></A> / <A HREF="delete_serviceconfig.asp?id=<%=aServicesInfo(i, 0)%>&n=<%=Server.URLEncode(aServicesInfo(i, 1))%>&sfid=<%=aServicesInfo(i, 2)%>"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#000000"><%Response.Write(asDescriptors(754)) 'Clear%></FONT></A></FONT></TD>
                </TR>

                <TR>
                  <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="3"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="3" BORDER="0" ALT=""></TD>
                </TR>

                <TR>
                  <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
                </TR>


              <% Next %>

            <TR>
              <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="3"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="3" BORDER="0" ALT=""></TD>
            </TR>

            <TR>
              <TD WIDTH="1%"></TD>
              <TD COLSPAN="5"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><A HREF="services_select.asp"><%Response.Write(asDescriptors(842))'Add another service%></A></FONT></B></TD>
            </TR>

          </TABLE>
          <BR />

<!--
          <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
            <A HREF="services_select.asp">.Add another service.</A>
            <%Response.Write(Replace(asDescriptors(741), "#", "<B>" & asDescriptors(335) & "</B>")) 'To configure another service, # click below. 'Descriptor:Next%><BR/>
            <BR />
          </FONT>
 -->
        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD COLSPAN="2">
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="left" NOWRAP WIDTH="1%">
              <FORM ACTION="adminOverview.asp"><INPUT type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT><INPUT TYPE="HIDDEN" NAME="section" VALUE="<%=SECTION_SERVICES%>" />&nbsp;</FORM>
            </TD>
            <TD ALIGN="left" NOWRAP WIDTH="98%">
              <FORM ACTION="finish.asp"><INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;</FORM>
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
	Erase aServicesInfo
%>