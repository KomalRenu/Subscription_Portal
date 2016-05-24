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
Dim aSiteProperties()
Redim aSiteProperties(MAX_SITE_PROP)

Dim aAnswers
Dim sAnswer

    'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_TITLE_PAGE) = STEP_SERVICES & " " & asDescriptors(362)'Descriptor:Services
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES

    lStatus = checkSiteConfiguration()

    If lErr = NO_ERR Then
        lErr = getSiteProperties(aSiteProperties)
    End If

    'The rest of the ASP server site code goes here...
    If lErr = NO_ERR Then
        Redim aAnswers(2, 1)

        aAnswers(0, 0) = "subscription." & ANSWER_SUBSCRIPTION_ID
        aAnswers(0, 1) = ANSWER_SUBSCRIPTION_ID

        aAnswers(1, 0) = "subscription." & ANSWER_USER_ID
        aAnswers(1, 1) = ANSWER_USER_ID

        aAnswers(2, 0) = "subscription." & ANSWER_ADDRESS_ID
        aAnswers(2, 1) = ANSWER_ADDRESS_ID

        sAnswer = aSiteProperties(SITE_PROP_DEFAULT_ANSWER)

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
    <TD VALIGN="TOP">
      <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
        <TR>
          <TD COLSPAN="6" HEIGHT="1%">
            <!-- begin header -->
              <!-- #include file="admin_header.asp" -->
            <!-- end header -->
          </TD>
        </TR>
        <TR>
          <TD WIDTH="1%" VALIGN="TOP">
            <!-- begin toolbar -->
              <!-- #include file="_toolbar_services.asp" -->
            <!-- end toolbar -->
          </TD>

          <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

          <TD WIDTH="96%" VALIGN="TOP">
            <%If lErr <> NO_ERR Then %>
              <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
            <%Else%>
              <BR />
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
                <%Call Response.Write(asDescriptors(726) & " ") 'MicroStrategy Narrowcast Subscription Portal will publish the services in the folder you select.  However, some services require additional configuration.%>
                <%Call Response.Write(asDescriptors(727) & " ") 'Please specify the default configuration for these services below.%>
                <%Call Response.Write(Replace(asDescriptors(728), "#", "<B>" & asDescriptors(729) & "</B>")) 'To configure a specific service individually, click #  below. 'Descriptor:Configure Services%>
              </FONT>
              <BR />
              <BR />
              <FORM ACTION="services_modify.asp">
              <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
                <TR>
                  <TD>
                    <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                      <B><%Response.Write(asDescriptors(766)) 'Default Service Configuration%>:</B>
                    </FONT>
                  </TD>
                </TR>

                <TR>
                  <TD BGCOLOR="#c2c2c2"><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="1"></TD>
                </TR>

                <TR>
                  <TD><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="5"></TD>
                </TR>

                <TR>
                  <TD>
                    <!-- Begin: Default Values -->
                      <!-- #include file="default_service_config_widget.asp" -->
                    <!-- Ends: Default Values -->
                  </TD>
                </TR>


                <TR>
                  <TD><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="5"></TD>
                </TR>

                <TR>
                  <TD ALIGN="center">
                    <INPUT NAME="CONF" TYPE="SUBMIT" CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(729)) 'Configure Services%>" />
                  </TD>
                </TR>

              </TABLE>


              <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
                <TR>
                  <TD COLSPAN="2">
                    <BR />
                  </TD>
                </TR>

                <TR>
                  <TD ALIGN="left" NOWRAP WIDTH="1%">
                    <INPUT NAME="BACK" TYPE="SUBMIT" CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
                  </TD>
                  <TD ALIGN="left" NOWRAP WIDTH="98%">
                    <INPUT NAME="NEXT" TYPE="SUBMIT" CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
                  </TD>
                </TR>
              </TABLE>
              </FORM>
            <%End If %>
          </TD>

          <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

          <TD WIDTH="1%" VALIGN=TOP>
              <!-- #include file="help_widget.asp" -->
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
	Erase aSiteProperties
%>