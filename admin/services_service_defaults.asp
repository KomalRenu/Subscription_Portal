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

Dim aAnswers
Dim sAnswer

    If oRequest("back").count > 0 Then
		Erase aAnswers
        Response.Redirect("services_overview.asp")
    End If

    If lErr = NO_ERR Then
        lErr = ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)

        'The following information is not valid anymore:
        aSvcConfigInfo(SVCCFG_SS_ID) = ""
        aSvcConfigInfo(SVCCFG_SS_NAME) = ""
        aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) = ""
        aSvcConfigInfo(SVCCFG_QO_ID) = ""
        aSvcConfigInfo(SVCCFG_QO_NAME) = ""
        aSvcConfigInfo(SVCCFG_QO_PARENT_ID) = ""
        aSvcConfigInfo(SVCCFG_AQ_ID) = ""
        aSvcConfigInfo(SVCCFG_AQ_NAME) = ""
        aSvcConfigInfo(SVCCFG_AQ_PARENT_ID) = ""
        aSvcConfigInfo(SVCCFG_MAP_ID) = ""
        aSvcConfigInfo(SVCCFG_MAP_NAME) = ""

        aSvcConfigInfo(SVCCFG_STEP) = "default"

    End If

    If lErr = NO_ERR Then
        lErr = getSvcConfigDefaultAnswer(aSvcConfigInfo, sAnswer)
    End If

    'The rest of the ASP server site code goes here...
    'If lErr = NO_ERR Then
        'Redim aAnswers(3, 1)

        'aAnswers(0, 0) = ANSWER_DEFAULT
        'aAnswers(0, 1) = "Default"

        'aAnswers(1, 0) = ANSWER_SUBSCRIPTION_ID
        'aAnswers(1, 1) = ANSWER_SUBSCRIPTION_ID

        'aAnswers(2, 0) = ANSWER_USER_ID
        'aAnswers(2, 1) = ANSWER_USER_ID

        'aAnswers(3, 0) = ANSWER_ADDRESS_ID
        'aAnswers(3, 1) = ANSWER_ADDRESS_ID
    'End If

    'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_DEFAULT & " " & asDescriptors(768) 'Descriptor:Service Defaults
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForSvcConfig(aSvcConfigInfo)

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
              <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & "Services Overview", "services_overview.asp") 'Descriptor:Services Overview%>
            <%Else%>
              <BR />
              <% Call RenderSvcConfigPath(aSvcConfigInfo) %>
              <BR />
              <BR />
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <%Response.Write(asDescriptors(767)) 'You have selected to configure the above service.  You can either  keep the default preferences listed below or specify additional settings for individual subscription sets in the next steps.%><BR />
                <BR />
              </FONT>
              <BR />
              <FORM ACTION="services_service_modify.asp">
              <% Call RenderSvcConfigInputs(aSvcConfigInfo) %>
              <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
                <TR>
                  <TD>
                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                      <B><%Response.Write(asDescriptors(768)) 'Service Defaults:%></B>
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
	Erase aAnswers
%>