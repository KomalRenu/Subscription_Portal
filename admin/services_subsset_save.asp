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

Dim aNormalQuestions
Dim aSlicingQuestions
Dim aExtraQuestions

Dim sNormalQOIds
Dim sSlicingQOIds
Dim sExtraQOIds

Dim i
Dim lCount


    If lErr = NO_ERR Then
        lErr = ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)

        'The following information is not valid anymore:
        aSvcConfigInfo(SVCCFG_QO_ID) = ""
        aSvcConfigInfo(SVCCFG_QO_NAME) = ""
        aSvcConfigInfo(SVCCFG_QO_PARENT_ID) = ""
        aSvcConfigInfo(SVCCFG_AQ_ID) = ""
        aSvcConfigInfo(SVCCFG_AQ_NAME) = ""
        aSvcConfigInfo(SVCCFG_AQ_PARENT_ID) = ""

    End If


    If lErr = NO_ERR Then
        lErr = GetSubscriptionSetConfig(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
    End If

    If lErr = NO_ERR Then
        lErr = SaveSubscriptionSetConfig(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)

        If lErr = NO_ERR Then
            Call DeleteCache(GetSvcConfigCacheName(aSvcConfigInfo), SVC_CONFIG_CACHE_FOLDER)
            Erase aNormalQuestions
			Erase aSlicingQuestions
			Erase aExtraQuestions

            If aSvcConfigInfo(SVCCFG_STEP) = DYNAMIC_SS Then
                Response.Redirect("services_dynamic.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo))
            Else
                Response.Redirect("services_static.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo))
            End If
        End If
    End If

    aPageInfo(S_NAME_PAGE) = "services_overview.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_OVERVIEW & " " & asDescriptors(362) '"Services"
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForSvcConfig(aSvcConfigInfo)

    lStatus = checkSiteConfiguration()



%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<!-- #include file="../NSStyleSheet.asp" -->

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
      <%If lErr <> 0 Then %>
            <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(362) , "services_overview.asp") 'Descriptor: Return to: 'Descriptor: Services%>
      <%Else%>
      <BR />
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>

      <TABLE BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
        <FORM ACTION="delete_subsset.asp" METHOD="POST">
        <% RenderSvcConfigInputs(aSvcConfigInfo) %>
          <TR>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
                Are you sure you want to clean the configuration of: <B><%=aSvcConfigInfo(SVCCFG_SS_NAME)%></B>?
              </FONT>
            </TD>
          </TR>
          <TR>
            <TD ALIGN=CENTER>
              <BR />
              <INPUT name=confirm type=HIDDEN value="yes"   ></INPUT>
              <INPUT name=ok      type=submit class="buttonClass" value="<%Response.Write(asDescriptors(543)) 'Descriptor:Ok%>"></INPUT> &nbsp;
              <INPUT name=cancel  type=submit class="buttonClass" value="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT>
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
	Erase aNormalQuestions
	Erase aSlicingQuestions
	Erase aExtraQuestions
%>