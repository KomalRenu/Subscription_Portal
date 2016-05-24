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

Dim aSubsSets
Dim i

Dim sDefaultSubsSetId


    If lErr = NO_ERR Then
        lErr = ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)
    End If

    If lErr = NO_ERR Then
        lErr = getSubscriptionSets(aSvcConfigInfo, DYNAMIC_SS, aSubsSets)
        sDefaultSubsSetId = aSvcConfigInfo(SVCCFG_SS_CONFIG_ID)
    End If

    If lErr = NO_ERR Then
        'The following information is not valid anymore:
        aSvcConfigInfo(SVCCFG_SS_ID) = ""
        aSvcConfigInfo(SVCCFG_SS_NAME) = ""
        aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) = ""
        aSvcConfigInfo(SVCCFG_SS_MAP_ID) = ""
        aSvcConfigInfo(SVCCFG_QO_ID) = ""
        aSvcConfigInfo(SVCCFG_QO_NAME) = ""
        aSvcConfigInfo(SVCCFG_QO_PARENT_ID) = ""
        aSvcConfigInfo(SVCCFG_AQ_ID) = ""
        aSvcConfigInfo(SVCCFG_AQ_NAME) = ""
        aSvcConfigInfo(SVCCFG_AQ_PARENT_ID) = ""

        aSvcConfigInfo(SVCCFG_STEP) = "dynamic"
    End If

    'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_DYNAMIC & " " & asDescriptors(723)'"Dynamic Subscriptions"
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
<BODY BGCOLOR="#ffffff" TOPMARGIN=0 LEFTMARGIN=0 ALINK="#ff0000" LINK="#0000ff" VLINK="#0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
  <TR>
    <TD COLSPAN="6" HEIGHT="1%"><!-- begin header --><!-- #include file="admin_header.asp" --><!-- end header -->
    </TD>
  </TR>
  <TR>
    <TD WIDTH="1%" valign="top"><!-- begin toolbar --><!-- #include file="_toolbar_services.asp" --><!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG alt="" border=0 height=1 src="../images/1ptrans.gif" width=21 ></TD>

    <TD WIDTH="96%" valign="top">
      <%If lErr <> NO_ERR Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & "Services Overview", "services_overview.asp") 'Descriptor:Services Overview%>
      <%Else%>
        <BR />
        <% Call RenderSvcConfigPath(aSvcConfigInfo) %>
        <BR />
        <BR />
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
		  <% If IsEmpty(aSubsSets) Then
				Call Response.Write(asDescriptors(814)) 'This Service has no Dynamic Subscription Sets
			 Else
				Call Response.Write(asDescriptors(792)) 'This service has one or more dynamic subscription sets associated with it. Subscribers in theses subscription sets ill be stored in the Subscription Book Repository.
			 End If %>
          <BR />
        </FONT>

        <BR >
        <TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD VALIGN=TOP WIDTH=15>
              <IMG SRC="../images/arrow_right.gif" WIDTH=13 HEIGHT=13>
            </TD>
            <TD>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
                <B><%Call Response.Write(asDescriptors(791)) 'Defaults for All Dynamic Subscription Sets%></B><BR />
              </FONT>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <%Call Response.Write(asDescriptors(794)) 'To define a single configuration for all Dynamic Subscription Sets of this Service, click below. %></BR>
              </FONT>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <B><%Call Response.Write(asDescriptors(745)) 'Configured:%></B>
                <%
                    If Len(sDefaultSubsSetId) > 3 Then
                        Call Response.Write(asDescriptors(119)) 'Descriptor: Yes
                        Call Response.Write(" (<A HREF=""services_subsset.asp?ssid=" & DYNAMIC_SS & "&sscfgid=" & sDefaultSubsSetId & "&" & CreateRequestForSvcConfig(aSvcConfigInfo) & """ >" & asDescriptors(353) & "</A> / <A HREF=""delete_subsset.asp?ssid=" & DYNAMIC_SS & "&sscfgid=" & sDefaultSubsSetId & "&" & CreateRequestForSvcConfig(aSvcConfigInfo) & """ >" & asDescriptors(754) & "</A>)") 'edit/clear
                    Else
                        Call Response.Write(asDescriptors(118)) 'Descriptor: No
                        Call Response.Write(" (<A HREF=""services_subsset.asp?ssid=" & DYNAMIC_SS & "&sscfgid=" & NEW_OBJECT_ID & "&" & CreateRequestForSvcConfig(aSvcConfigInfo) & """ >" & asDescriptors(746) & "</A>)")
                    End If
                %>
              </FONT>
              <BR />
            </TD>
          </TR>

          <TR>
            <TD>
              <IMG SRC="../images/1ptrans.gif" HEIGHT="10" WIDTH="1">
            </TD>
          </TR>

          <TR>
            <TD VALIGN=TOP WIDTH=15>
              <IMG SRC="../images/arrow_right.gif" WIDTH=13 HEIGHT=13>
            </TD>
            <TD>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
                <B><%Call Response.Write(asDescriptors(795)) 'Configure Dynamic Subscription Sets Individually%></B><BR />
              </FONT>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <%Call Response.Write(asDescriptors(796)) 'To configure a dynamic subscription set individually, click on it below. Subscription sets that have not been configured will use the defaults from above.%></BR>
              </FONT><IMG SRC="../images/1ptrans.gif" HEIGHT="3" WIDTH="1"></BR>

<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
  <TR BGCOLOR="#c2c2c2">
    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="18" /></TD>
    <TD NOWRAP="1" ALIGN="LEFT"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  SIZE="<%=(N_SMALL_FONT)%>" ><%Call Response.Write(asDescriptors(731)) 'Dynamic Subscription Set%></FONT></B></TD>
    <TD>&nbsp;&nbsp;</TD>
    <TD NOWRAP="1" ALIGN="CENTER"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><%Call Response.Write(asDescriptors(825)) 'Configured%></FONT></B></TD>
    <TD>&nbsp;&nbsp;</TD>
    <TD>&nbsp;&nbsp;</TD>
  </TR>

  <TR>
    <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#000000"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
  </TR>

  <% If IsEmpty(aSubsSets) Then %>
    <TR>
      <TD COLSPAN=6 ALIGN=CENTER>
        <BR/>
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"><B>
          <%Call Response.Write(asDescriptors(814)) 'This Service has no Dynamic Subscription Sets%>
        </B></FONT>
        <BR />&nbsp;
      </TD>
    </TR>
  <% Else %>
    <% For i = 0 to UBound(aSubsSets) %>
      <TR>
        <TD WIDTH="1%"></TD>
        <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"><%=aSubsSets(i, 1)%></FONT></TD>
        <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"></FONT></TD>
        <TD ALIGN="CENTER"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"><%If Len(aSubsSets(i, 2)) > 0 Then Response.Write(asDescriptors(119)) Else Response.Write(asDescriptors(118)) %></FONT></TD>
        <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"></FONT></TD>
        <TD ALIGN="CENTER"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>">
        <%
          If Len(aSubsSets(i, 2)) > 0 Then
              Call Response.Write("<A HREF=""services_subsset.asp?ssid=" & aSubsSets(i, 0) & "&ssn=" & aSubsSets(i, 1) & "&sscfgid=" & aSubsSets(i, 2) & "&" & CreateRequestForSvcConfig(aSvcConfigInfo) & """ >" & asDescriptors(353) & "</A> / <A HREF=""delete_subsset.asp?ssid=" & aSubsSets(i, 0) & "&ssn=" & aSubsSets(i, 1) & "&sscfgid=" & aSubsSets(i, 2) & "&" & CreateRequestForSvcConfig(aSvcConfigInfo) & """ >" & asDescriptors(754) & "</A>" ) 'edit/clear
          Else
              Call Response.Write("<A HREF=""services_subsset.asp?ssid=" & aSubsSets(i, 0) & "&sscfgid=new&" & CreateRequestForSvcConfig(aSvcConfigInfo) & """ >" & asDescriptors(746) & "</A>" )
          End If
        %>
      </TR>

      <TR>
        <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
      </TR>
    <% Next %>
  <% End If %>

  <TR>
    <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="5" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
  </TR>
</TABLE>


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
              <FORM ACTION="services_static.asp" METHOD=POST>
                <INPUT type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
                <% RenderSvcConfigInputs(aSvcConfigInfo) %>
              </FORM>
            </TD>
            <TD ALIGN="left" NOWRAP WIDTH="98%">
              <FORM ACTION="services_overview.asp" METHOD=POST id=form2 name=form2>
                <INPUT type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
                <% RenderSvcConfigInputs(aSvcConfigInfo) %>
              </FORM>
            </TD>
          </TR>
        </TABLE>
        </FORM>

      <%End If %>
    </TD>

    <TD WIDTH="1%"><IMG alt="" border=0 height=1 src="../images/1ptrans.gif" width=21 ></TD>

    <TD WIDTH="1%"><!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
	Erase aSubsSets
%>