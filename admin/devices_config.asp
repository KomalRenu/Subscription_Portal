<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CustomLib/SiteConfigCuLib.asp" -->
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<%
Dim aSiteProperties()
Redim aSiteProperties(MAX_SITE_PROP)

Dim lStatus
Dim sFid
Dim sId
Dim sDescString

    'Get the Channels list request from the request object:
    If Len(oRequest("back")) > 0 Then
        Response.Redirect "deviceTypes.asp"
    End If

    If Len(oRequest("default")) > 0 Then
        sFid = oRequest("dfid")
        sId = oRequest("did")

        Response.Redirect "select_devices.asp?device=default&id=" & sId & "&fid=" & sFid
    End If


    If Len(oRequest("portal")) > 0 Then
        sFid = oRequest("pfid")
        sId = oRequest("pid")

        Response.Redirect "select_devices.asp?device=portal&id=" & sId & "&fid=" & sFid
    End If


    If lErr = NO_ERR Then
        lErr = getSiteProperties(aSiteProperties)
    End If


    If lErr = NO_ERR Then

        If Len(oRequest("deviceNext")) > 0 Then
            aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION) = CStr(oRequest("defValidation"))
            lErr = setSiteProperties(aSiteProperties, FLAG_PROP_GROUP_DEVICES)
            If lErr = NO_ERR Then
                lErr = ResetApplicationVariables()
            End If

            If lErr = NO_ERR Then
                Response.Redirect "is_config.asp"
            End If
        End If

    End If

    aPageInfo(S_TITLE_PAGE) = STEP_SITE_DEVICES & " " & asDescriptors(579) 'Descriptor:Site Devices
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3

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
        <!-- #include file="_toolbar_site_preferences.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
      <%Else%>
      <BR />
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
		<% Response.Write(asDescriptors(674)) 'Descriptor:A Device in this context is a specific way of delivering content to the end user.   The defaults provided are ## and ###.  You may change the defaults if have other needs.%>
      </FONT>
      <BR />
      <BR />
      <FORM ACTION="devices_config.asp" METHOD="POST">
      <TABLE  WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0 >
        <TR>
          <TD COLSPAN=2>
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
              <B><%Response.Write(asDescriptors(559)) 'Descriptor:Default device:%></B>
            </FONT>
          </TD>
        </TR>

        <TR>
          <TD COLSPAN="2" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#cccccc"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
        </TR>

        <TR>
          <TD VALIGN=TOP WIDTH="80%">
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <%If aSiteProperties(SITE_PROP_DEFAULT_DEV_ID) = "" Then
                Response.Write asDescriptors(552) 'Descriptor:(DISABLED)
              Else
                Response.Write aSiteProperties(SITE_PROP_DEFAULT_DEV_NAME)
            %>
            </FONT>
            <BR />
            <IMG SRC="../images/1ptrans.gif" ALT="" BORDER="" WIDTH="15" HEIGHT="10">
            <IMG SRC="../images/1ptrans.gif" ALT="" BORDER="" ALIGN="LEFT" WIDTH="15" HEIGHT="15">
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(501) 'Descriptor: Address format:%></FONT><BR />
            <TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0">
              <TR>
                  <TD BGCOLOR="#cccccc">
                      <TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>
                          <TR>
                              <TD BGCOLOR="#ffffff">
                                  <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
                                      <TR>
                                          <TD VALIGN=TOP>
                                              <INPUT TYPE="RADIO" NAME="defValidation" VALUE="email" <%If (Len(aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION)) = 0)(aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION) = "email") Then%>CHECKED<%End If%> />
                                          </TD>
                                          <TD VALIGN=TOP>
                                              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
                                                  <B><%Response.Write asDescriptors(502) 'Descriptor: E-mail%></B><BR />
                                                  <%Response.Write asDescriptors(503) 'Descriptor: This is the standard format for internet e-mail addresses.%><BR />
                                                  <%Response.Write asDescriptors(504) 'Descriptor: Format: xxxx@xxxxxx.xxx%>
                                              </FONT>
                                          </TD>
                                      </TR>
                                      <TR>
                                          <TD VALIGN=TOP>
                                              <INPUT NAME="defValidation" VALUE="number" TYPE="RADIO" <%If aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION) = "number" Then%>CHECKED<%End If%> />
                                          </TD>
                                          <TD VALIGN=TOP>
                                              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
                                                  <B><%Response.Write asDescriptors(505) 'Descriptor: Numeric%></B><BR />
                                                  <%Response.Write asDescriptors(524) 'Descriptor: Use this format for a string of numbers only.%><BR />
                                                  <%Response.Write asDescriptors(613) 'Descriptor: Format: any numbers and the following characters - ( )%>
                                              </FONT>
                                          </TD>
                                      </TR>
                                      <TR>
                                          <TD VALIGN=TOP>
                                              <INPUT NAME="defValidation" VALUE="none" TYPE="RADIO" <%If aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION) = "none" Then%>CHECKED<%End If%> />
                                          </TD>
                                          <TD VALIGN=TOP>
                                              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
                                                  <B><%Response.Write asDescriptors(611) 'Descriptor: No validation%></B><BR />
                                                  <%Response.Write asDescriptors(612) 'Descriptor: Address value can be any text string%><BR />
                                              </FONT>
                                          </TD>
                                      </TR>
                                  </TABLE>
                              </TD>
                          </TR>
                      </TABLE>
                  </TD>
                </TR>
              </TABLE>
              <%End If%>
          </TD>

          <TD VALIGN=TOP WIDTH="20%">
            <INPUT name=did     type=HIDDEN value="<%=aSiteProperties(SITE_PROP_DEFAULT_DEV_ID)%>" ></INPUT>
            <INPUT name=dfid    type=HIDDEN value="<%=aSiteProperties(SITE_PROP_DEFAULT_FOLDER_ID)%>" ></INPUT>
            <INPUT name=default type=SUBMIT value="<%Response.Write(asDescriptors(268)) 'Descriptor:Change%>" class="buttonClass" ></INPUT> &nbsp;
            <BR>
            <%
              If aSiteProperties(SITE_PROP_DEFAULT_DEV_ID) <> "" Then
               Response.Write "<A HREF=""modify_devices.asp?device=default&id=""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(562) & "</FONT></A>" 'Descriptor:Disable
              End If
            %>
          </TD>
        </TR>

        <TR>
          <TD COLSPAN="2" ALIGN="CENTER" HEIGHT="40" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="10" BORDER="0" ALT=""></TD>
        </TR>

        <TR>
          <TD COLSPAN=2>
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
              <B><%Response.Write(asDescriptors(574)) 'Descriptor:Portal Device:%></B>
            </FONT>
          </TD>
        </TR>

        <TR>
          <TD COLSPAN="2" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#cccccc"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
        </TR>

        <TR>
          <TD VALIGN=TOP WIDTH="80%">
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
              <%
                If aSiteProperties(SITE_PROP_PORTAL_DEV_ID) = "" Then
                  Response.Write asDescriptors(552) 'Descriptor:(DISABLED)
                Else
                  Response.Write aSiteProperties(SITE_PROP_PORTAL_DEV_NAME)
                End If
              %>
            </FONT>
          </TD>
          <TD VALIGN=TOP WIDTH="20%" >
            <INPUT name=pid    type=HIDDEN value="<%=aSiteProperties(SITE_PROP_PORTAL_DEV_ID)%>" ></INPUT>
            <INPUT name=pfid   type=HIDDEN value="<%=aSiteProperties(SITE_PROP_PORTAL_FOLDER_ID)%>" ></INPUT>
            <INPUT name=portal type=SUBMIT class="buttonClass" value="<%Response.Write(asDescriptors(268)) 'Descriptor:Change%>"></INPUT> &nbsp;
            <BR>
            <%
              If aSiteProperties(SITE_PROP_PORTAL_DEV_ID) <> "" Then
               Response.Write "<A HREF=""modify_devices.asp?device=portal&id=""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(562) & "</FONT></A>" 'Descriptor:Disable
              End If
            %>
          </TD>
        </TR>

      </TABLE>


      <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
        <TR>
          <TD COLSPAN="2">
            <BR /><BR />
          </TD>
        </TR>

        <TR>
          <TD ALIGN="left" NOWRAP WIDTH="1%">
            <INPUT NAME=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
          </TD>
          <TD ALIGN="left" NOWRAP WIDTH="98%">
            <INPUT NAME=deviceNext type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
          </TD>
        </TR>
      </FORM>
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
	Erase aSiteProperties
%>