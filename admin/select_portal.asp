<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next

'files to be checked - select_portal.asp, delete_portal.asp, PortalConfigCuLib.asp, PortalConfigCoLib.asp, _toolbar_portal_management.asp


%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/PortalConfigCuLib.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->

<%
Dim lStatus
Dim aPortals()
Dim nCount
Dim i
Dim Redirect_String
Dim sPortalName
Dim sDefaultPortal
Dim sDeleteUrl
Dim sVirtualDirectory
Dim sVD

    If Request.Form("new_portal") <> "" Then
        sPortalName = oRequest("portal_name")

        lErr = CreateNewPortal(sPortalName)
        If lErr <> NO_ERR And lErr <> ERR_VD_ALREADY_EXIST Then
            Call LogErrorXML(aConnectionInfo, lErr, "", "", "select_portal.asp", "", "", "Error calling CreateNewPortal", LogLevelTrace)
        Else
            Response.Redirect "../../" & sPortalName & "/admin/select_site.asp"
        End If

    End If

    If Request.Form("back") <> "" Then
        Response.Redirect "select_site.asp"
    End If

    If Request.Form("next") <> "" Then
        sVD = oRequest("vd")
        Redirect_String =  "../../" & sVD & "/admin/select_site.asp"
        Response.redirect Redirect_String
    End If

    If lErr = NO_ERR Then
        lErr = cu_GetAllPortals(aPortals, nCount)
    End If

    'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_PORTAL & " " & asDescriptors(628)'Descriptor:Portal Management
    aPageInfo(N_CURRENT_OPTION_PAGE) = 2

    lStatus = checkSiteConfiguration()

    sVirtualDirectory = GetVirtualDirectoryName()

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<SCRIPT LANGUAGE=javascript>
<!--

  function validateForm() {
  var sMsg;
  var sValue;


    sMsg = "";
    sValue = FormPortal.portal_name.value;
    if (sValue == "" || isBlank(sValue)) {
      <%Call Response.Write("sMsg = """ & asDescriptors(699) & """") 'Descriptor:Please provide a name for the portal %>
    } else {
        if (checkIsAlphaNumeric(sValue) == false) {
            <%Call Response.Write("sMsg = """ & asDescriptors(881) & """;") 'Descriptor:Portal names cannot contain special characters. Please enter a name which only uses a-z, 0-9 or underscore. %>
        }
    }

    if (sMsg != "") {
      if(document.all){
         document.all("validation").innerHTML = sMsg;
         document.all("validation").style.display = "block";
      }
      return false;
    }
  }
//-->
</SCRIPT>

<!-- #include file="validationJS.asp" -->
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
        <!-- #include file="_toolbar_portal_management.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> NO_ERR Then %>
        <% Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(628), "select_portal.asp") 'Descriptor: Return to: 'Descriptor:Portal Management%>
      <%Else%>
        <BR />
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
          <%Response.Write(asDescriptors(666)) 'Descriptor:Once the Portal has been defined it may be necessary to modify, delete, or create a new subscription portal.%>&nbsp;
          <%Response.Write(asDescriptors(667)) 'Descriptor:The following Subscription Portals have been detected on this web server.%>&nbsp;
          <%Response.Write(asDescriptors(668)) 'Descriptor:Click next to a Portal to edit its configuration.%>&nbsp;
        </FONT>

        <BR />
        <BR />

      <FORM NAME="FormPortal" ACTION="select_portal.asp" METHOD="POST">

        <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">

          <TR BGCOLOR="#6699CC">
            <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="18" /></TD>
            <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#FFFFFF"> <%Response.Write(asDescriptors(306)) 'Descriptor: Name%> </FONT></B></TD>
            <TD>&nbsp;&nbsp;</TD>
            <TD><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#FFFFFF"> <%Response.Write(asDescriptors(105)) 'Descriptor:Details%> </FONT></B></TD>
            <TD>&nbsp;&nbsp;</TD>
            <TD ALIGN=CENTER>&nbsp;</TD>
          </TR>

          <TR>
            <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#000000"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
          </TR>

          <% For i=0 to nCount - 1 %>
              <TR>
                <TD WIDTH="1%"><INPUT NAME=vd TYPE=radio VALUE="<%=aPortals(i, 6)%>" <% If StrComp(aPortals(i,0), sVirtualDirectory, vbTextCompare) = 0 Then Response.write "CHECKED" %> /></TD>
                <TD NOWRAP="1"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#000000"><%=Server.HTMLEncode(aPortals(i,6))%></FONT></TD>
                <TD>&nbsp;&nbsp;</TD>
                <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#000000"><%Response.Write(asDescriptors(697)) 'Descriptor:Associated Site%>: <b><%=aPortals(i,1)%></b></FONT></TD>
                <TD>&nbsp;&nbsp;</TD>
                <%If nCount > 1 Then %>
                  <TD ALIGN=CENTER><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" COLOR="#000000"><A HREF="<%="delete_portal.asp?name=" & aPortals(i,0)%>"><%Response.write(asDescriptors(249))'Descriptor:Delete%></A></TD>
                <%End If%>
              </TR>

              <TR>
                <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
              </TR>
          <% Next %>

          <TR>
            <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="10" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="10" BORDER="0" ALT=""></TD>
          </TR>

        </TABLE>

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
          <TR>
            <TD>
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="CENTER" NOWRAP >
              <INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(543)) 'Descriptor:OK%>"></INPUT> &nbsp;
              <%If nCount > 1 Then %>
                  <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT>
              <%End If%>
            </TD>
          </TR>

          <TR>
            <TD>
              <BR />
            </TD>
          </TR>
        </TABLE>

        <BR />
        <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
          <TR><TD>
            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
              <B><%Response.Write(asDescriptors(629)) 'Descriptor:Create New Subscription Portal.%><Br></B>
          </TR>
          <TR><TD>
            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
              <%Response.Write(asDescriptors(632)) 'Descriptor:To create a new Subscription Portal, specify the following information:%>
            </FONT>
          </TR>
        </TABLE>
        <BR />

        <TABLE WIDTH="70%" BORDER="1" CELLSPACING="0" CELLPADDING="10">
          <TR >
            <TD>
              <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="2">
                <TR>
                  <TD WIDTH="1%"  NOWRAP VALIGN=MIDDLE ALIGN=LEFT><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" ><%Response.Write(asDescriptors(631)) & ":" 'Descriptor:Portal Name:%></FONT></TD>
                  <TD WIDTH="99%" NOWRAP VALIGN=MIDDLE ALIGN=LEFT><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >http://localhost/</FONT><INPUT NAME="portal_name"   VALUE=""  CLASS="textBoxclass" SIZE=35></INPUT><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" COLOR="#ff0000"></FONT></TD>
                </TR>
                <TR>
                  <TD ALIGN=LEFT></TD>
                  <TD>
                    <DIV STYLE="display:none;" class="validation" id="validation"></DIV>
                  </TD>
                </TR>
                <TR>
                  <TD COLSPAN=2 ALIGN=CENTER>
                    <INPUT name=new_portal type=submit class="buttonClass" value="<%Response.Write(asDescriptors(630)) 'Descriptor:Create New Portal%>" onClick="return validateForm();"></INPUT>
                  </TD>
                </TR>
              </TABLE>
            </TD>
          </TR>
        </TABLE>

      </FORM>

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
	Erase aPortals
%>