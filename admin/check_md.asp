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
Dim aMDConn(1)
Dim sConfirmed
Dim bTablesReady
Dim lStatus


    'Check for actions cancelled:
    If oRequest("cancel") <> "" Then
        'Roll back values:
        lErr = setMDConn(aMDConn)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "check_md.asp", "", "", "Error calling setMDConn", LogLevelTrace)
        Erase aMDConn

        Response.Redirect("select_md.asp")
    End If

    aPageInfo(S_NAME_PAGE) = "select_md.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_MD & " " & asDescriptors(569) 'Descriptor:Metadata Connection
    aPageInfo(N_CURRENT_OPTION_PAGE) = 1

    lStatus = checkSiteConfiguration()

    'Check if the MDTables exist with the current configuration:
    If lErr = NO_ERR Then
        lErr = checkMDTables(bTablesReady)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "check_md.asp", "", "", "Error calling checkMDTables", LogLevelTrace)
    End If

    'The MD Tables are ok... continue picking up a site:
    If lErr = NO_ERR Then
        If bTablesReady Then
			Erase aMDConn
            Call Response.Redirect("select_site.asp")
        End If
    End If

    If lErr = NO_ERR Then
        'If we have already confirmed with the user, just continue creating the
        'tables
        sConfirmed = oRequest("confirm")

        'Check if the MDTables exist with the current configuration:
        If sConfirmed = "yes" Then
            lErr = CreateMDTables()
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, "", "", "check_md.asp", "", "", "Error calling createMDTables", LogLevelTrace)
            Else
				Erase aMDConn
                Call Response.Redirect("select_site.asp")
            End If
        End If

    End If

    'Present a message to the user requesting confirmation to create tables.
    'For the messages get current values:
    If lErr = NO_ERR Then
        lErr = getMDConn(aMDConn)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "check_md.asp", "", "", "Error calling getMDConn", LogLevelTrace)
    End If

    'Since the tables are not ready, mark as if they are missing:
    If lErr = NO_ERR Then
        lStatus = lStatus Or CONFIG_MISSING_MD
    End If


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
        <!-- #include file="_toolbar_engine_config.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(569), "select_md.asp") 'Descriptor: Return to: 'Descriptor:Metadata Connection %>
      <%Else%>
      <BR />
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>

      <TABLE BORDER="0" WIDTH="100%"  CELLSPACING="0" CELLPADDING="0">
        <FORM ACTION="check_md.asp">
          <TR>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
               <%Response.Write(Replace(Replace(asDescriptors(590), "#1", "<B>" & aMDConn(0) & "</B>"), "#2", "<B>" & aMDConn(1) & "</B>")) 'Descrtiptor:The Database "#1" does not contain Metadata tables with the prefix "#2".%>
               <BR />
               <%Response.Write(asDescriptors(591)) 'Descriptor:Do you wish to create them?%>
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
	Erase aMDConn
%>
