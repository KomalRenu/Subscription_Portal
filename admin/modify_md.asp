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
Dim aMDConn(1)
'Dim aCurrentMDConn(1)
Dim lStatus

Dim sConfirmed

    'Check for actions cancelled:
    If oRequest("cancel").count > 0 Then
		Erase aMDConn
        Call Response.Redirect("select_md.asp?dba=" & oRequest("dba") & "&pre=" & oRequest("pre"))
    End If

    If oRequest("back").count > 0 Then
		Erase aMDConn
        Call Response.Redirect("welcome.asp")
    End If

    lStatus = checkSiteConfiguration()

    aPageInfo(S_NAME_PAGE) = "select_md.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_MD & " " & asDescriptors(569) 'Descriptor:Metadata Connection
    aPageInfo(N_CURRENT_OPTION_PAGE) = 1

    'Read request variables:
    aMDConn(0) = oRequest("dba")
    aMDConn(1) = oRequest("pre")

    'Check if it is trying to use a new MD Connection, if not, just set the logging and continue
    'to next page, if it is, confirm before commiting changes:
    If lErr = NO_ERR Then
		'If the configuration is already completed, verify with the user
        'before any changes take place, if not, just commit changes:
        sConfirmed = oRequest("confirm")

        'If confirmed to replace current MD, then commit changes:
        If Strcomp(sConfirmed, "modifyMD") = 0 Then
           'Save MD Information and continue to next page
           lErr = setMDConn(aMDConn)
           If lErr <> NO_ERR Then
				Call LogErrorXML(aConnectionInfo, lErr, "", "", "select_md.asp", "", "", "Error calling setMDConn", LogLevelTrace)
				sErrorHeader  = "Error saving the Metadata Connection"
                sErrorMessage = "An unexpected error happened while setting the new Metadata Connection. Please confirm the name """ & aMDConn(0) & """ is correct."
           Else
				'After saving changes, if everything went fine, redirect to the select_portal page
           		Call ResetApplicationVariables()
           		Erase aMDConn
				Response.Redirect("adminSummary.asp?section=1")
           End If

         ElseIf Strcomp(sConfirmed, "createMDTables") = 0 Then

            'Save MD Information and continue to next page
            lErr = setMDConn(aMDConn)

            If lErr = NO_ERR Then
                'Creating MD tables
                lErr = CreateMDTables()
                If lErr <> NO_ERR Then
                	Application.Value("MD_CONN") = ""
                Else
                    ''After saving changes, if everything went fine, tables will be saved and
                    'will continue to next page
                    Call ResetApplicationVariables()

                    If lErr = NO_ERR Then
                	    Response.Redirect("adminSummary.asp?section=1")
                	End If
                End If
            End If

         End If

    End If

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
        <!-- #include file="_toolbar_engine_config.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(569), "select_md.asp") 'Descriptor: Return to:'Descriptor:Metadata Connection  %>
      <%Else%>
     <BR />
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>

      <TABLE BORDER="0" WIDTH="100%"  CELLSPACING="0" CELLPADDING="0">
        <FORM ACTION="modify_md.asp">
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
			  <INPUT name=dba      type=HIDDEN value="<%=aMDConn(0)%>"></INPUT>
              <INPUT name=pre     type=HIDDEN value="<%=aMDConn(1)%>"></INPUT>
              <INPUT name=confirm type=HIDDEN value="createMDTables"   ></INPUT>
              <INPUT name=ok      type=submit class="buttonClass" value="<%Response.Write(asDescriptors(119)) 'Descriptor:Yes%>"></INPUT> &nbsp;
              <INPUT name=cancel  type=submit class="buttonClass" value="<%Response.Write(asDescriptors(118)) 'Descriptor:No%>"></INPUT>
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