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
Dim lStatus
Dim lCheckError

Dim sSiteId
Dim sCurrentSite
Dim sConfirm

Dim aSites()
Dim nCount
Dim i

    If oRequest("back") <> "" Then
		Erase aSiteProperties
		response.redirect "adminOverview.asp?section=2"
    End If

    sSiteId = oRequest("sid")
    sConfirm = oRequest("confirm")
    sCurrentSite = Application.Value("SITE_ID")

    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_SITE & " " & asDescriptors(623) 'Descriptor:Site Definition
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT

    lStatus = checkSiteConfiguration()

    If Len(sConfirm) > 0 Then

        If sSiteId = Application.Value("SITE_ID") And Len(sSiteId) > 0 Then
			Response.Redirect("modify_site.asp?sid=" & sSiteId)

        Else
			lCheckError = CheckSiteRepositories(sSiteId)

			If lCheckError = NO_ERR Then
				Response.Redirect("modify_site.asp?sid=" & sSiteId)
		    Else

                Select CASE lCheckError
	        	Case ERR_WRONG_DBALIAS_DEFINITION, ERR_PROPERTY_NOT_DEFINED
	        		sErrorMessage =  asDescriptors(714) 'Descriptor: Incorrect database alias definition for metadata tables

	        	Case ERR_WRONG_TABLE_VERSION
	        		sErrorMessage = asDescriptors(715) 'Descriptor: Database tables exist, but are not the correct version. Please enter another prefix.

	        	Case ERR_NO_TABLES_EXIST
	        		sErrorMessage = Replace(Replace(asDescriptors(925), "#1", "<B>" & sDBAlias & "</B>"), "#2", "<B>" & sPrefix & "</B>") 'Descriptor: There is no repository defined in '#1' with prefix '#2'.  Please enter a valid prefix.

	            Case Else
	                lErr = lCheckError

	        	End Select

	    	End If

		End If

    End If

    If lErr = NO_ERR Then
        lErr = getAllSites(aSites, nCount)
    End If

    'If no Site Definitions, continue automatically to create a new one:
    If lErr = NO_ERR Then
        If nCount = 0 Then
            If Len(oRequest("skip")) = 0 Then
                Response.Redirect("site_name.asp")
            Else
                Response.Redirect("adminOverview.asp?section=" & SECTION_PORTAL_MANAGEMENT)
            End If
        End If
    End If

    'Select current site:
    If lErr = NO_ERR Then

        'Select the first on the list
        sSiteId = aSites(0, 0)

        'Loop until we find the good one:
        For i = 0 To nCount - 1
            If StrComp(aSites(i, 0), sCurrentSite) = 0 Then
                sSiteId = sCurrentSite
                Exit For
            End If
        Next

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
		<!-- #include file="_toolbar_portal_management.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
	  <%If lErr <> NO_ERR Then
			Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor: Return to: 'Descriptor:Site definition
        Else%>
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV <%If lCheckError = NO_ERR Then Response.write "STYLE=""display:none;""" %> class="validation" id="validation"><LI><%=sErrorMessage%></LI></DIV></FONT>
      <BR />

      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
        <%Response.Write(Replace(asDescriptors(841), "#", "<B>" & GerPortalName(GetVirtualDirectoryName()) & "</B>")) 'You are currently configuring the Portal called #.  To configure a different Portal, or to create a new Portal,%> <A HREF="select_portal.asp"><%Response.write(asDescriptors(285))'click here%></A>.&nbsp;
        <P>
        <%Response.Write(asDescriptors(638)) 'Descriptor:The concept of a "site definition" is set of properties that define what the user will actually see through the browser when they access the URL of any given virtual directory.%>&nbsp;
        <%Response.Write(asDescriptors(639)) 'Descriptor:A portal can use one of any site definitions.%>
        <P>
        <%Response.Write(asDescriptors(640)) 'Descriptor:The following list contains site definitions that already exist in the current metadata.%>&nbsp;
        <%Response.Write(asDescriptors(641)) 'Descriptor:You may either select an existing site definition or create a new site definition in the current metadata.%>&nbsp;
        <P>
        <%Response.Write(asDescriptors(578)) 'Descriptor:Select the site definition to be used by this server:%>&nbsp;
      </FONT>
      <BR />

      <BR />
      <FORM NAME="FormSite" ACTION="select_site.asp">
        <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">

          <% If nCount > 0 Then %>

              <TR BGCOLOR="#6699CC">
                <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="18" /></TD>
                <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#FFFFFF"> <%Response.Write(asDescriptors(306)) 'Descriptor: Name%> </FONT></B></TD>
                <TD>&nbsp;&nbsp;</TD>
                <TD><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#FFFFFF"> <%Response.Write(asDescriptors(22)) 'Descriptor:Description%> </FONT></B></TD>
                <TD>&nbsp;&nbsp;</TD>
                <TD ALIGN=CENTER><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#FFFFFF"> </FONT></B></TD>
              </TR>

              <TR>
                <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#000000"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
              </TR>

              <% For i = 0 to nCount - 1 %>
                  <TR>
                    <TD WIDTH="1%"><INPUT NAME=sid TYPE=radio VALUE="<%=aSites(i, 0)%>"  <%If (sSiteId = aSites(i, 0)) Then Response.Write "CHECKED" %> /></TD>
                    <TD NOWRAP="1"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><%=Server.HTMLEncode(aSites(i, 1))%></FONT></TD>
                    <TD>&nbsp;&nbsp;</TD>
                    <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><%=Server.HTMLEncode(aSites(i, 2))%></FONT></TD>
                    <TD>&nbsp;&nbsp;</TD>
                    <% If aSites(i,3) = 0 Then %>
                        <TD ALIGN=CENTER><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><A HREF="site_name.asp?sid=<%=aSites(i, 0)%>"><%Response.Write(asDescriptors(353)) 'Descriptor:Edit%></A>&nbsp;/&nbsp;<A HREF="delete_site.asp?id=<%=aSites(i, 0)%>&n=<%=Server.URLEncode(aSites(i, 1))%>&tp=<%=TYPE_SITE%>"><%Response.Write(asDescriptors(249)) 'Descriptor:Delete%></A></TD>
                  	<% Else %>
                  		<TD>&nbsp;&nbsp;</TD>
                  	<% End If %>
                  </TR>

                  <TR>
                    <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
                  </TR>
              <% Next %>

              <TR>
                <TD COLSPAN="6" ALIGN="CENTER" HEIGHT="10" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="10" BORDER="0" ALT=""></TD>
              </TR>

          <% End If %>
          <TR>
            <TD COLSPAN="6"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
          </TR>

          <TR>
            <TD></TD>
            <TD COLSPAN="5">
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000">
                <A HREF="site_name.asp?sid=new"><%Response.write(asDescriptors(622)) 'Descriptor:Add a New Site Definition%></A>
              </FONT>
            </TD>
          </TR>
        </TABLE>

        <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
          <TR>
            <TD>
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="left" NOWRAP>
			  <INPUT name=confirm type=HIDDEN value="true"   ></INPUT>
              <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
              <INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
            </TD>
          </TR>

        </TABLE>

      </FORM>
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
	Erase aSites
%>