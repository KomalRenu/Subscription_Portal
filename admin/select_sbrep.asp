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

<!-- #include file="dbaliases_widget.asp" -->
<%
Dim lStatus
Dim lCheckError
Dim sSiteId
Dim i

Dim aDBAliases
Dim aConnectionsInfo

Dim sDBAlias
Dim sPrefix

Dim sConfirm
Dim bSaveChanges
Dim bUnexpectedError

Dim aSiteProperties
Redim aSiteProperties(MAX_SITE_PROPERTIES)


    If oRequest("back") <> "" Then
		Erase aMDConn
		Erase aSiteProperties
		Erase aDBAliases
        Call Response.Redirect("select_aurep.asp")
    End If


    If oRequest("return") <> "" Then
		Erase aMDConn
		Erase aSiteProperties
		Erase aDBAliases
        Call Response.Redirect("welcome.asp")
    End If

    aPageInfo(S_TITLE_PAGE) = STEP_SITE_SBREP & " " & asDescriptors(581) 'Descriptor:Subscription Book Repository
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT

    lStatus = checkSiteConfiguration()
    sSiteId = Application.Value("SITE_ID")

    'Getting DBAlias and prefix from form
	sDBAlias = CStr(oRequest("dba"))
	sPrefix = CStr(oRequest("pre"))
	sConfirm = CStr(oRequest("confirm"))

	'Retrieve the connections for this site:
	lErr = GetSiteConnections(sSiteId, aConnectionsInfo)

	'Check if only need to confirm values.
	If lErr = NO_ERR Then

		If Len(sConfirm) > 0 Then

			'Only save if necessary:
			bSaveChanges = False

			'Confirm valid values:
			If sDBAlias <> aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_DSN) Or _
			   sPrefix <> aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_PREFIX) Then
				bSaveChanges = True
			End If
			lCheckError = CheckDBAlias(sDBAlias, sPrefix, REPOSITORY_SBREP)

			If lCheckError = NO_ERR Then

				If bSaveChanges Then

					aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_DSN) = sDBAlias
					aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_PREFIX) = sPrefix

					lErr = SetSiteConnections(sSiteId, aConnectionsInfo)

					If lErr = NO_ERR Then
						Call ResetApplicationVariables()
					End If

				End If

				If lErr = NO_ERR Then
					aSiteProperties(SITE_PROP_ID) = sSiteId
					lErr = SetSite(sSiteId)
				End If

				'Saved correctly, redirect to new page.
				If lErr = NO_ERR Then
					Call Response.Redirect("adminSummary.asp?section=2")
				End If

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


		'Get a list of DBAliases from the Engine:
		If lErr = NO_ERR Then
		    Call getDBAliases(aDBAliases)
		End If

		'Pick the Alias to use as default for each Connection
		If lErr = NO_ERR Then

		    If Len(sDBAlias) = 0 Then
		        sDBAlias = selectDefaultDBALias(aDBAliases, Array(aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_DSN), "sbr", "sub", "Portal", "rep"))
		        sPrefix = aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_PREFIX)
		    End If

		End If

	End If

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<SCRIPT LANGUAGE=javascript>
<!--

  function validateForm() {
  var sMsg

    sMsg = "";
    if (FormSBrep.pre.value != "") {
      if (checkInvalidCharacters(FormSBrep.pre.value) == false) sMsg += "<LI>Please enter a Table Prefix without the following characters: " + invalidChars();
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
	  <%If lErr <> NO_ERR Then
			Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(581), "select_sbrep.asp") 'Descriptor:SBR %>
      <%  Else%>
      <FONT FACE="Verdana,Arial,MS Sans Serif"  COLOR="#ff0000"><DIV <%If lCheckError = NO_ERR Then Response.write "STYLE=""display:none;"""%> class="validation" id="validation"><LI><%=sErrorMessage%></DIV></FONT>
      <BR />

      <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
        <FORM NAME="FormSBrep" ACTION="select_sbrep.asp" METHOD="POST">
          <TR>
            <TD>
              <FONT FACE="Verdana,Arial,MS Sans Serif" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" COLOR="#000000">
                <%Response.write(asDescriptors(657)) 'Descriptor:The Subscription Book Repository stores subscription data including end users, addresses, personalization and subscription sets.%>&nbsp;
                <%Response.write(asDescriptors(655)) 'Descriptor:The following table contains existing Database Aliases where data may be stored.%>&nbsp;
                <P>
                <%Response.write(asDescriptors(619)) 'Descriptor:Select the database alias where the Subscription Book Repository is located, or click Add a new database alias.%>&nbsp;
              </FONT>
              <BR />
              <BR />
            </TD>
          </TR>

          <TR>
            <TD>
              <!--Start DBAlias list: -->
              <%Call displayDBAliasWidget(aDBAliases, asDescriptors, sPrefix, 2)%>
              <!--End DBAlias list -->
            </TD>
          </TR>

          <TR>
            <TD>
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="left" NOWRAP>
			  <INPUT name=confirm type=HIDDEN value="true"   ></INPUT>
              <INPUT name=back   onClick="return validateForm();" type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
              <INPUT name=next   onClick="return validateForm();" type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
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
	Erase aSiteProperties
	Erase aDBAliases
%>