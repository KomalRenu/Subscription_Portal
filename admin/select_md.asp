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
Dim aDBAliases
Dim bFolderShared

Dim aMDConn(1)
Dim aCurrentMDConn(1)
Dim sDBAlias
Dim sPrefix
Dim sContinue

Dim i

    aPageInfo(S_NAME_PAGE) = "select_md.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_MD & " " & asDescriptors(569) 'Descriptor:Metadata Connection
    aPageInfo(N_CURRENT_OPTION_PAGE) = 1

    If oRequest("back") <> "" Then
		Erase aMDConn
		Erase aCurrentMDConn
		Erase aDBAliases
        Call Response.Redirect("welcome.asp")
    End If

    lStatus = checkSiteConfiguration()

	'Getting DBAlias and prefix from form
	sDBAlias = CStr(oRequest("dba"))
	sPrefix = CStr(oRequest("pre"))

	'Get MDConn values
	If lErr = NO_ERR Then
		lErr = getMDConn(aCurrentMDConn)
		If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "select_md.asp", "", "", "Error calling getMDConn", LogLevelTrace)
	End If

	If lErr = NO_ERR Then

		'User has submitted DBalias and a table prefix
		'Need to check if there tables for this MD and make
		'sure they have the correct version
		If Len(oRequest("next")) > 0 Then

			'If trying to use the same MD, just redirect to summary:
			'to next page, if it is, confirm before commiting changes:
			If ((Strcomp(aCurrentMDConn(0), sDBAlias) = 0) And (Strcomp(aCurrentMDConn(1), sPrefix) = 0) And (Len(sDBAlias) > 0)) Then
				Response.Redirect("adminSummary.asp?section=1")

			Else

				If (lStatus = CONFIG_OK) Then
					sContinue = CStr(oRequest("continue"))
				Else
					sContinue = "yes"
				End If

				If Len(sContinue) > 0 Then
					'Updating array whose values will be used
					'in confirmation form
					aMDConn(0) = sDBAlias
					aMDConn(1) = sPrefix

					'Save MD Information and continue to next page
					lErr = setMDConn(aMDConn)

					'Trying to use a new DBAlias, check it:
					lCheckError = CheckDBAlias(sDBAlias, sPrefix, REPOSITORY_MD)

					If lCheckError = NO_ERR Then

						Response.Redirect("adminSummary.asp?section=1")

					Else
						lErr = setMDConn(aCurrentMDConn)
						Select CASE lCheckError
							Case ERR_WRONG_DBALIAS_DEFINITION, ERR_PROPERTY_NOT_DEFINED
								sErrorMessage =  asDescriptors(714) 'Descriptor: Incorrect database alias definition for metadata tables

							Case ERR_WRONG_TABLE_VERSION
								sErrorMessage = asDescriptors(715) 'Descriptor: Database tables exist, but are not the correct version. Please enter another prefix.

							Case ERR_NO_TABLES_EXIST
								Response.Redirect("modify_md.asp?dba=" & server.URLEncode(oRequest("dba")) & "&pre=" & server.URLEncode(oRequest("pre")))

							Case Else
								lErr = lCheckError

						End Select
					End If
				End If
			End If
		End If
	End If


	'Get a list of DBAliases from the Engine:
	If lErr = NO_ERR Then
		lErr = getDBAliases(aDBAliases)
		If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "select_md.asp", "", "", "Error calling getDBAliases", LogLevelTrace)
	End If


	'Pick the Alias to use as default for each Connection only if there
	'is no error to show
	If lErr = NO_ERR And Not bDisplayError Then
		'Select default connections for Aurep
		If (Strcomp(sDBAlias, "") = 0) And (Strcomp(sPrefix, "") = 0) Then
			sDBAlias = selectDefaultDBALias(aDBAliases, Array(aCurrentMDConn(0), "SITE", "MD", "META"))
			sPrefix  = aCurrentMDConn(1)
		End If
	End If

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--

  function validateForm() {
  var sMsg

    sMsg = "";
    if (FormMD.pre.value != "") {
      if (checkInvalidCharacters(FormMD.pre.value) == false) <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(597) & """ + invalidChars();") 'Descriptor:Please enter a name without the following characters: %>
    }

    if (sMsg != "") {
      sMsg = sMsg + "<P>";
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
      <%If lErr <> NO_ERR Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(283), "welcome.asp") 'Descriptor: Return to:'Descriptor:Welcome" %>

      <%ElseIf Len(oRequest("next")) = 0 Then%>
      <BR />
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="<%If Not bDisplayError Then Response.write "display:none;"%>" class="validation" id="validation"><%If bDisplayError Then Response.write "<LI>" & sErrorMessage%><P></DIV></FONT>

      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
      <%Response.write(asDescriptors(659)) 'The Portal Repository stores configuration settings that define the end user's subscription portal interface.%>
      <%Response.write(asDescriptors(660)) 'This information includes:%>
      <UL>
        <LI><%Response.write(asDescriptors(661)) 'portal name and description%></LI>
        <LI><%Response.write(asDescriptors(817)) 'site definition%></LI>
        <UL>
          <LI><%Response.write(asDescriptors(662)) 'channel settings%></LI>
          <LI><%Response.write(asDescriptors(663)) 'published services%></LI>
          <LI><%Response.write(asDescriptors(664)) 'device folders%></LI>
          <LI><%Response.write(asDescriptors(665)) 'default preferences%></LI>
        </UL>
      </UL>
      <BR />
      </FONT>
      <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
        <FORM NAME="FormMD" ACTION="select_md.asp" METHOD="POST">
          <TR>
            <TD>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" COLOR="#000000">
                <%Response.write(asDescriptors(617)) 'Descriptor:Select the database connection for the Portal Repository.%>
              </FONT>
              <BR />
              <BR />
            </TD>
          </TR>

          <TR>
            <TD>
              <!--Start DBAlias list: -->
              <%Call displayDBAliasWidget(aDBAliases, asDescriptors, sPrefix, 0)%>
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
              <INPUT name=back    type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
              <INPUT name=next    type=submit class="buttonClass" onClick="return validateForm();" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;

            </TD>
          </TR>
		<TR>
		  <TD>
			<BR/>
		  </TD>
          </TR>
		<TR>
		  <TD>
		  <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
			<%Response.write(asDescriptors(933)) 'Descriptor:Tip: The system prefix allows you to create the repository tables for several Narrowcast Server systems in the same database location. For example, you might use the prefix TEST for a test system and the prefix PROD for a production system. A valid system prefix consists of characters and numbers, and starts with a character. The prefix cannot exceed eight characters in length.%>
			</FONT>
		  </TD>
          </TR>
        </FORM>
      </TABLE>
      <%Else%>
       <BR />
       <FORM ACTION="select_md.asp" METHOD="POST" id=form1 name=form1>
        <TABLE BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
          <TR>
            <TD>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
                <B><%Call Response.write(asDescriptors(322)) 'Descriptor:Warning!%></B><BR />
                <%Response.Write(asDescriptors(554)) 'Descriptor:After modifying the Metadata connection, you will need to re-configure the Site Info. %><BR />
                <%Response.Write(asDescriptors(563)) 'Descriptor:Do you want to continue?%>
              </FONT>
            </TD>
          </TR>
          <TR>
            <TD ALIGN=CENTER>
              <BR />
              <INPUT name=dba      type=HIDDEN value="<%=sDBAlias%>"></INPUT>
              <INPUT name=pre      type=HIDDEN value="<%=sPrefix%>"></INPUT>
              <INPUT name=continue type=HIDDEN value="yes"   ></INPUT>
              <INPUT name=next     type=submit class="buttonClass" value="<%Response.Write(asDescriptors(543)) 'Descriptor:Ok%>"></INPUT> &nbsp;
              <INPUT name=cancel   type=submit class="buttonClass" value="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT>
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
	Erase aMDConn
	Erase aCurrentMDConn
	Erase aDBAliases
%>