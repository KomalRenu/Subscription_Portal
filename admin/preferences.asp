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
Dim aSitePreferences()
Redim aSitePreferences(MAX_SITE_PROP)

Dim aLocales
Dim aValues
Dim aNewUsers
Dim aAttachments
Dim aUseDHTML
Dim aTimeZones

Dim sExpDate
Dim sExpCount
Dim lStatus

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "preferences.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_PREFERENCES & " " & asDescriptors(286) 'Descriptor:Preferences
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3

    lStatus = checkSiteConfiguration()

    If lErr = NO_ERR Then
        lErr = getSiteProperties(aSitePreferences)
    End If

	'Code to maintain compatibility with sites created using NCS version 7.1GA
	'If property for login mode does not exist then create it.
	If aSitePreferences(SITE_LOGIN_MODE) = "" Then
		lErr = CreateSiteLoginModeProperty()
	    If lErr = NO_ERR Then
			aSitePreferences(SITE_LOGIN_MODE) = "NC_NORMAL"  'Set the login mode to the default value
	    End IF
	End IF

	'If property for prompt block counts does not exist then create it.
	If aSitePreferences(SITE_ELEMENT_PROMPT_BLOCK_COUNT) = "" Or aSitePreferences(SITE_OBJECT_PROMPT_BLOCK_COUNT) = "" Then
		lErr = CreateSitePromptCountProperty()
	    If lErr = NO_ERR Then
			aSitePreferences(SITE_ELEMENT_PROMPT_BLOCK_COUNT) = "30"  'Set the prompt block count to the default value
			aSitePreferences(SITE_OBJECT_PROMPT_BLOCK_COUNT) = "30"  'Set the prompt block count to the default value
	    End IF
	End IF

	'If property for attachments does not exist then create it.
	If aSitePreferences(SITE_PROP_STREAM_ATTACHMENTS) = "" Then
		lErr = CreateStreamAttachmentsProperty()
	    If lErr = NO_ERR Then
			aSitePreferences(SITE_PROP_STREAM_ATTACHMENTS) = "1"  'Set the default value
	    End IF
	End IF

	'If property for prompt match case does not exist then create it
	If aSitePreferences(SITE_PROMPT_MATCH_CASE) = "" Then
		lErr = CreatePromptMatchCaseProperty()
	    If lErr = NO_ERR Then
			aSitePreferences(SITE_PROMPT_MATCH_CASE) = "1"  'Set the default value
	    End If
	End If

	'If property for I-server port does not exist then create it
	If aSitePreferences(SITE_AUTHENTICATION_SERVER_PORT) = "" Then
		lErr = CreateIserverPortProperty()
		If lErr = NO_ERR Then
			aSitePreferences(SITE_PROSITE_AUTHENTICATION_SERVER_PORT) = "0"  'Set the default value
		End If
	End If

	'If property for timezone does not exist then create it
	If aSitePreferences(SITE_PROP_TIMEZONE) = "" Then
		lErr = CreateTimeZoneProperty()
		If lErr = NO_ERR Then
			aSitePreferences(SITE_PROP_TIMEZONE) = "0"  'Set the default value
		End If
	End If


    If lErr = NO_ERR Then
        lErr = getLocales(aLocales)
    End If

    If lErr = NO_ERR Then
        Redim aNewUsers(1, 1)

        aNewUsers(0, 0) = "1"
        aNewUsers(0, 1) = asDescriptors(119) 'Descriptor:yes

        aNewUsers(1, 0) = "0"
        aNewUsers(1, 1) = asDescriptors(118) 'Descriptor:no
    End If

    If lErr = NO_ERR Then
        Redim aAttachments(1, 1)

        aAttachments(0, 0) = "1"
        aAttachments(0, 1) = asDescriptors(119) 'Descriptor:yes

        aAttachments(1, 0) = "0"
        aAttachments(1, 1) = asDescriptors(118) 'Descriptor:no
    End If

    If lErr = NO_ERR Then
        Redim aUseDHTML(2, 1)

        aUseDHTML(0, 0) = "2"
        aUseDHTML(0, 1) = asDescriptors(304) 'Descriptor: Determine automatically

        aUseDHTML(1, 0) = "1"
        aUseDHTML(1, 1) = asDescriptors(119) 'Descriptor:yes

        aUseDHTML(2, 0) = "0"
        aUseDHTML(2, 1) = asDescriptors(118) 'Descriptor:no
    End If


    If lErr = NO_ERR Then
        sExpDate = "12/31/" & Year(Now())
        sExpCount = "10"
        If aSitePreferences(SITE_PROP_NEW_EXPIRE) = "1" Then
            sExpDate = aSitePreferences(SITE_PROP_EXPIRE_VALUE)
        ElseIf aSitePreferences(SITE_PROP_NEW_EXPIRE) = "2" Then
            sExpCount = aSitePreferences(SITE_PROP_EXPIRE_VALUE)
        Else
            aSitePreferences(SITE_PROP_NEW_EXPIRE) = "0"
        End If
    End If

    If lErr = NO_ERR Then
    	PopulateTimeZones aTimeZones
	End If

    'Change reserved characters:
    asReservedChars = Array("*", "|", "<", """", "?", ">")

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
    if (FormPreferences.cache[1].checked) {
      if (FormPreferences.dir.value == "" || isBlank(FormPreferences.dir.value)) {
        <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(599) & """") 'Descriptor:Please provide a directory for the temp files %>
      } else {
        if (checkInvalidCharacters(FormPreferences.dir.value) == false) <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(600) & """ + invalidChars();") 'Descriptor:Please enter a path for the temp files directory without the following characters:  %>
      }
    }

    if (FormPreferences.exp[1].checked) {
      if (checkIsDate(FormPreferences.expDate.value) == false) <%Call Response.write("sMsg += ""<LI>" & asDescriptors(294) & """ ;") 'Descriptor:Please enter a valid date %>
    } else {
      if (FormPreferences.exp[2].checked) {
        if (checkIsNumeric(FormPreferences.expCount.value) == false) <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(601) & """;") 'Descriptor:Please enter a valid number for the number of days.;%>
      }
    }

    if (checkIsNumeric(FormPreferences.element_count.value) == false) <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(900) & """;") 'Descriptor:Please enter a valid number for the element prompt items.;%>
    if (checkIsNumeric(FormPreferences.object_count.value) == false) <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(901) & """;") 'Descriptor:Please enter a valid number for the object prompt items.;%>

    if (FormPreferences.is_normal.checked || FormPreferences.is_nt.checked) {
      if (FormPreferences.is_server_name.value == "" || isBlank(FormPreferences.is_server_name.value)){
		<%Call Response.write("sMsg += ""<LI>" & asDescriptors(896) & """ ;") 'Descriptor:Please enter the name of the Intelligence Server that you want to validate the users against. %>
	  }
    }

    if (!FormPreferences.nc_normal.checked && !FormPreferences.is_normal.checked && !FormPreferences.is_nt.checked)
		FormPreferences.LoginMode.value = "000";
    if (FormPreferences.nc_normal.checked && !FormPreferences.is_normal.checked && !FormPreferences.is_nt.checked)
		FormPreferences.LoginMode.value = "001";
    if (!FormPreferences.nc_normal.checked && FormPreferences.is_normal.checked && !FormPreferences.is_nt.checked)
		FormPreferences.LoginMode.value = "010";
    if (FormPreferences.nc_normal.checked && FormPreferences.is_normal.checked && !FormPreferences.is_nt.checked)
		FormPreferences.LoginMode.value = "011";
    if (!FormPreferences.nc_normal.checked && !FormPreferences.is_normal.checked && FormPreferences.is_nt.checked)
		FormPreferences.LoginMode.value = "100";
    if (FormPreferences.nc_normal.checked && !FormPreferences.is_normal.checked && FormPreferences.is_nt.checked)
		FormPreferences.LoginMode.value = "101";
    if (!FormPreferences.nc_normal.checked && FormPreferences.is_normal.checked && FormPreferences.is_nt.checked)
		FormPreferences.LoginMode.value = "110";
    if (FormPreferences.nc_normal.checked && FormPreferences.is_normal.checked && FormPreferences.is_nt.checked)
		FormPreferences.LoginMode.value = "111";

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
        <!-- #include file="_toolbar_site_preferences.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
      <%Else%>
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>
      <BR />
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
        <%Response.Write(asDescriptors(682)) 'Descriptor:The site preferences control how the default values are set for users of the Portal GUI.  You may change any of the settings by modifying the information that appears on this page.%>
        <P><%Response.Write(asDescriptors(681)) 'Descriptor:Select the settings for the remainder of the site properties.%>
      </FONT>
      <BR />
      <BR />
      <FORM ACTION="modify_preferences.asp" Name="FormPreferences">
      <TABLE  WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0 >
      <TR>
        <TD COLSPAN=2>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <B><%Response.Write(asDescriptors(576)) 'Descriptor:Project Settings:%></B>
          </FONT>
        </TD>
      </TR>

      <TR>
        <TD COLSPAN=2 ALIGN="CENTER" HEIGHT="1" BGCOLOR="#cccccc"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
      </TR>

      <TR>
        <TD VALIGN=CENTER>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <%Response.Write(asDescriptors(564)) 'Descriptor:Enable creation of new users:%>
          </FONT>
        </TD>
        <TD>
          <% Call RenderDropDownList("usrs", aNewUsers, aSitePreferences(SITE_PROP_NEW_USERS), "") %>
        </TD>
      </TR>

      <!--
      <TR>
        <TD VALIGN=CENTER>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <%Response.Write(asDescriptors(560)) 'Descriptor:Default new users Locale:%>
          </FONT>
        </TD>
        <TD>
          <% Call RenderDropDownList("locale", aLocales, aSitePreferences(SITE_PROP_NEW_LOCALE) & ";" & aSitePreferences(SITE_PROP_GUI_LANG), "") %>
        </TD>
      </TR>
      -->

      <TR>
        <TD VALIGN=CENTER>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <%Response.Write(asDescriptors(570)) 'Descriptor:New users expiration:%>
          </FONT>
        </TD>
        <TD>
          <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
            <TR>
              <TD WIDTH="10%" NOWRAP VALIGN=CENTER><INPUT NAME=exp TYPE=radio VALUE="0" <%If aSitePreferences(SITE_PROP_NEW_EXPIRE) = "0" Then Response.Write "CHECKED" %> /><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><%Response.Write(asDescriptors(571)) 'Descriptor:No expiration%></FONT></TD>
              <TD WIDTH="35%">&nbsp;</TD>
              <TD WIDTH="10%" NOWRAP VALIGN=CENTER><INPUT NAME=exp TYPE=radio VALUE="1" <%If aSitePreferences(SITE_PROP_NEW_EXPIRE) = "1" Then Response.Write "CHECKED" %> /><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><%Response.Write(asDescriptors(572)) 'Descriptor:On:%> </FONT><INPUT NAME=expDate class="textBoxClass" SIZE=12 VALUE="<%=sExpDate%>" /> <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">(mm/dd/yyyy)</FONT></TD>
              <TD WIDTH="35%">&nbsp;</TD>
              <TD WIDTH="10%" NOWRAP VALIGN=CENTER><INPUT NAME=exp TYPE=radio VALUE="2" <%If aSitePreferences(SITE_PROP_NEW_EXPIRE) = "2" Then Response.Write "CHECKED" %> /><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><%Response.Write(Replace(asDescriptors(556), "##", "</FONT><INPUT NAME=expCount class=""textBoxClass"" SIZE=2 VALUE=""" & sExpCount & """ /><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>")) 'Descriptor:After:  ## days%></FONT></TD>
            </TR>
          </TABLE>
        </TD>
      </TR>

      <TR>
        <TD VALIGN=CENTER>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <%Response.Write("Timezone") 'Descriptor:Timezone:%>
          </FONT>
        </TD>
        <TD>
          <% Call RenderDropDownList("timezone", aTimeZones, aSitePreferences(SITE_PROP_TIMEZONE), "") %>
        </TD>
      </TR>

      <TR>
        <TD COLSPAN=2 ALIGN="CENTER" HEIGHT="15" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="15" BORDER="0" ALT=""></TD>
      </TR>

      <TR>
        <TD COLSPAN=2>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <B><%Response.Write(asDescriptors(567)) 'Descriptor:GUI Settings:%></B>
          </FONT>
        </TD>
      </TR>

      <TR>
        <TD COLSPAN=2 ALIGN="CENTER" HEIGHT="1" BGCOLOR="#cccccc"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
      </TR>

      <TR>
        <TD VALIGN=CENTER>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <%Response.Write(asDescriptors(561)) 'Descriptor:Default Use DHTML:%>
          </FONT>
        </TD>
        <TD>
          <%Call RenderDropDownList("dhtml", aUseDHTML, aSitePreferences(SITE_PROP_USE_DHTML), "" ) %>
        </TD>
      </TR>

      <TR>
        <TD VALIGN=CENTER>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
           <%Response.Write(asDescriptors(835)) & ": " 'Descriptor: Cache Storage Mechanism:%>
          </FONT>
        </TD>
        <TD>
          <TABLE CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
            <TR>
              <TD VALIGN=CENTER NOWRAP><INPUT NAME=cache TYPE=radio VALUE="2" <%If aSitePreferences(SITE_PROP_PROMPT_CACHE) <> "1" Then Response.Write "CHECKED" %> /><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><%Response.Write(asDescriptors(696)) 'Descriptor:Session Variables%></FONT></TD>
              <TD>&nbsp;&nbsp;&nbsp;</TD>
              <TD NOWRAP VALIGN=CENTER><INPUT NAME=cache TYPE=radio VALUE="1" <%If aSitePreferences(SITE_PROP_PROMPT_CACHE) = "1" Then Response.Write "CHECKED" %> /><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><%Response.Write(asDescriptors(584))'Descriptor:Temp Files Directory:%> </FONT><INPUT NAME=dir class="textBoxClass" SIZE=40 VALUE="<%=Server.HTMLEncode(aSitePreferences(SITE_PROP_TMP_DIR))%>"></INPUT></TD>
              <TD>&nbsp;&nbsp;&nbsp;</TD>
            </TR>
          </TABLE>
        </TD>
      </TR>

      <TR>
        <TD VALIGN=CENTER>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
           <%Response.Write(asDescriptors(838)) & ":" 'Descriptor: Display summary page for personalized service:%>
          </FONT>
        </TD>
        <TD>
          <TABLE CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
            <TR>
              <TD WIDTH="10%" VALIGN=CENTER NOWRAP><INPUT NAME="summary" TYPE=radio VALUE="1" <%If aSitePreferences(SITE_PROP_SUMMARY_PAGE) = "1" Then Response.Write "CHECKED" %> /><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><%Response.Write asDescriptors(35) 'Descriptor:Always%></FONT></TD>
              <TD WIDTH="35%">&nbsp;</TD>
              <TD WIDTH="10%" NOWRAP VALIGN=CENTER><INPUT NAME="summary" TYPE=radio VALUE="2" <%If aSitePreferences(SITE_PROP_SUMMARY_PAGE) = "2" Then Response.Write "CHECKED" %> /><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><%Response.Write asDescriptors(836) 'Descriptor:only when there are more than 1 question%> </FONT></TD>
              <TD WIDTH="35%">&nbsp;</TD>
              <TD WIDTH="10%" NOWRAP VALIGN=CENTER><INPUT NAME="summary" TYPE=radio VALUE="3" <%If aSitePreferences(SITE_PROP_SUMMARY_PAGE) = "3" Then Response.Write "CHECKED" %> /><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>"><%Response.Write asDescriptors(36) 'Descriptor:never%></FONT></TD>
            </TR>
          </TABLE>
        </TD>
      </TR>


      <TR>
        <TD COLSPAN=2 ALIGN="CENTER" HEIGHT="15" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="15" BORDER="0" ALT=""></TD>
      </TR>

      <TR>
        <TD COLSPAN=2>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <B><%Response.Write(asDescriptors(553)) 'Descriptor:Administrator Info:%></B>
          </FONT>
        </TD>
      </TR>

      <TR>
        <TD COLSPAN=2 ALIGN="CENTER" HEIGHT="1" BGCOLOR="#cccccc"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
      </TR>

      <TR>
        <TD VALIGN=CENTER>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <%Response.Write(asDescriptors(431)) 'Descriptor:Email:%>
          </FONT>
        </TD>
        <TD>
          <INPUT NAME=email class="textBoxClass" SIZE=40 VALUE="<%=Server.HTMLEncode(aSitePreferences(SITE_PROP_EMAIL))%> "></INPUT>
        </TD>
      </TR>

      <TR>
        <TD VALIGN=CENTER>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <%Response.Write(asDescriptors(573)) 'Descriptor:Phone number:%>
          </FONT>
        </TD>
        <TD>
          <INPUT NAME=phone class="textBoxClass" SIZE=40 VALUE="<%=Server.HTMLEncode(aSitePreferences(SITE_PROP_PHONE))%> "></INPUT>
        </TD>
      </TR>

      <TR>
        <TD COLSPAN=2 ALIGN="CENTER" HEIGHT="15" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="15" BORDER="0" ALT=""></TD>
      </TR>

      <TR>
        <TD COLSPAN=2>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <B><%Response.Write asDescriptors(904) 'Descriptor:Login Mode: %></B>
          </FONT>
        </TD>
      </TR>

      <TR>
        <TD COLSPAN=2 ALIGN="CENTER" HEIGHT="1" BGCOLOR="#cccccc"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
      </TR>

      <TR>
		<TD COLSPAN=2>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
				</TR>
					<TD WIDTH="80%" VALIGN=CENTER>
					  <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
					    <% Response.Write asDescriptors(892) 'Descriptor: 1. Narrowcast Server Normal (User Name and Password) " %>
					  </FONT>
					</TD>
					<TD>
					  <% If (aSitePreferences(SITE_LOGIN_MODE) = "NC_NORMAL") Or (aSitePreferences(SITE_LOGIN_MODE) = "NC_IS_NORMAL") Or (aSitePreferences(SITE_LOGIN_MODE) = "NC_NT_NORMAL") Or (aSitePreferences(SITE_LOGIN_MODE) = "NC_IS_NT_NORMAL") Then %>
					      <INPUT NAME="nc_normal" CHECKED TYPE="CHECKBOX"></INPUT>
					  <% Else %>
					      <INPUT NAME="nc_normal" TYPE="CHECKBOX"></INPUT>
					  <% End If %>
					</TD>
				</TR>

				<TR>
				  <TD WIDTH="80%" VALIGN=CENTER>
				    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
				      <%Response.Write asDescriptors(893) 'Descriptor:2. MicroStrategy Intelligence Server Normal (User Name and Password)%>
				    </FONT>
				  </TD>
				  <TD>
					  <% If (aSitePreferences(SITE_LOGIN_MODE) = "IS_NORMAL") Or (aSitePreferences(SITE_LOGIN_MODE) = "IS_NT_NORMAL") Or (aSitePreferences(SITE_LOGIN_MODE) = "NC_IS_NORMAL") Or (aSitePreferences(SITE_LOGIN_MODE) = "NC_IS_NT_NORMAL") Then %>
				        <INPUT NAME="is_normal" CHECKED TYPE="CHECKBOX"></INPUT>
				    <% Else %>
				        <INPUT NAME="is_normal" TYPE="CHECKBOX"></INPUT>
				    <% End If %>
				  </TD>
				</TR>

				<TR>
				  <TD WIDTH="80%" VALIGN=CENTER>
				    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
				      <%Response.Write asDescriptors(894) 'Descriptor:3. MicroStrategy Intelligence Server NT (NT User Name)%>
				    </FONT>
				  </TD>
				  <TD>
					  <% If (aSitePreferences(SITE_LOGIN_MODE) = "NT_NORMAL") Or (aSitePreferences(SITE_LOGIN_MODE) = "IS_NT_NORMAL") Or (aSitePreferences(SITE_LOGIN_MODE) = "NC_NT_NORMAL") Or (aSitePreferences(SITE_LOGIN_MODE) = "NC_IS_NT_NORMAL") Then %>
				        <INPUT NAME="is_nt" CHECKED TYPE="CHECKBOX"></INPUT>
				    <% Else %>
					      <INPUT NAME="is_nt" TYPE="CHECKBOX"></INPUT>
				    <% End If %>
				  </TD>
				</TR>

				<TR>
				  <TD WIDTH="80%" VALIGN=CENTER>
				    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
				      <%Response.Write asDescriptors(895) 'Descriptor:Please enter the MicroStrategy Intelligence Server name if option 2 or 3 is selected.%>
				    </FONT>
				  </TD>
				  <TD>
				    <INPUT NAME=is_server_name class="textBoxClass" SIZE=40 VALUE="<%=Server.HTMLEncode(aSitePreferences(SITE_AUTHENTICATION_SERVER_NAME))%>"></INPUT>
				  </TD>
				</TR>

				<TR>
				  <TD WIDTH="80%" VALIGN=CENTER>
				    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
				      <%Response.Write asDescriptors(9) 'Descriptor:Port%>
				    </FONT>
				  </TD>
				  <TD>
				    <INPUT NAME=is_server_port class="textBoxClass" SIZE=40 VALUE="<%=Server.HTMLEncode(aSitePreferences(SITE_AUTHENTICATION_SERVER_PORT))%>"></INPUT>
				  </TD>
				</TR>

			</TABLE>
		</TD>
      </TR>

      <TR>
        <TD COLSPAN=2 ALIGN="CENTER" HEIGHT="15" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="15" BORDER="0" ALT=""></TD>
      </TR>

      <TR>
        <TD COLSPAN=2>
          <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
            <B><%Response.Write asDescriptors(902) 'Descriptor:Prompt Settings: %></B>
          </FONT>
        </TD>
      </TR>

      <TR>
        <TD COLSPAN=2 ALIGN="CENTER" HEIGHT="1" BGCOLOR="#cccccc"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
      </TR>

      <TR>
		<TD COLSPAN=2>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
				</TR>
					<TD WIDTH="50%" VALIGN=CENTER>
					  <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
					    <%Response.Write  asDescriptors(898) 'Descriptor:Number of items to be returned for element prompts:%>
					  </FONT>
					</TD>
					<TD>
			          <INPUT NAME=element_count class="textBoxClass" SIZE=5 VALUE="<%=Server.HTMLEncode(aSitePreferences(SITE_ELEMENT_PROMPT_BLOCK_COUNT))%> "></INPUT>
					</TD>
				</TR>

				<TR>
					<TD WIDTH="50%" VALIGN=CENTER>
					  <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
					    <%Response.Write  asDescriptors(899) 'Descriptor:Number of items to be returned for object prompts:%>
					  </FONT>
					</TD>
					<TD>
			          <INPUT NAME=object_count class="textBoxClass" SIZE=5 VALUE="<%=Server.HTMLEncode(aSitePreferences(SITE_OBJECT_PROMPT_BLOCK_COUNT))%> "></INPUT>
					</TD>
				</TR>

				<TR>
					<TD WIDTH="50%" VALIGN=CENTER>
						<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
						    <%Response.Write  asDescriptors(926) 'Descriptor:Match case sensitivity by default:%>
	  				  </FONT>
					</TD>
					<TD>
			          <INPUT NAME="match_case" TYPE="checkbox" VALUE="1" <% If(aSitePreferences(SITE_PROMPT_MATCH_CASE) = 1) Then Response.Write "CHECKED=""1"""%></INPUT>
					</TD>
				</TR>

			</TABLE>
		</TD>
      </TR>

      </TABLE>
      <TABLE WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
      <TR>
        <TD>
          <BR />
        </TD>
      </TR>

      <TR>
        <TD ALIGN="left" NOWRAP>
          <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
          <INPUT name=next type=submit class="buttonClass" onClick="return validateForm();" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
        </TD>
      </TR>
      </TABLE>
      <INPUT TYPE=HIDDEN NAME=LoginMode VALUE=""></INPUT>
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
	Erase aSitePreferences
	Erase aLocales
	Erase aValues
	Erase aNewUsers
	Erase aAttachments
	Erase aUseDHTML
	Erase aTimeZones
%>