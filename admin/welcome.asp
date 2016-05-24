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
	Dim lStatus
	Dim sDoesBackupExists
	Dim bAskToRestore
	Dim bFailedToRestore
	Dim bFailedToReset

	bFailedToRestore = false
	bFailedToReset = false
	If Len(Cstr(Request("restore_no"))) > 0 Then
		Response.Redirect("select_md.asp")
	End If

	If Len(Cstr(Request("restore_yes"))) > 0 Then
		lErr = RestoreBackupPropertyFiles()
		If lErr = NO_ERR Then
			lErr = ResetSubscriptionEngine()
			If lErr = NO_ERR Then
				lStatus = GetConfigSettings()
				Response.Redirect("select_md.asp")
			Else
				bFailedToReset = true
			End If
		Else
			bFailedToRestore = true
		End If
	End If

    aPageInfo(S_NAME_PAGE) = "welcome.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_WELCOME & " " & asDescriptors(587)'Descriptor:Welcome to Microstrategy Narrowcast Server
    aPageInfo(N_CURRENT_OPTION_PAGE) = 1

    lStatus = checkSiteConfiguration()

	bAskToRestore = false
	If Not bFailedToRestore and Not bFailedToReset Then
		If Len(CStr(Application.Value("SE"))) = 0  and Len(CStr(Application.Value("MD_CONN"))) = 0 Then
			lErr = FindBackupPropertyFiles(sDoesBackupExists)

			If lErr = NO_ERR Then
				If StrComp(UCase(sDoesBackupExists), "TRUE") = 0 Then
					bAskToRestore = true
				End If
			End If
		End If
	End If
%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>

<!-- #include file="../NSStyleSheet.asp" -->

<BODY BGCOLOR="FFFFFF" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
  <TR>
    <TD VALIGN="TOP" COLSPAN="5">
		<!-- begin header -->
		<!-- #include file="admin_header.asp" -->
		<!-- end header -->
    </TD>
  </TR>
  <TR HEIGHT="100%">
	<TD VALIGN="TOP" WIDTH="157" HEIGHT="100%" BGCOLOR="#666666">
		      <!-- begin toolbar -->
		        <!-- #include file="_toolbar_engine_config.asp" -->
		      <!-- end toolbar -->
   	</TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
	      <BR />
	      <IMG SRC="../images/repositories_large.gif" ALT="" BORDER="0" ALIGN="RIGHT" />
	      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>">
  	      <P>
	        <%Response.Write(asDescriptors(585))'Descriptor:This Portal empowers end users to easily self administrate Narrowcast Applications through a Web browser and so much other nice stuff that marketing guys know.%>
	      </P>
	        <P>
              <%Response.Write(asDescriptors(642))'As the Portal Designer you have the following responsibilities:%>
            <UL>
                <LI><%Response.Write(asDescriptors(643))'Configure and manage the subscription portal to allow end users to subscribe to a variety of services the Web%></LI>
                <LI><%Response.Write(asDescriptors(644))'Configure data sources and portal layouts%></LI>
                <LI><%Response.Write(asDescriptors(645))'Publish services and device types%></LI>
                <LI><%Response.Write(asDescriptors(646))'Select default devices for the subscription portal%></LI>
                <LI><%Response.Write(asDescriptors(647))'Specify information source properties and default portal preferences%></LI>
            </UL>
            </P>
            <P>
              <%Response.Write(asDescriptors(648))'The Portal Designer provides you with the capability to:%>
            <UL>
              <LI><%Response.Write(asDescriptors(650))'select or edit the portal configuration%></LI>
              <LI><%Response.Write(asDescriptors(651))'create or select the site definition%></LI>
              <LI><%Response.Write(asDescriptors(652))'configure static and dynamic service offerings%></LI>
            </UL>

			<% If bAskToRestore Then %>
				<P>
					<%Response.Write(asDescriptors(929)) 'There are backup property files found.  Would you like to restore and use configuration from the backup files?%>
				</P>
				<TABLE WIDTH="100%" CELLSPACING="0" CELLBORDER="0">
				  <TR>
					<TD ALIGN="left" NOWRAP>
					  <FORM ACTION="welcome.asp">
					  	<INPUT name=restore_yes type=submit class="buttonClass" value="<%Response.Write(asDescriptors(119))'Yes%>"></INPUT>
					  	<INPUT name=restore_no type=submit class="buttonClass" value="<%Response.Write(asDescriptors(118))'No%>"></INPUT> &nbsp;
					  </FORM>
					</TD>
				  </TR>
				</TABLE>


			<% Else %>
				<P>
				<% 	If bFailedToRestore Then
						Response.Write("<FONT COLOR=""RED""><B>" & asDescriptors(930) & "</B></FONT>") 'Failed to restore backup property files.  Please reconfigure your portal by clicking NEXT.
					ElseIf bFailedToReset Then
						Response.Write("<FONT COLOR=""RED""><B>" & asDescriptors(931) & "</B></FONT>") 'The backup property files have been restored.  Please restart MicroStrategy Subscription Server service before continuing.
					Else
					 	Response.Write(asDescriptors(653)) 'Click NEXT to begin.
					End If
				%>
				</P>

				<TABLE WIDTH="100%" CELLSPACING="0" CELLBORDER="0">
				  <TR>
					<TD>
					  <BR />
					</TD>
				  </TR>

				  <TR>
					<TD ALIGN="left" NOWRAP>
					  <FORM ACTION="select_md.asp"><INPUT type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;</FORM>
					</TD>
				  </TR>
				</TABLE>
			<%End If%>

	    </TD>



	    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    	<TD WIDTH="1%" VALIGN="TOP">
			<!-- #include file="help_widget.asp" -->
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>