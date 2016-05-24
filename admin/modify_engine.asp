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
Dim sEngine
Dim sConfirmed
Dim lStatus
Dim bFolderShared

    'Check for actions cancelled:
    If oRequest("cancel") <> "" Then
        Response.Redirect("select_engine.asp")
    End If

    'Check for actions cancelled:
    If oRequest("back") <> "" Then
        Response.Redirect("welcome.asp")
    End If

    aPageInfo(S_NAME_PAGE) = "select_engine.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_ENGINE & " " & asDescriptors(582) 'Descriptor:Subscription Engine Location"
    aPageInfo(N_CURRENT_OPTION_PAGE) = 1

    lStatus = checkSiteConfiguration()

    'Read rest of request variables:
    sEngine = oRequest("se")

    'Check if it is trying to use the same SE, just go to the next page,
    'if not, set the new location:
    If (sEngine = Application.Value("SE")) Then
        Call Response.Redirect("select_md.asp")

    Else

	lErr = IsEngineFolderShared(sEngine)
    If lErr = 0 then		'0 means folder is shared, no error
		bFolderShared = True
	Elseif lErr = 1 then	'1 means folder not shared, but no error
		bFolderShared = False
		lErr = NO_ERR
	Else					'error in checking if folder is shared
		sErrorMessage = asDescriptors(912)   'Descriptor: Unable to check if the Subscription Engine folder is shared.
	End If


		If Len(oRequest("share")) <> 0 Then
			lErr = ShareSubEngineFolder(sEngine)
			If lErr = NO_ERR Then
				bFolderShared = True
			Else
				sErrorMessage =  FillEngineAndDriveInfo(asDescriptors(911),sEngine, GetEngineInstallDrive(sEngine))  'Descriptor: Unable to share drive.  Please manually share the #1 drive on machine #2 for that machine's local Administrators group
				bFolderShared = False
			End If
		End If

		If bFolderShared Then

			'If the configuration is already completed, verify with the user
			'before any changes take place, if not, just commit changes:
			If (lStatus = CONFIG_OK) Then
			    sConfirmed = oRequest("ok")
			Else
			    sConfirmed = "yes"
			End If

			'If the user has confirmed changes, set new Engine
			If sConfirmed = "yes" Then
				lErr = setSubscriptionEngine(sEngine)
				If lErr = NO_ERR Then
				    Call ResetApplicationVariables()
				    Call Response.Redirect("select_md.asp")
				Else
					sErrorMessage = asDescriptors(822) 'Descriptor: "There was an error while trying to establish connection to Subscription Engine. Either the server name is incorrect or Subscription engine is not running on this server."
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
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(582) , "select_engine.asp") 'Descriptor: Return to: 'Descriptor:Subscription Engine Location"%>
      <%ElseIf bFolderShared Then%>

      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>
      <BR />

        <TABLE BORDER=0 WIDTH=80% CELLSPACING=0 CELLPADDING=0>
          <FORM ACTION="modify_engine.asp">
          <TR>
            <TD>
            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
               <%Response.Write(asDescriptors(555)) 'Descriptor:After modifying the Subscription Engine location, you will need to re-configure the MD Connection and the Site Info. %><BR />
               <%Response.Write(asDescriptors(563)) 'Descriptor:Do you want to continue?%>
              </FONT>
            </TD>
          </TR>
          <TR>
            <TD ALIGN=CENTER>
              <BR />
              <INPUT name=origin type=HIDDEN value="<%=sOrigin%>"   ></INPUT>
              <INPUT name=se     type=HIDDEN value="<%=Server.HTMLEncode(sEngine)%>"   ></INPUT>
              <INPUT name=ok     type=HIDDEN value="yes"   ></INPUT>
              <INPUT             type=submit class="buttonClass" value="<%Response.Write(asDescriptors(543)) 'Descriptor:Ok%>"></INPUT> &nbsp;
              <INPUT name=cancel type=submit class="buttonClass" value="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT>
            </TD>
          </TR>
          </FORM>
        </TABLE>
      <%Else %>
          <TABLE BORDER=0 WIDTH=80% CELLSPACING=0 CELLPADDING=0>
          <FORM ACTION="modify_engine.asp">
          <TR>
            <TD>
            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
               <%Response.Write(FillEngineAndDriveInfo(asDescriptors(910),sEngine, GetEngineInstallDrive(sEngine))) 'Descriptor:The Subscription Engine's configuration settings could not be accessed. For the Subscription Portal to function, either the folder or the drive where the Subscription Engine is installed must be shared.  The Subscription Portal will not function if this requirement is not met.  Would you like to automatically share the #1 drive on machine #2 for that machine's local Administrators group?.%><BR />
               </FONT>
            </TD>
          </TR>
          <TR>
            <TD ALIGN=CENTER>
              <BR />
              <INPUT name=se     type=HIDDEN value="<%=Server.HTMLEncode(sEngine)%>"   ></INPUT>
              <INPUT name=ok     type=HIDDEN value="yes"   ></INPUT>
              <INPUT name=share  type=submit class="buttonClass" value="<%Response.Write(asDescriptors(543)) 'Descriptor:Ok%>"></INPUT> &nbsp;
              <INPUT name=cancel type=submit class="buttonClass" value="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT>
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