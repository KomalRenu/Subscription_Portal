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
<!-- #include file="../CustomLib/ChannelsCuLib.asp" -->
<%
Const STEP_WIZARD_INTRO  = 0
Const STEP_WIZARD_FOLDER = 1
Const STEP_WIZARD_NAME   = 2
Const STEP_WIZARD_FINISH = 3

Dim sChannelsXML
Dim sFolderXML
Dim nNextStep

Dim aWizardInfo(12)
Dim bStatus
Dim lStatus

    'If cancelled, goto back to channels.asp
    If oRequest("cancel").count > 0 Then
		Erase aWizardInfo
        Response.Redirect "channels.asp"
    End If

    'Get request info:
    If lErr = NO_ERR Then
        lErr = ParseWizardRequest(oRequest, aWizardInfo)
    End If

    'This is added to get sites to render tabs.  This will eventually be a transaction
    If lErr = NO_ERR Then
        lErr = cu_GetChannels(sChannelsXML)
    End If

    'Define which is the next step:
    If Len(aWizardInfo(CHANNEL_NAVIGATION)) = 0 Then
        nNextStep = CInt(aWizardInfo(CHANNEL_STEP))

    ElseIf aWizardInfo(CHANNEL_NAVIGATION) = "next" Then
        nNextStep = CInt(aWizardInfo(CHANNEL_STEP)) + 1

    ElseIf aWizardInfo(CHANNEL_NAVIGATION) = "back" Then
        nNextStep = CInt(aWizardInfo(CHANNEL_STEP)) - 1

    ElseIf aWizardInfo(CHANNEL_NAVIGATION) = "finish" Then
        nNextStep = STEP_WIZARD_FINISH

    End If

    'Validate, there is no Step 0:
    If nNextStep = 0 Then nNextStep = STEP_WIZARD_FOLDER

    'We already applied the navigation:
    aWizardInfo(CHANNEL_NAVIGATION) = ""
    aWizardInfo(CHANNEL_STEP) = nNextStep

    'If On the final step:
    If lErr = NO_ERR Then
        If nNextStep = STEP_WIZARD_FINISH Then
            'Call modify:
            lErr = ModifyChannel(aWizardInfo)
            If lErr = NO_ERR Then
                Erase aWizardInfo
                Response.Redirect "channels.asp"
            ElseIf lErr = ERR_INVALID_NAME Then
                sErrorMessage = "A channel with the name """ & aWizardInfo(CHANNEL_NAME) & """ already exists, please select a different name."
                nNextStep = STEP_WIZARD_NAME
            End If

        End If
    End If

    'If a channel id was sent, make sure we already
    'have its information, if not request it:
    If lErr = NO_ERR Then

        'A channel was requested:
        If Len(aWizardInfo(CHANNEL_ID)) > 0 Then
            'Check if we already have its name (all channels must have a name:
            If Len(aWizardInfo(CHANNEL_NAME)) = 0 Then
                lErr = GetWizardChannelInfo(aWizardInfo, sChannelsXML, sErrorMessage)
            End If
        End If

    End If

    'Get the folder Info
    If lErr = NO_ERR Then

        If Len(aWizardInfo(CHANNEL_FOLDER_ID)) > 0 Then
            lErr = GetWizardFolderInfo(aWizardInfo, sErrorMessage)
            If lErr <> NO_ERR Then aWizardInfo(CHANNEL_FOLDER_ID) = ""
        End If

    End If

    'Wrap up page properties:
    If nNextStep = STEP_WIZARD_FOLDER Then

        aPageInfo(S_TITLE_PAGE) = STEP_CHANNELS_FOLDER & " " & asDescriptors(589)'Descriptor: Select the Channels Folder

        If lErr = NO_ERR Then
            'If it is a new channel, and the name is the default one, clean it back on this step:
            If Len(aWizardInfo(CHANNEL_ID)) = 0 Then
                If aWizardInfo(CHANNEL_NAME) = NewChannelName(aWizardInfo(CHANNEL_FOLDER_NAME)) Then
                    aWizardInfo(CHANNEL_NAME) = ""
                    aWizardInfo(CHANNEL_DESC) = ""
                End If
            End If

            'Finally, get the Parent folder XML:
            lErr = co_GetFolderXML(aWizardInfo(CHANNEL_PARENT_ID), ROOT_APP_FOLDER_TYPE, sFolderXML)
            If lErr <> NO_ERR Then aWizardInfo(CHANNEL_PARENT_ID) = ""

        End If


    ElseIf nNextStep = STEP_WIZARD_NAME Then

        aPageInfo(S_TITLE_PAGE) = STEP_CHANNELS_NAME & " " & asDescriptors(669)'Descriptor: Name & Description

    End If

    'Get the Channels list request from the request object:
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateWizardRequest(aWizardInfo)

    lStatus = checkSiteConfiguration()


%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<%If nNextStep = STEP_WIZARD_NAME Then%>

<SCRIPT LANGUAGE=javascript>
<!--

  function validateForm() {
  var sMsg

    sMsg = "";
    if ((FormChannelName.cnm.value == "") || isBlank(FormChannelName.cnm.value)) {
      sMsg += <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(720) & """;") 'Descriptor:Please enter a name without the following characters: %>
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
<%End If%>
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
    <%If lErr <> NO_ERR And lErr <> ERR_INVALID_NAME Then %>

 	    <%
 			If lErr = ERR_NO_ACTIVE_FOLDER Then sErrorMessage = asDescriptors(936)
 			If lErr = ERR_INACTIVE_FOLDER_ANCESTOR Then sErrorMessage = asDescriptors(942)
 	     	Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(471), "channels.asp") 'Descriptor: Return to:'Descriptor: Channels
 	  	%>
    <%Else%>
    <BR />
    <TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0>
    <TR>
      <TD><IMG SRC="../images/1ptrans.gif" WIDTH=10></TD>
      <TD VALIGN="TOP" WIDTH="100%">
        <%
        Call Response.write("<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """  COLOR=""#ff0000"">")
        If lErr <> NO_ERR Then
            Call Response.write("<DIV class=""validation"" id=""validation""><LI>" & sErrorMessage & "</DIV>")
        Else
            Call Response.write("<DIV STYLE=""display:none;"" class=""validation"" id=""validation""></DIV>")
        End If
        Call Response.write(" </FONT>")

        Select Case nNextStep
        Case STEP_WIZARD_INTRO
            Call RenderWizardWelcome(aWizardInfo)

        Case STEP_WIZARD_FOLDER
            Call Response.Write( "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """>" & asDescriptors(823) & "<BR /></FONT>") 'Descriptor: Defining a new channel requires selecting a folder from the Object Repository.  Select the channel folder to be displayed.
            Call RenderSelectFolder(aWizardInfo, sFolderXML)

        Case STEP_WIZARD_NAME
            Call RenderSelectName(aWizardInfo)

        Case Else:
            Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(250) & " channels.asp", "channels.asp") 'Descriptor: Return to:

        End Select
        %>
      </TD>
    </TR>
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
	Erase aWizardInfo
%>