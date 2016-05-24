<%'** Copyright © 1996-2012  MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/ChannelsCoLib.asp" -->
<%
'*** aWizardInfo vars:
Const CHANNEL_STEP             =  0
Const CHANNEL_ACTION           =  1
Const CHANNEL_ID               =  2
Const CHANNEL_NAME             =  3
Const CHANNEL_DESC             =  4
Const CHANNEL_ACTIVE           =  5
Const CHANNEL_FOLDER_ID        =  6
Const CHANNEL_FOLDER_NAME      =  7
Const CHANNEL_FOLDER_DESC      =  8
CONST CHANNEL_PARENT_ID        =  9
Const CHANNEL_NAVIGATION       = 10


Function ParseWizardRequest(oRequest, aWizardInfo)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
Dim lErr
Dim chkHide

    On Error Resume Next
    lErr = NO_ERR

    'Clean all vars:
    aWizardInfo(CHANNEL_NAVIGATION)     = ""
    aWizardInfo(CHANNEL_STEP)           = ""
    aWizardInfo(CHANNEL_ACTION)         = ""
    aWizardInfo(CHANNEL_ID)             = ""
    aWizardInfo(CHANNEL_NAME)           = ""
    aWizardInfo(CHANNEL_DESC)           = ""
    aWizardInfo(CHANNEL_ACTIVE)         = ""
    aWizardInfo(CHANNEL_FOLDER_ID)      = ""
    aWizardInfo(CHANNEL_FOLDER_NAME)    = ""
    aWizardInfo(CHANNEL_FOLDER_DESC)    = ""
    aWizardInfo(CHANNEL_PARENT_ID)      = ""

    aWizardInfo(CHANNEL_STEP)           = Trim(CStr(oRequest("st")))
    aWizardInfo(CHANNEL_ACTION)         = Trim(CStr(oRequest("action")))
    aWizardInfo(CHANNEL_ID)             = Trim(CStr(oRequest("cid")))
    aWizardInfo(CHANNEL_NAME)           = Trim(CStr(oRequest("cnm")))
    aWizardInfo(CHANNEL_DESC)           = Trim(CStr(oRequest("cds")))
    aWizardInfo(CHANNEL_ACTIVE)         = Left(CStr(oRequest("cac")),1)
    aWizardInfo(CHANNEL_FOLDER_ID)      = Trim(CStr(oRequest("cfid")))
    aWizardInfo(CHANNEL_FOLDER_NAME)    = Trim(CStr(oRequest("fnm")))
    aWizardInfo(CHANNEL_FOLDER_DESC)    = Trim(CStr(oRequest("fds")))
    aWizardInfo(CHANNEL_PARENT_ID)      = Trim(CStr(oRequest("fid")))

    If oRequest("back") <> "" Then
        aWizardInfo(CHANNEL_NAVIGATION) = "back"
    ElseIf oRequest("next") <> "" Then
        aWizardInfo(CHANNEL_NAVIGATION) = "next"
    ElseIf oRequest("finish") <> "" Then
        aWizardInfo(CHANNEL_NAVIGATION) = "finish"
    End If

    'If no STEP info, assume step 0:
    If aWizardInfo(CHANNEL_STEP) = "" Then aWizardInfo(CHANNEL_STEP) = "0"

    'Check for Show Next Time setting
    chkHide = Trim(CStr(oRequest("chkHide")))
    If chkHide <> "" Then
        SetHideChannelWizardIntro(Left(chkHide,1))
    End If


    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AdminCuLib.asp", "ParseWizardRequest", "", "Error setting variables equal to Request variables", LogLevelError)
    End If

    ParseWizardRequest = lErr
    Err.Clear

End Function

Function CreateWizardRequest(aWizardInfo)
'********************************************************
'*Purpose:  Shows the path of the selected folder
'*Inputs:   sFolderXML: with the information of a document
'*Outputs:  none
'********************************************************
Dim sRequest

    On Error Resume Next

    sRequest = "st=" & aWizardInfo(CHANNEL_STEP)
    If aWizardInfo(CHANNEL_ACTION) <> "" Then sRequest = sRequest & "&action=" & aWizardInfo(CHANNEL_ACTION)
    If aWizardInfo(CHANNEL_ID) <> "" Then sRequest = sRequest & "&cid=" & aWizardInfo(CHANNEL_ID)
    If aWizardInfo(CHANNEL_NAME) <> "" Then sRequest = sRequest & "&cnm=" & Server.URLEncode(aWizardInfo(CHANNEL_NAME))
    If aWizardInfo(CHANNEL_DESC) <> "" Then sRequest = sRequest & "&cds=" & Server.URLEncode(aWizardInfo(CHANNEL_DESC))
    If aWizardInfo(CHANNEL_ACTIVE) <> "" Then sRequest = sRequest & "&cac=" & aWizardInfo(CHANNEL_ACTIVE)
    If aWizardInfo(CHANNEL_FOLDER_ID) <> "" Then sRequest = sRequest & "&cfid=" & aWizardInfo(CHANNEL_FOLDER_ID)
    If aWizardInfo(CHANNEL_FOLDER_NAME) <> "" Then sRequest = sRequest & "&fnm=" & Server.URLEncode(aWizardInfo(CHANNEL_FOLDER_NAME))
    If aWizardInfo(CHANNEL_FOLDER_DESC) <> "" Then sRequest = sRequest & "&fds=" & Server.URLEncode(aWizardInfo(CHANNEL_FOLDER_DESC))
    If aWizardInfo(CHANNEL_PARENT_ID) <> "" Then sRequest = sRequest & "&fid=" & aWizardInfo(CHANNEL_PARENT_ID)
    If aWizardInfo(CHANNEL_NAVIGATION) <> "" Then sRequest = sRequest & "&nav=" & aWizardInfo(CHANNEL_NAVIGATION)

    CreateWizardRequest = sRequest
    Err.Clear

End Function

Function cu_DeleteChannel(sChannelID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "cu_DeleteChannel"
    Dim lErrNumber
    Dim sSiteID

    lErrNumber = NO_ERR
    sSiteID = CStr(Application.Value("SITE_ID"))

    lErrNumber = co_DeleteChannel(sSiteID, sChannelID)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ChannelsCuLib.asp", PROCEDURE_NAME, "", "Error calling co_DeleteChannel", LogLevelTrace)
    End If

    cu_DeleteChannel = lErrNumber
    Err.Clear
End Function

Function getChannelXML(aWizardInfo)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim sChannelXML

    'Create XML:
    sChannelXML = ""
    sChannelXML = sChannelXML & "<mi>"
    sChannelXML = sChannelXML & "  <in>"
    sChannelXML = sChannelXML & "    <oi tp='" & TYPE_CHANNEL & "'>"
    sChannelXML = sChannelXML & "      <prs>"
    sChannelXML = sChannelXML & "        <pr id='NAME' v=""" & Server.HTMLEncode(aWizardInfo(CHANNEL_NAME)) & """ />"
    sChannelXML = sChannelXML & "        <pr id='DESC' v=""" & Server.HTMLEncode(aWizardInfo(CHANNEL_DESC)) & """ />"
    sChannelXML = sChannelXML & "        <pr id='CHANNEL_PUBLISHED' v='" & aWizardInfo(CHANNEL_ACTIVE) & "' />"
    sChannelXML = sChannelXML & "        <pr id='SERVICE_FOLDER_ID' v='" & aWizardInfo(CHANNEL_FOLDER_ID) & "' />"
    sChannelXML = sChannelXML & "      </prs>"
    sChannelXML = sChannelXML & "    </oi>"
    sChannelXML = sChannelXML & "  </in>"
    sChannelXML = sChannelXML & "</mi>"

    GetChannelXML = sChannelXML
    Err.Clear
End Function

Function GetWizardFolderInfo(aWizardInfo, strErrorMessage)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "GetWizardFolderInfo"
    Dim lErr
    Dim oFolderDOM
    Dim oFolder
    Dim sFolderXML

    lErr = NO_ERR

    If lErr = NO_ERR Then
        lErr = co_GetFolderXML(aWizardInfo(CHANNEL_FOLDER_ID), ROOT_APP_FOLDER_TYPE, sFolderXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), Err.source, "ChannelsCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetFolderXML", LogLevelTrace)
    End If

    'Load the XML File:
    If lErr = NO_ERR Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sFolderXML, oFolderDOM)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), Err.source, "ChannelsCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
        Else
	          'There must exist an ObjectInfo with the Folder Id:
	          Set oFolder = oFolderDOM.selectSingleNode("//fd[@id=""" & aWizardInfo(CHANNEL_FOLDER_ID) & """]")
	          If (Err.number <> NO_ERR) Or (oFolder Is Nothing) Then
	              lErr = ERR_XML_LOAD_FAILED
	              strErrorMessage = asDescriptors(942) 'Descriptor: Error loading folder information.  The folder may be inactive.
				  Call LogErrorXML(aConnectionInfo, CStr(lErr), strErrorMessage, "", "ChannelsCuLib.asp", PROCEDURE_NAME, "", "Error retrieving channel node", LogLevelError)
		  End If
        End If
    End If

    'Set the Info from the XML into the array:
    If lErr = NO_ERR Then
        aWizardInfo(CHANNEL_FOLDER_NAME) = oFolder.getAttribute("n")
        aWizardInfo(CHANNEL_FOLDER_DESC) = oFolder.getAttribute("des")
        aWizardInfo(CHANNEL_PARENT_ID)   = ofolder.parentNode.previousSibling.getAttribute("id")
    End If

    Set oFolder = Nothing
    Set oFolderDOM = Nothing

    GetWizardFolderInfo = lErr
    Err.Clear
End Function

Function GetWizardChannelInfo(aWizardInfo, sChannelsXML, sErrorMessage)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "GetWizardChannelInfo"
    Dim lErr
    Dim oChannelsDOM
    Dim oChannel

    lErr = NO_ERR

    lErr = LoadXMLDOMFromString(aConnectionInfo, sChannelsXML, oChannelsDOM)
    If lErr <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "ChannelsCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
    End If

    If lErr = NO_ERR Then
        'There must exist an ObjectInfo with the ChannelId:
        Set oChannel = oChannelsDOM.selectSingleNode("//oi[@id=""" & aWizardInfo(CHANNEL_ID) & """]")
        If oChannel Is Nothing Then
            lErr = ERR_XML_LOAD_FAILED
            sErrorMessage = "Error loading XML"
            Call LogErrorXML(aConnectionInfo, CStr(lErr), sErrorMessage, "", "ChannelsCuLib.asp", PROCEDURE_NAME, "", "Error retrieving channel node", LogLevelError)
        End If
    End If

    If lErr = NO_ERR Then

        aWizardInfo(CHANNEL_NAME)      = oChannel.getAttribute("n")
        aWizardInfo(CHANNEL_DESC)      = oChannel.getAttribute("des")

        'Select all the property nodes:
        aWizardInfo(CHANNEL_ACTIVE)    = GetPropertyValue(oChannel, "active")
        aWizardInfo(CHANNEL_FOLDER_ID) = GetPropertyValue(oChannel, "serviceFolderID")

    End If

    Set oChannelsDOM = Nothing
    Set oChannel = Nothing

    GetWizardChannelInfo = lErr
    Err.Clear
End Function

Function RenderChannelsList(sChannelsXML)
'********************************************************
'*Purpose:  Shows the path of the selected folder
'*Inputs:   sFolderXML: with the information of a document
'*Outputs:  none
'********************************************************
Dim oChannelsDOM
Dim oChannels
Dim oProp
Dim lErr

Dim i
Dim sImage
Dim sName
Dim sDesc
Dim sId
Dim sFeedback

    On Error Resume Next
    lErr = NO_ERR

    lErr = LoadXMLDOMFromString(aConnectionInfo, sChannelsXML, oChannelsDOM)
    If lErr <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ChannelsCuLib.asp", "RenderChannelsList", "", "Error calling LoadXMLDOMFromString: sChannelsXML", LogLevelTrace)
        Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#ff0000"">" & asDescriptors(272) & "</FONT></B>" 'Descriptor: One of the XML nodes needed is not available. Please ask the Administrator for more details.

    Else
        'Header:
        Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
        Response.Write "  <TR BGCOLOR=""#6699CC"">"
        Response.Write "    <TD><img src=""../images/1ptrans.gif"" WIDTH=""2"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
        Response.Write "    <TD WIDTH=""1%""><IMG SRC=""../images/active_status.gif"" WIDTH=""19"" HEIGHT=""10"" BORDER=""0"" ALT="""" /></TD>"
        Response.Write "    <TD><img src=""../images/1ptrans.gif"" WIDTH=""2"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
        Response.Write "    <TD NOWRAP=""1""><B><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">" & asDescriptors(306) & "</FONT></B></TD>" 'Descriptor: Name
        Response.Write "    <TD>&nbsp;&nbsp;</TD>"
        Response.Write "    <TD NOWRAP=""1""><B><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">" & asDescriptors(22) & "</FONT></B></TD>" 'Descriptor: Description
        Response.Write "    <TD>&nbsp;&nbsp;</TD>"
        Response.Write "    <TD NOWRAP=""1"" ALIGN=""CENTER""</TD>" 'Descriptor: Delete
        Response.Write "  </TR>"
        Response.Write "  <TR>"
        Response.Write "    <TD COLSPAN=""8"" ALIGN=""CENTER"" HEIGHT=""1"" BGCOLOR=""#000000""><img src=""../images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
        Response.Write "  </TR>"

        Set oChannels = oChannelsDOM.selectNodes("//oi")

        If oChannels.length = 0 Then
            Response.Write "<TR><TD COLSPAN=8 ALIGN=CENTER>"
            Response.Write "<BR/><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>"
            Response.Write asDescriptors(588) 'Descriptor:There are no channels defined for this site
            Response.Write "</B></FONT>"
            Response.Write "<BR />&nbsp;</TD></TR>"

        Else
            For i = 0 To oChannels.length - 1

                sId   = oChannels.item(i).getAttribute("id")
                sName = oChannels.item(i).getAttribute("n")
                sDesc = oChannels.item(i).getAttribute("des")
                If oChannels(i).selectSingleNode("prs/pr[@id='active']").getAttribute("v") = "0" Then
                    sImage = "inactive.gif"
                    sFeedback =  asDescriptors(475)'Descriptor:Inactive channel
                Else
                    sImage = "active.gif"
                    sFeedback = asDescriptors(474)'Descriptor:Active channel
                End If

                Response.Write "  <TR>"
                Response.Write "    <TD COLSPAN=""3"" ALIGN=""CENTER""><img src=""../images/" & sImage & """ WIDTH=""15"" HEIGHT=""15"" BORDER=""0"" ALT=""" & sFeedback &  """></TD>"
                Response.Write "    <TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#000000"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & Server.HTMLEncode(sName) & "</FONT></A></TD>"
                Response.Write "    <TD>&nbsp;&nbsp;</TD>"
                Response.Write "    <TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#000000"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & Server.HTMLEncode(sDesc) & "</FONT></TD>"
                Response.Write "    <TD>&nbsp;&nbsp;</TD>"
                Response.Write "    <TD ALIGN=CENTER><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#000000""><A HREF=""channels_wizard.asp?st=1&action=edit&cid=" & sId & """>" & asDescriptors(353) & "</A></FONT><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#000000"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """>&nbsp;/&nbsp;</FONT><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#000000""><A HREF=""delete_channel.asp?id=" & sId & "&n=" & Server.URLEncode(sName) & "&tp=" & TYPE_CHANNEL & """>" & asDescriptors(249) & "</A></TD>"  'Descriptor:Edit 'Descriptor:Delete
                Response.Write "  </TR>"
                Response.Write "  <TR>"
                Response.Write "    <TD COLSPAN=""8"" ALIGN=""CENTER"" HEIGHT=""1"" BGCOLOR=""#6699CC""><img src=""../images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
                Response.Write "  </TR>"

            Next
        End If

        Response.Write "  <TR>"
        Response.Write "    <TD COLSPAN=""8"" ALIGN=""CENTER"" HEIGHT=""5""><img src=""../images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""5"" BORDER=""0"" ALT=""""></TD>"
        Response.Write "  </TR>"
        Response.Write "  <TR>"
        Response.Write "    <TD COLSPAN=""3""><img src=""../images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=0 ALT=""""/></TD>"
        Response.Write "    <TD COLSPAN=""5""><A HREF=""channels_wizard.asp?action=add""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>"  & asDescriptors(473) & "</B></FONT></A></TD>" 'Descriptor: Add a new channel
        Response.Write "  </TR>"
        Response.Write "</TABLE>"

    End If

    Set oChannelsDOM = Nothing
    Set oChannels = Nothing

End Function

Function RenderWizardWelcome(aWizardInfo)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:  none
'********************************************************
Dim sTitle

    On Error Resume Next

    If aWizardInfo(CHANNEL_ACTION) = "add" Then
        sTitle = asDescriptors(476) 'Descriptor: Create New Channel
    Else
        sTitle = asDescriptors(477) 'Descriptor: Edit Channel
    End If

    Response.Write "      <FORM Action=""channels_wizard.asp"" METHOD=""POST"">"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""st"" VALUE=""0"">"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""action"" VALUE=""" & aWizardInfo(CHANNEL_ACTION) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cid"" VALUE=""" & aWizardInfo(CHANNEL_ID) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cnm"" VALUE=""" & aWizardInfo(CHANNEL_NAME) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cds"" VALUE=""" & aWizardInfo(CHANNEL_DESC) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cac"" VALUE=""" & aWizardInfo(CHANNEL_ACTIVE) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cfid"" VALUE=""" & aWizardInfo(CHANNEL_FOLDER_ID) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""fnm"" VALUE=""" & aWizardInfo(CHANNEL_FOLDER_NAME) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""fds"" VALUE=""" & aWizardInfo(CHANNEL_FOLDER_DESC) & """>"

    Response.Write "      <TABLE BORDER=""0"" WIDTH=100% CELLSPACING=0 CELLPADDING=0 >"
    Response.Write "        <TR>"
    Response.Write "          <TD>"
    Response.Write "            <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """ >"
    Response.Write "              <B>" & sTitle & "</B>"
    Response.Write "            </FONT>"
    Response.Write "          </TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >" & asDescriptors(283) & "!</FONT></TD>" 'Descriptor: Welcome
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD HEIGHT=1 BGCOLOR=""#000000""><IMG SRC=""../images/1ptrans.gif"" /></TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD HEIGHT=""10""><IMG SRC=""../images/1ptrans.gif"" /></TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >" & asDescriptors(478) & "</FONT></TD>" 'Descriptor: To create a new channel, you'll need to perform the following steps:
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD HEIGHT=""5""><IMG SRC=""../images/1ptrans.gif"" /></TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD>"
    Response.Write "            <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >"
    Response.Write "              <B>"
    Response.Write "                <OL>"
    Response.Write "                  <LI>" & asDescriptors(479) & "</LI>" 'Descriptor: Publish a folder of Services for the Channel
    Response.Write "                  <LI>" & asDescriptors(480) & "</LI>" 'Descriptor: Give the Channel a Name
    Response.Write "                </OL>"
    Response.Write "              </B>"
    Response.Write "            </FONT>"
    Response.Write "          </TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD>"
    Response.Write "            <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >"
    Response.Write "              <INPUT name=chkHide type=checkbox VALUE=1 CHECKED>" & asDescriptors(481) & "</INPUT>" 'Descriptor: Do not show this page again
    Response.Write "              <INPUT NAME=chkHide TYPE=HIDDEN   VALUE=0 />"
    Response.Write "            </FONT>"
    Response.Write "          </TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD HEIGHT=""10""><IMG SRC=""../images/1ptrans.gif""></TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD>"
    Response.Write "            <TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=0 CELLSPACING=0 >"
    Response.Write "              <TR>"
    Response.Write "                <TD WIDTH=100% ALIGN=""RIGHT"">"
    Response.Write "                  <INPUT name=next    type=SUBMIT class=""buttonClass"" value=""" & asDescriptors(335) & """ ></INPUT>"  'Descriptor:Next
    Response.Write "                </TD>"
    Response.Write "                <TD WIDTH=""20""><IMG SRC=""../images/1ptrans.gif"" WIDTH=""20"" /></TD>"
    Response.Write "                <TD ALIGN=""RIGHT"">"
    Response.Write "                  <INPUT name=cancel type=SUBMIT class=""buttonClass"" value=""" & asDescriptors(120)  & """></INPUT>"  'Descriptor:Cancel
    Response.Write "                </TD>"
    Response.Write "              </TR>"
    Response.Write "            </TABLE>"
    Response.Write "          </TD>"
    Response.Write "        </TR>"
    Response.Write "      </TABLE>"
    Response.Write "      </FORM>"

End Function

Function RenderSelectFolder(aWizardInfo, sFolderXML)
'********************************************************
'*Purpose:  Shows the path of the selected folder
'*Inputs:   sFolderXML: with the information of a document
'*Outputs:  none
'********************************************************
Dim lErr
Dim oFolderDOM
Dim oFolderContent
Dim i
Dim oSubFolders
Dim sFolderId
Dim sChecked
Dim sTitle
Dim oParent

    On Error Resume Next
    lErr = NO_ERR

    lErr = LoadXMLDOMFromString(aConnectionInfo, sFolderXML, oFolderDOM)
    If lErr <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.description), CStr(Err.source), "AdminCuLib.asp", "RenderChannelsList", "", "Error loading sChannelsXML", LogLevelError)
    Else
        Set oFolderContent = Nothing
        Set oFolderContent = oFolderDOM.selectSingleNode("mi/fct")
    End If

    If lErr <> NO_ERR Then
        Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">" & asDescritpros(1315) & "</FONT></B></TD>" 'Descriptor: One of the XML nodes needed is not available. Please ask the Administrator for more details.
    Else

        If aWizardInfo(CHANNEL_ACTION) = "add" Then
            sTitle = asDescriptors(476) 'Descriptor: Create New Channel
        Else
            sTitle = asDescriptors(477) 'Descriptor: Edit Channel
        End If


        Response.Write "      <FORM Action=""channels_wizard.asp"" METHOD=""POST"">"
        Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""st"" VALUE=""1"">"
        Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""action"" VALUE=""" & aWizardInfo(CHANNEL_ACTION) & """>"
        Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cid"" VALUE=""" & aWizardInfo(CHANNEL_ID) & """>"
        Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cnm"" VALUE=""" & Server.HTMLEncode(aWizardInfo(CHANNEL_NAME)) & """>"
        Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cds"" VALUE=""" & Server.HTMLEncode(aWizardInfo(CHANNEL_DESC)) & """>"
        Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cac"" VALUE=""" & aWizardInfo(CHANNEL_ACTIVE) & """>"

        Response.Write "      <TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH=""100%"" >"
        Response.Write "        <TR>"
        Response.Write "          <TD>"

        Call RenderFolderPath(aWizardInfo, sFolderXML)

        Response.Write "          </TD>"
        Response.Write "        </TR>"
        Response.Write "        <TR>"
        Response.Write "          <TD HEIGHT=""10""><IMG SRC=""../images/1ptrans.gif""></TD>"
        Response.Write "        </TR>"

        'Header:
        Response.Write "        <TR>"
        Response.Write "          <TD>"
        Response.Write "            <TABLE WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
        Response.Write "              <TR BGCOLOR=""#6699CC"">"
        Response.Write "                <TD><IMG SRC=""../images/1ptrans.gif"" WIDTH=""18"" /></TD>"
        Response.Write "                <TD><IMG SRC=""../images/1ptrans.gif"" WIDTH=""18"" /></TD>"
        Response.Write "                <TD NOWRAP=""1""><B><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">" & asDescriptors(306) & "</FONT></B></TD>" 'Descriptor: Name
        Response.Write "                <TD>&nbsp;&nbsp;</TD>"
        Response.Write "                <TD NOWRAP=""1""><B><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">" & asDescriptors(34) & "</FONT></B></TD>" 'Descriptor: Modified
        Response.Write "                <TD>&nbsp;&nbsp;</TD>"
        Response.Write "                <TD NOWRAP=""1""><B><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">" & asDescriptors(22) & "</FONT></B></TD>" 'Descriptor: Description
        Response.Write "              </TR>"
        Response.Write "              <TR>"
        Response.Write "                <TD COLSPAN=""7"" ALIGN=""CENTER"" HEIGHT=""1"" BGCOLOR=""#000000""><img src=""../images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
        Response.Write "              </TR>"

        'Subfolders:
        Set oSubFolders = oFolderContent.selectNodes("oi[@tp=""2""]")
        Set oParent = oFolderDOM.selectSingleNode("//fd[@id='" & aWizardInfo(CHANNEL_PARENT_ID) & "']")

        'Keep the folder ID
        sFolderId = aWizardInfo(CHANNEL_FOLDER_ID)
        aWizardInfo(CHANNEL_FOLDER_ID) = ""
        If sFolderId = "" Then sFolderId = oSubFolders.item(0).getAttribute("id")


        If oSubFolders.length > 0 Then

            For i = 0 To oSubFolders.length - 1

                aWizardInfo(CHANNEL_PARENT_ID) = oSubFolders.item(i).getAttribute("id")

                'If this is the selected folder, mark it:
                If sFolderId = aWizardInfo(CHANNEL_PARENT_ID) Then
                    sChecked = " CHECKED "
                Else
                    sChecked = ""
                End If

                Response.Write "              <TR>"
                Response.Write "                <TD><INPUT NAME=cfid TYPE=radio VALUE=""" & oSubFolders.item(i).getAttribute("id") & """ " & sChecked &  " /></TD>"
                Response.Write "                <TD><IMG SRC=""../images/folder.gif"" /></TD>"
                Response.Write "                <TD NOWRAP=""1""><A HREF=""channels_wizard.asp?" & CreateWizardRequest(aWizardInfo) & """><B><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#000000"">" & Server.HTMLEncode(oSubFolders.item(i).getAttribute("n")) & "</FONT></B></A></TD>"
                Response.Write "                <TD>&nbsp;&nbsp;</TD>"
                Response.Write "                <TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#000000"">" & Server.HTMLEncode(oSubFolders.item(i).getAttribute("mdt")) & "</FONT></TD>"
                Response.Write "                <TD>&nbsp;&nbsp;</TD>"
                Response.Write "                <TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#000000"">" & Server.HTMLEncode(oSubFolders.item(i).getAttribute("des")) & "</FONT></TD>"
                Response.Write "              </TR>"

                Response.Write "              <TR>"
                Response.Write "                <TD COLSPAN=""7"" ALIGN=""CENTER"" HEIGHT=""1"" BGCOLOR=""#6699CC""><img src=""../images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
                Response.Write "              </TR>"

            Next

        Else

            Response.Write "<TR>"
            Response.Write "  <TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""10"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD>"
            Response.Write "</TR>"

            Response.Write "<TR>"
            Response.Write " <TD ALIGN=""CENTER"" COLSPAN=13 BGCOLOR=""#ffffff"">"
            Response.Write "    <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"
            Response.Write "      <B>" & asDescriptors(837) & "</B>" 'There are no valid objects in this folder
            Response.Write "    </FONT></TD>"
            Response.Write "  </TD>"
            Response.Write "</TR>"

            Response.Write "<TR>"
            Response.Write "  <TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""10"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD>"
            Response.Write "</TR>"

            If Not oParent Is Nothing Then
                aWizardInfo(CHANNEL_FOLDER_ID) = ""
                aWizardInfo(CHANNEL_PARENT_ID) = oParent.getAttribute("id")

                Response.Write "<TR>"
                Response.Write "  <TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""5"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD>"
                Response.Write "</TR>"

                Response.Write "<TR>"
                Response.Write "<TD COLSPAN=13 BGCOLOR=""#ffffff"">"
                Response.Write "    <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"
                Response.Write "     <B><A HREF=""channels_wizard.asp?" & CreateWizardRequest(aWizardInfo) & """>" & asDescriptors(147) & "</A></B>" 'Back to Parent Folder
                Response.Write "</FONT></TD>"
                Response.Write "</TD>"
                Response.Write "</TR>"
            End If

        End If

        'Restore the original FolderId
        aWizardInfo(CHANNEL_FOLDER_ID) = sFolderId

        Response.Write "            </TABLE>"
        Response.Write "          </TD>"
        Response.Write "        <TR>"
        Response.Write "          <TD HEIGHT=""30""><IMG SRC=""../images/1ptrans.gif""></TD>"
        Response.Write "        </TR>"
        Response.Write "        <TR>"
        Response.Write "          <TD WIDTH=100% ALIGN=""LEFT"" NOWRAP>"
        Response.Write "            <INPUT name=cancel type=SUBMIT class=""buttonClass"" value=""" & asDescriptors(334) & """></INPUT>&nbsp;"  'Descriptor:Back

        If oSubFolders.length > 0 Then
            Response.Write "                  <INPUT name=next    type=SUBMIT class=""buttonClass"" value=""" & asDescriptors(335) & """ ></INPUT>"  'Descriptor:Next
        End If

        Response.Write "          </TD>"
        Response.Write "        </TR>"
        Response.Write "      </TABLE>"
        Response.Write "      </FORM>"

    End If

    Set oFolderDOM = Nothing
    Set oFolderContent = Nothing
    Set oSubFolders = Nothing

End Function

Function RenderSelectName(aWizardInfo)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:  none
'********************************************************
Dim sName
Dim sDesc
Dim bActive
Dim sTitle

    On Error Resume Next

    'Set the values
    sName = Server.HTMLEncode(aWizardInfo(CHANNEL_NAME))
    sDesc = Server.HTMLEncode(aWizardInfo(CHANNEL_DESC))
    bActive = (aWizardInfo(CHANNEL_ACTIVE) = "1")

    If aWizardInfo(CHANNEL_NAME) = "" Then
        sName = Server.HTMLEncode(NewChannelName(aWizardInfo(CHANNEL_FOLDER_NAME)))
        sDesc = Server.HTMLEncode(aWizardInfo(CHANNEL_FOLDER_DESC))
    End If

    If aWizardInfo(CHANNEL_ACTIVE) = "" Then bActive = True

    If aWizardInfo(CHANNEL_ACTION) = "add" Then
        sTitle = asDescriptors(476) 'Descriptor: Create New Channel
    Else
        sTitle = asDescriptors(477) 'Descriptor: Edit Channel
    End If


    Response.Write "      <FORM Action=""channels_wizard.asp"" METHOD=""POST"" NAME=""FormChannelName"">"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""st"" VALUE=""2"">"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""action"" VALUE=""" & aWizardInfo(CHANNEL_ACTION) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cid"" VALUE=""" & aWizardInfo(CHANNEL_ID) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""cfid"" VALUE=""" & aWizardInfo(CHANNEL_FOLDER_ID) & """>"
    Response.Write "      <INPUT TYPE=""HIDDEN"" NAME=""fnm"" VALUE=""" & Server.HTMLEncode(aWizardInfo(CHANNEL_FOLDER_NAME)) & """>"

    Response.Write "      <TABLE BORDER=""0"" WIDTH=""100%"" CELLSPACING=0 CELLPADDING=0>"
    Response.Write "        <TR>"
    Response.Write "          <TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >" & asDescriptors(483) & "</FONT></TD>" 'Descriptor: Enter a name for this new channel:
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD>"
    Response.Write "            <INPUT CLASS=""textBoxClass""  SIZE=40 NAME=cnm VALUE=""" & sName & """></INPUT></TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD HEIGHT=""10""><IMG SRC=""../images/1ptrans.gif""></TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD>"
    Response.Write "            <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >" & asDescriptors(484) & "</FONT></TD>" 'Descriptor: Enter a description for this channel: (Optional)
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD WIDTH=""5""><TEXTAREA CLASS=""textBoxClass""  COLS=60 NAME=cds ROWS=5 >" & sDesc & "</TEXTAREA></TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD HEIGHT=""10""><IMG SRC=""../images/1ptrans.gif""></TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD><INPUT name=cac type=checkbox value=""1"" "
    If bActive Then Response.Write "checked "
    Response.Write ">"
    Response.Write "              <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >" & asDescriptors(146) & "</FONT>" 'Descriptor:  Channel Enabled (Channel is visible to end-users)
    Response.Write "            </INPUT>"
    Response.Write "            <INPUT name=cac type=hidden value=""0"" >"
    Response.Write "          </TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD HEIGHT=""15""><IMG SRC=""../images/1ptrans.gif""></TD>"
    Response.Write "        </TR>"
    Response.Write "        <TR>"
    Response.Write "          <TD>"
    Response.Write "            <TABLE WIDTH=""100%"" BORDER=0 CELLPADDING=0 CELLSPACING=0 >"
    Response.Write "              <TR>"
    Response.Write "                <TD WIDTH=""5"" ALIGN=""LEFT"">"
    Response.Write "                  <INPUT name=back  type=SUBMIT class=""buttonClass"" value=""" & asDescriptors(334) & """ ></INPUT>"  'Descriptor:Previous
    Response.Write "                </TD>"
    Response.Write "                <TD WIDTH=""5""><IMG SRC=""../images/1ptrans.gif"" WIDTH=""5"" /></TD>"
    Response.Write "                <TD ALIGN=""LEFT"">"
    Response.Write "                  <INPUT name=next  type=SUBMIT class=""buttonClass"" value=""" & asDescriptors(335) & """ onClick=""return validateForm();"" ></INPUT>"  'Descriptor:Next
    Response.Write "                </TD>"
    Response.Write "                <TD ALIGN=""RIGHT"">"
    'Response.Write "                  <INPUT name=cancel type=SUBMIT class=""buttonClass"" value=""" & asDescriptors(120)  & """></INPUT>"  'Descriptor:Cancel
    Response.Write "                </TD>"
    Response.Write "              </TR>"
    Response.Write "            </TABLE>"
    Response.Write "          </TD>"
    Response.Write "        </TR>"
    Response.Write "      </TABLE>"
    Response.Write "      </FORM>"


End Function

Function RenderFolderPath(aWizardInfo, sFolderXML)
'********************************************************
'*Purpose:  Shows the path of the selected folder
'*Inputs:   sFolderXML: with the information of a document
'*Outputs:  none
'********************************************************
Dim lErr
Dim oContentsDOM
Dim oFolder
Dim iNumFolders
Dim i
Dim sFolderId

    On Error Resume Next
    lErr = NO_ERR

    'Load XML into a DOM Object:
    Set oContentsDOM = Server.CreateObject("Microsoft.XMLDOM")
    oContentsDOM.async = False
    If oContentsDOM.loadXML(sFolderXML) = False Then
        lErr = ERR_XML_LOAD_FAILED
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ChannelsCuLib.asp", "RenderFolderPath", "", "Error loading folder xml file", LogLevelError)
    Else
        iNumFolders = CInt(oContentsDOM.selectNodes("//a").length)
        If Err.number <> NO_ERR Then
            lErr = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "ChannelsCuLib.asp", "RenderFolderPath", "", "Error retrieving oi nodes", LogLevelError)
        End If
    End If

    'Start by saying: you are here:
    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"
    Response.Write asDescriptors(26) & " " 'Descriptor: You are here:

    If lErr = NO_ERR Then

        'No folders? We must be on the root:
        If iNumFolders = 0 Then
            Response.Write "<b>" & "&gt;" & "</b>" 'Root

        Else
            'Get the path node:
            Set oFolder = oContentsDOM.selectSingleNode("/mi/as")

            'Keep the Folder Id
            sFolderId = aWizardInfo(CHANNEL_FOLDER_ID)

            'Go recursively:
            For i=1 To iNumFolders

                Set oFolder = oFolder.selectSingleNode("a")
                Response.Write " &gt; "

                If i = iNumFolders Then
                    Response.Write "<b>" & oFolder.selectSingleNode("fd").getAttribute("n") & "</b>"
                Else
                    aWizardInfo(CHANNEL_FOLDER_ID) = ""
                    aWizardInfo(CHANNEL_PARENT_ID) = oFolder.selectSingleNode("fd").getAttribute("id")
                    Response.Write "<A HREF=""channels_wizard.asp?" & CreateWizardRequest(aWizardInfo) & """><font color=""#0000"">" & Server.HTMLEncode(oFolder.selectSingleNode("fd").getAttribute("n")) & "</font></A>"
                End If
            Next

            'Restore the Folder Id
            aWizardInfo(CHANNEL_FOLDER_ID) = sFolderId

        End If
    Else
        'add handling
    End If

    Response.Write "</font>"

    Set oContentsDOM = Nothing
    Set oFolder = Nothing

    Err.Clear
End Function

Function RenderChannelProperties(aWizardInfo)
Dim bLine

    Response.Write "<TABLE BORDER=0 CELLSPACING=10 CELLPADDING=0>"

    If Len(aWizardInfo(CHANNEL_NAME)) > 0 Then
        Response.Write "<TR>"
        Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(306) & ":</B></font></TD>"'Descriptor: Name
        Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & aWizardInfo(CHANNEL_NAME) & "</font></TD>"
    End If

    If Len(aWizardInfo(CHANNEL_DESC)) > 0 Then
        Response.Write "<TR>"
        Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(22)  & ":</B></font></TD>"'Descriptor:Description
        Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & aWizardInfo(CHANNEL_DESC) & "</font></TD>"
    End If

    If Len(aWizardInfo(CHANNEL_ACTIVE)) > 0 Then
        Response.Write "<TR>"
        Response.Write "<TD><B><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & "Active: (ND)" & "</B></font></TD>" 'Descriptor: Active
        Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & aWizardInfo(CHANNEL_ACTIVE) & "</font></TD>"
    End If

    Response.Write "<TR>"
    Response.Write "<TD COLSPAN=2 BGCOLOR=""#000000"" HEIGHT=1><IMG SRC=""../images/1ptrans.gif"" /></TD>"
    Response.Write "</TR>"

    bLine = False

    If aWizardInfo(CHANNEL_FOLDER_NAME) <> "" Then
        Response.Write "<TR>"
        Response.Write "<TD><B><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(255) & "</B></font></TD>" 'Descriptor: Folder name:
        Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & aWizardInfo(CHANNEL_FOLDER_NAME) & "</font></TD>"
        bLine = True
    End If

    If bLine Then
        Response.Write "<TR>"
        Response.Write "<TD COLSPAN=2 BGCOLOR=""#000000"" HEIGHT=1><IMG SRC=""../images/1ptrans.gif"" /></TD>"
        Response.Write "</TR>"
    End If

    Response.Write "</TABLE>"

End Function



Function ModifyChannel(aWizardInfo)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
CONST PROCEDURE_NAME = "ModifyChannel"
Dim lErr
Dim sSiteId

    On Error Resume Next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    If lErr = NO_ERR Then

        Select Case aWizardInfo(CHANNEL_ACTION)
        Case "add"
            'If NO ID, generate one:
            If aWizardInfo(CHANNEL_ID) = "" Then aWizardInfo(CHANNEL_ID) = GetGUID()
            If aWizardInfo(CHANNEL_ACTIVE) = "" Then aWizardInfo(CHANNEL_ACTIVE) = "0"

            lErr = ValidateChannelName(aWizardInfo)
            If lErr <> 0 Then
                Call LogErrorXML(aConnectionInfo, lErr, "", Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "", "Error calling co_AddChannel", LogLevelTrace)
            Else
                lErr = co_AddChannel(sSiteId, aWizardInfo(CHANNEL_ID), GetChannelXML(aWizardInfo))
                If lErr <> 0 Then Call LogErrorXML(aConnectionInfo, lErr, "", Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "", "Error calling co_AddChannel", LogLevelTrace)
            End If

        Case "edit"
            If aWizardInfo(CHANNEL_ACTIVE) = "" Then aWizardInfo(CHANNEL_ACTIVE) = "0"

            lErr = ValidateChannelName(aWizardInfo)
            If lErr <> 0 Then
                Call LogErrorXML(aConnectionInfo, lErr, "", Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "", "Error calling co_AddChannel", LogLevelTrace)
            Else
                lErr = co_UpdateChannel(sSiteId, aWizardInfo(CHANNEL_ID), GetChannelXML(aWizardInfo))
                If lErr <> 0 Then Call LogErrorXML(aConnectionInfo, lErr, "", Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "", "Error calling co_UpdateChannel", LogLevelTrace)
            End If

        Case "delete":
            sReturn = co_DeleteChannel(sSiteId, aWizardInfo(CHANNEL_ID))
            If lErr <> 0 Then Call LogErrorXML(aConnectionInfo, lErr, "", Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "", "Error calling co_deleteChannel", LogLevelTrace)

        Case Else:
            lErr = URL_MISSING_PARAMETER
            Call LogErrorXML(aConnectionInfo, lErr, "", Err.source, "ChannelsCoLib.asp", PROCEDURE_NAME, "", "No valid action specified", LogLevelError)

        End Select

    End If

    If lErr = NO_ERR Then
        Call ResetApplicationVariables()
    End If

    ModifyChannel = lErr
    Err.Clear

End Function


Function ValidateChannelName(aWizardInfo)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
CONST PROCEDURE_NAME = "ValidateChannelName"
Dim lErr
Dim sSiteId
Dim aChannels

Dim sChannelsXML
Dim oChannelsDOM

Dim oChannels
Dim sId
Dim sName
Dim i

Dim bSameName

    On Error Resume Next
    lErr = NO_ERR

    If lErr = NO_ERR Then
        lErr = cu_GetChannels(sChannelsXML)
    End If

    If lErr = NO_ERR Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sChannelsXML, oChannelsDOM)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sObjectsXML", LogLevelTrace)
    End If

    If lErr = NO_ERR Then
        Set oChannels = oChannelsDOM.selectNodes("//oi")

        If oChannels.length > 0 Then
            Redim aChannels(oChannels.length - 1)
            bSameName = False

            For i = 0 To oChannels.length - 1
                sId = oChannels(i).getAttribute("id")
                sName = oChannels(i).getAttribute("n")

                If sId = aWizardInfo(CHANNEL_ID) Then
                    If StrComp(aWizardInfo(CHANNEL_NAME), sName, vbTextCompare) = 0 Then
                        bSameName = True
                        Exit For
                    End If
                End If

                aChannels(i) = sName
            Next

            If bSameName = False Then
                If aWizardInfo(CHANNEL_NAME) <> GetNewName(aChannels, aWizardInfo(CHANNEL_NAME)) Then
                    lErr = ERR_INVALID_NAME
                End If
            End If

        End If
    End If

    ValidateChannelName = lErr
    Err.Clear

End Function


Function NewChannelName(sDefaultName)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
CONST PROCEDURE_NAME = "NewChannelName"
Dim lErr
Dim sSiteId
Dim aChannels

Dim sChannelsXML
Dim oChannelsDOM

Dim oChannels
Dim i

    On Error Resume Next
    lErr = NO_ERR

    If lErr = NO_ERR Then
        lErr = cu_GetChannels(sChannelsXML)
    End If

    If lErr = NO_ERR Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sChannelsXML, oChannelsDOM)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sObjectsXML", LogLevelTrace)
    End If

    If lErr = NO_ERR Then
        Set oChannels = oChannelsDOM.selectNodes("//oi")

        If oChannels.length > 0 Then
            Redim aChannels(oChannels.length - 1)

            For i = 0 To oChannels.length - 1
                aChannels(i) = oChannels(i).getAttribute("n")
            Next

        End If
    End If

    NewChannelName = GetNewName(aChannels, sDefaultName)
    Err.Clear

End Function

%>
