<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!-- #include file="../CoreLib/AdminCoLib.asp" -->
<%
'=--Steps on the wizard
'Engine configuration:
Const SECTION_ENGINE_CONFIG = 1
Const STEP_WELCOME = "1."
Const STEP_SELECT_MD = "2."
Const STEP_SELECT_MD_DBALIAS = "2.1."
Const STEP_SUMMARY_MD = "3."

'Portal Management
Const SECTION_PORTAL_MANAGEMENT = 2
Const STEP_SELECT_SITE = "2."
Const STEP_SELECT_PORTAL = "2."
Const STEP_SITE_NAME = "2.1"
Const STEP_SITE_AUREP = "2.2"
Const STEP_SITE_AUREP_DBALIAS = "2.2.1"
Const STEP_SITE_SBREP = "2.3"
Const STEP_SITE_SBREP_DBALIAS = "2.3.1"
Const SECTION_PORTAL_SUMMARY = "3."


'Site Management
Const SECTION_SITE_MANAGEMENT = 3
Const STEP_CHANNELS = "2."
Const STEP_CHANNELS_NAME = "2.2."
Const STEP_CHANNELS_FOLDER = "2.1."
Const STEP_DEVICE_TYPES = "3."
Const STEP_DEVICE_TYPES_EDIT = "3.1."
Const STEP_DEVICE_TYPES_FOLDER = "3.2."
Const STEP_SITE_DEVICES = "4."
Const STEP_SELECT_DEVICE = "4.1."
Const STEP_IS = "5."
Const STEP_PREFERENCES = "6."
Const STEP_SUMMARY_SITE = "7."

'Services
Const SECTION_SERVICES = 4
Const STEP_SERVICES = "2."
Const STEP_SERVICES_OVERVIEW = "2.1."
Const STEP_SERVICES_SELECT = "2.1."
Const STEP_SERVICES_DEFAULT = "2.3."
Const STEP_SERVICES_STATIC = "2.2."
Const STEP_SERVICES_STATIC_SUBSET = "2.2.1"
Const STEP_SERVICES_STATIC_SELECT_QO = "2.2.1.1"
Const STEP_SERVICES_STATIC_SELECT_MAP = "2.2.1.2"
Const STEP_SERVICES_STATIC_MAP_TABLES = "2.2.1.2.1"
Const STEP_SERVICES_STATIC_MAP_COLUMNS = "2.2.1.2.2"

Const STEP_SERVICES_DYNAMIC = "2.3."
Const STEP_SERVICES_DYNAMIC_SUBSET = "2.3.1"
Const STEP_SERVICES_DYNAMIC_SELECT_QO = "2.3.1.1"
Const STEP_SERVICES_DYNAMIC_SELECT_MAP = "2.3.1.2"
Const STEP_SERVICES_DYNAMIC_MAP_TABLES = "2.3.1.2.1"
Const STEP_SERVICES_DYNAMIC_MAP_COLUMNS = "2.3.1.2.2"
Const STEP_SERVICES_DYNAMIC_TABLES = "2.3.2"
Const STEP_SERVICES_DYNAMIC_COLUMNS = "2.3.3"
Const STEP_SERVICES_SUMMARY = "3."

Const STEP_FINISH = "3."


Function DisplayAdminError(sErrorHeader, sErrorMessage, lErrorId, sButtonCaption, szParentPage)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************

	On Error Resume Next

    Select Case lErrorId
    Case ERR_FOLDER_NOT_FOUND
        Response.Write "<TABLE BORDER=0 CELLPADDING=8 CELLSPACING=0 WIDTH=""100%"">"
        Response.Write "<FORM METHOD=""GET"" ACTION=""" & szParentPage & """><TR>"
        Response.Write "<TD VALIGN=TOP WIDTH=""1%"">"
        Response.Write "<IMG SRC=""../images/jobError.gif"" WIDTH=""55"" HEIGHT=""65"" BORDER=""0"" ALT="""">"
        Response.Write "</TD>"
        Response.Write "<TD VALIGN=TOP WIDTH=""99%"">"
        Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """ COLOR=""#cc0000""><B>" & asDescriptors(876) & "</B></FONT><BR /> <BR/>" ' "Folder Not Found"
        Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """>" & asDescriptors(877) & "</FONT><BR />"  '"The folder requested was not found in the Object Repository."
        Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """>" & "<A HREF=""" & aPageInfo(S_NAME_PAGE) & "?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) & "&fid="">" & asDescriptors(878) & "</A><BR/></FONT>" '"Click here to browse the Project's root folder."
        Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & Replace(asDescriptors(879), "#", oRequest("fid")) & "</FONT><BR /></TD>" '"(Folder Id: #)"
        Response.Write "<TR><TD></TD><TD><A HREF=""ShowErrors.asp?src=" & aPageInfo(S_NAME_PAGE) & """><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(880) & "</font></A></TD></TR>" 'Show Error Log
        Response.Write "<TR><TD></TD><TD><BR /><input TYPE=""SUBMIT"" CLASS=""buttonClass"" VALUE=""" & sButtonCaption & """ id=1 name=1></TD></TR>"
        Response.Write "</FORM></TABLE>"

    Case Else:
        Response.Write "<TABLE BORDER=0 CELLPADDING=8 CELLSPACING=0 WIDTH=""100%"">"
        Response.Write "<FORM METHOD=""GET"" ACTION=""" & szParentPage & """><TR>"
        Response.Write "<TD VALIGN=TOP WIDTH=""1%"">"
        Response.Write "<IMG SRC=""../images/jobError.gif"" WIDTH=""55"" HEIGHT=""65"" BORDER=""0"" ALT="""">"
        Response.Write "</TD>"
        Response.Write "<TD VALIGN=TOP WIDTH=""99%"">"
        Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """><FONT COLOR=""#cc0000""><b>" & sErrorHeader & "</b></font><BR> <BR>"
        Response.Write sErrorMessage & "<BR>(Error code: "& lErrorId & ")</font></TD>"
        Response.Write "<TR><TD></TD><TD><A HREF=""ShowErrors.asp?src=" & aPageInfo(S_NAME_PAGE) & """><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(880) & "</font></A></TD></TR>" 'Show Error Log
        Response.Write "<TR><TD></TD><TD><BR /><input TYPE=""SUBMIT"" CLASS=""buttonClass"" VALUE=""" & sButtonCaption & """></TD></TR>"
        Response.Write "</FORM></TABLE>"
    End Select

	DisplayAdminError = Err.number
End Function

Function ResetSubscriptionEngine()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "ResetSubscriptionEngine"
	Dim lErrNumber

	lErrNumber = NO_ERR

	lErrNumber = co_ResetSubscriptionEngine()
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_ResetSubscriptionEngine", LogLevelTrace)
	End If

	ResetSubscriptionEngine = lErrNumber
	Err.Clear
End Function

Function ConnectToSubscriptionEngines()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "ConnectToSubscriptionEngines"
	Dim lErrNumber

	lErrNumber = NO_ERR

	lErrNumber = co_ConnectToSubscriptionEngines()
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_ConnectToSubscriptionEngine", LogLevelTrace)
	End If

	ConnectToSubscriptionEngine = lErrNumber
	Err.Clear
End Function


Function selectDefaultValue(aValueList, aCandidates)
'********************************************************
'*Purpose:  Returns the name of the value that best fit to
'           the list of candidates.
'           If no one fits, returns the first value in the List
'*Inputs:   aValueList: list of available values
'           aCandidates: List of possible names. The more
'             top in the order is the candidate, the most
'             relevance it gets.
'*Outputs:  Name of the most probable alias
'********************************************************
Dim lErr
Dim sBestValue

Dim lBestPosition
Dim lPos
Dim i
Dim j

    On Error Resume Next
    lErr = NO_ERR

    'Search on all the aValueList each name.
    'If we found a DBAlias that exactly start with that string
    'of one Candidate, We assume that is the lucky guy:
    'If not, we'll use the aValueList which contains one
    'of the Candidates closer.
    lBestPosition = 0
    For i = 0 To UBound(aValueList)
        If (Len(aValueList(i)) + 1) > lBestPosition Then lBestPosition = (Len(aValueList(i)) + 1)
    Next

    For i = 0 To UBound(aCandidates)

        If Len(aCandidates(i)) > 0 Then
            For j = 0 To UBound(aValueList)

                'If found exact match, end looking:
                If StrComp(aCandidates(i), aValueList(j), vbTextCompare) = 0 Then
                    sBestValue = aValueList(j)
                    lBestPosition = 1
                    Exit For
                End If

                'If the candidate is found, check if the Alias is closer than
                'the previous:
                lPos = InStr(1, aValueList(j), aCandidates(i), vbTextCompare)
                If lPos > 0 Then
                    If lPos < lBestPosition Then
                        lBestPosition = lPos
                        sBestValue = aValueList(j)
                    End If
                End If

            Next
        End If

        'If we're already found an Alias that start with
        'one of the candidates, don't check the others
        if lBestPosition = 1 Then Exit For

    Next


    'Check if we already have selected an alias, if not,
    'by default select the first on the list:
    If sBestValue = "" Then sBestValue = aValueList(0)

    selectDefaultValue = sBestValue
    Err.Clear

End Function

Function selectDefaultDBALias(aValueList, aCandidates)
'********************************************************
'*Purpose:  Returns the name of the value that best fit to
'           the list of candidates.
'           If no one fits, returns the first value in the List
'*Inputs:   aValueList: list of available values
'           aCandidates: List of possible names. The more
'             top in the order is the candidate, the most
'             relevance it gets.
'*Outputs:  Name of the most probable alias
'********************************************************
Dim lErr
Dim sBestValue

Dim lBestPosition
Dim lPos
Dim i
Dim j

    On Error Resume Next
    lErr = NO_ERR

    'Search on all the aValueList each name.
    'If we found a DBAlias that exactly start with that string
    'of one Candidate, We assume that is the lucky guy:
    'If not, we'll use the aValueList which contains one
    'of the Candidates closer.
    lBestPosition = 0
    For i = 0 To UBound(aValueList)
        If (Len(aValueList(i,0)) + 1) > lBestPosition Then lBestPosition = (Len(aValueList(i,0)) + 1)
    Next

    For i = 0 To UBound(aCandidates)

        If Len(aCandidates(i)) > 0 Then
            For j = 0 To UBound(aValueList)

                'If found exact match, end looking:
                If StrComp(aCandidates(i), aValueList(j,0), vbTextCompare) = 0 Then
                    sBestValue = aValueList(j,0)
                    lBestPosition = 1
                    Exit For
                End If

                'If the candidate is found, check if the Alias is closer than
                'the previous:
                lPos = InStr(1, aValueList(j,0), aCandidates(i), vbTextCompare)
                If lPos > 0 Then
                    If lPos < lBestPosition Then
                        lBestPosition = lPos
                        sBestValue = aValueList(j,0)
                    End If
                End If

            Next
        End If

        'If we're already found an Alias that start with
        'one of the candidates, don't check the others
        if lBestPosition = 1 Then Exit For

    Next


    'Check if we already have selected an alias, if not,
    'by default select the first on the list:
    If sBestValue = "" Then sBestValue = aValueList(0,0)

    selectDefaultDBALias = sBestValue
    Err.Clear

End Function

Function RenderDropDownList(sDropDownName, aItems, sSelectedElement, sOnChange)
'********************************************************
'*Purpose:  Create a dropdown list based on the array
'*Inputs:   sDropDownName: The name of the SELECT
'*          aItems: Items to be displayed
'*          sSElectedElement: The current selection:
'*Outputs:
'********************************************************
Dim i
Dim sSelected

    On Error Resume Next

    Call Response.Write("<SELECT NAME=""" & sDropDownName & """ class=""pullDownClass"" ")
    If sOnChange <> "" Then
        Call Response.Write(" onchange=""" & sOnChange & """ ")
    End If
    Call Response.Write(">")

    For i = 0 to UBound(aItems)
        If sSelectedElement = aItems(i, 0) Then
            sSelected = " SELECTED"
        Else
            sSelected = ""
        End If

        Call Response.Write("<OPTION" & sSelected & " VALUE=""" &  aItems(i, 0) & """>" & aItems(i, 1) & "</OPTION>")
    Next

    Call Response.Write("</SELECT>")

End Function


Function RenderAdminWizardStep(nStep, sTitle, nIndentation, sLink, bSelected, bEnabled)
'********************************************************
'*Purpose: Renders a step of a Wizard
'*Inputs:  nStep:  Step number
'          sTitle: Step title
'          sDescription:Step Description, exactly below the title
'          bSelected: If this is the current step
'          bEnabled: If this step is enabled:
'*Outputs: Sends to the output a row being the
'********************************************************
Dim sStartFont
Dim sEndFont

Dim sStartLink
Dim sEndLink

Dim sLineImage
Dim sLineColor

Dim i

    sStartLink = ""
    sEndLink = ""

    If bSelected Then
        sStartFont = "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFF80""><B>"
        sEndFont   = "</B></FONT>"

        sLineImage   = "../images/1pgrey.gif"
        sLineColor = " BGCOLOR=""#999999"" "

    ElseIf bEnabled Then
        sStartFont = "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">"
        sEndFont   = "</FONT>"

        sLineImage   = "../images/1ptrans.gif"

        If sLink <> "" Then
            sStartLink = "<A HREF=""" & sLink & """>"
            sEndLink = "</A>"
        End If

    Else
        sStartFont = "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#c2c2c2"">"
        sEndFont   = "</FONT>"

        sLineImage   = "../images/1ptrans.gif"

    End If

    Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"

    Response.Write "<TR>"

    If nIndentation > 0 Then
        For i = 0 to nIndentation - 1
            Response.Write "  <TD WIDTH=""5"" ROWSPAN=""6""><IMG SRC=""../images/1ptrans.gif"" WIDTH=""10"" HEIGHT=""1"" BORDER=""0"" ALT="""" /></TD>" & vbCrLf
        Next
        Response.Write "<TD COLSPAN=""5""><IMG SRC=""../images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" ALT="""" BORDER=""0"" /></TD>"
    Else
        Response.Write "<TD COLSPAN=""5""><IMG SRC=""../images/1ptrans.gif"" HEIGHT=""10"" WIDTH=""1"" ALT="""" BORDER=""0"" /></TD>"
    End If

    Response.Write "</TR>"

    Response.Write "<TR><TD COLSPAN=""5"" HEIGHT=""2"" ALIGN=""RIGHT""><IMG SRC=""" & sLineImage & """ WIDTH=""1"" HEIGHT=""2"" BORDER=""0"" ALT="""" /></TD></TR>" & vbCrLf
    Response.Write "<TR><TD COLSPAN=""5"" HEIGHT=""1"" " & sLineColor & "><IMG SRC=""../images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT="""" /></TD></TR>" & vbCrLf
    Response.Write "<TR>" & vbCrLf


    Response.Write "  <TD WIDTH=""1"" VALIGN=""TOP"" " & sLineColor & " ><IMG SRC=""../images/1ptrans.gif"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD>" & vbCrLf
    Response.Write "  <TD WIDTH=""5""><IMG SRC=""../images/1ptrans.gif"" WIDTH=""5"" HEIGHT=""1"" BORDER=""0"" ALT="""" /></TD>" & vbCrLf
    Response.Write "  <TD ALIGN=""LEFT"" VALIGN=""TOP"" WIDTH=""15"">" & sStartFont & nStep & sEndFont & "</TD>" & vbCrLf
    Response.Write "  <TD WIDTH=""5""><IMG SRC=""../images/1ptrans.gif"" WIDTH=""5"" HEIGHT=""1"" BORDER=""0"" ALT="""" /></TD>" & vbCrLf
    Response.Write "  <TD ALIGN=""LEFT"" VALIGN=""TOP"">" & sStartLink & sStartFont & sTitle & sEndFont & sEndLink & "</TD>" & vbCrLf
    Response.Write "</TR>" & vbCrLf
    Response.Write "<TR><TD COLSPAN=""5"" HEIGHT=""1"" " & sLineColor & "><IMG SRC=""../images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT="""" /></TD></TR>" & vbCrLf
    Response.Write "<TR><TD COLSPAN=""5"" HEIGHT=""2"" ALIGN=""RIGHT""><IMG SRC=""" & sLineImage & """ WIDTH=""1"" HEIGHT=""2"" BORDER=""0"" ALT="""" /></TD></TR>" & vbCrLf

    Response.Write "</TABLE>"



End Function


Function RenderAdminSection(sTitle, sLink, bSelected, bEnabled)
'********************************************************
'*Purpose: Renders a Section of the Admin Wizard
'*Inputs:  sTitle: Section title
'          sLink:  Link to the default page of the section
'          bSelected: If this is the current Section
'          bEnabled: If this Section is enabled.
'*Outputs:
'********************************************************
Dim sStartFont
Dim sEndFont

Dim sStartLink
Dim sEndLink

Dim sLineImage


    sStartLink = ""
    sEndLink = ""

    If bSelected Then
        sStartFont = "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFF80""><B>"
        sEndFont   = "</B></FONT>"

    ElseIf bEnabled Then
        sStartFont = "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#FFFFFF"">"
        sEndFont   = "</FONT>"

        If sLink <> "" Then
            sStartLink = "<A HREF=""" & sLink & """>"
            sEndLink = "</A>"
        End If

    Else
        sStartFont = "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#c2c2c2"">"
        sEndFont   = "</FONT>"

    End If

    Response.Write "<TD>" & sStartLink & sStartFont & sTitle & sEndFont & sEndLink & "</TD>" & vbCrLf
    Response.Write "<TD WIDTH=""15""><IMG SRC=""../images/1ptrans.gif"" WIDTH=""15"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"

End Function

Function RenderFolderPath(oFolderDOM, sLink, sFolderToken)
'********************************************************
'*Purpose:  Shows the path of the selected folder
'*Inputs:   oFolderDOM: DOM Object for the folder.
'*          sLink: Target for links to the folders in the path:
'*Outputs:  none
'********************************************************
Dim oFolder
Dim iNumFolders
Dim i

    On Error Resume Next
    lErr = NO_ERR

    'Start by saying: you are here:
    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"
    Response.Write asDescriptors(26) & " " 'Descriptor: You are here:

    'Get the number of folders:
    iNumFolders = CInt(oFolderDOM.selectNodes("//a").length)

    'No folders? We must be on the root:
    If iNumFolders = 0 Then
        Response.Write "<b>" & "&gt;" & "</b>" 'Root

    Else
        'Get the path node:
        Set oFolder = oFolderDOM.selectSingleNode("/mi/as")

        'Go recursively:
        For i = 1 To iNumFolders

            Set oFolder = oFolder.selectSingleNode("a")
            Response.Write " &gt; "

            If i = iNumFolders Then
                Response.Write "<b>" & oFolder.selectSingleNode("fd").getAttribute("n") & "</b>"
            Else
                Response.Write "<A HREF=""" & sLink & "&" & sFolderToken & "=" & oFolder.selectSingleNode("fd").getAttribute("id") & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & oFolder.selectSingleNode("fd").getAttribute("n") & "</b></font></A>"
            End If

        Next

    End If

    Response.Write "</font>"

    Set oFolder = Nothing
    Err.Clear

End Function


Function GetOverviewSettings(aPageInfo, lAdminSection, sRedirectPage, sPreviousPage, sErrorMessage)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "getOverviewSettings"
	Dim lErr
	Dim sLink

	lErr = NO_ERR

	'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_NAME_PAGE) = "adminOverview.asp"

    If IsEmpty(lAdminSection) Or (lAdminSection = 0) Then
		lAdminSection = SECTION_ENGINE_CONFIG
    End If

    aPageInfo(S_TITLE_PAGE) = STEP_WELCOME & " " & asDescriptors(709) 'Descriptor: Overview

    'Based upon DB type, set variables
    Select Case lAdminSection
		Case SECTION_ENGINE_CONFIG:
				sRedirectPage = "select_engine.asp"
				aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_ENGINE_CONFIG
		Case SECTION_PORTAL_MANAGEMENT:
				sRedirectPage = "select_site.asp"
				aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT
				aPageInfo(S_TITLE_PAGE) = STEP_WELCOME & " " & asDescriptors(815) 'Descriptor: Edit Portal Configuration Overview
		Case SECTION_SITE_MANAGEMENT:
				sRedirectPage = "channels.asp"
				aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SITE_MANAGEMENT
				aPageInfo(S_TITLE_PAGE) = STEP_WELCOME & " " & asDescriptors(819) 'Descriptor: Site Preferences Overview
		Case SECTION_SERVICES:
				sRedirectPage = "services_overview.asp?next=Next"
				aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    End Select


    If lAdminSection = SECTION_ENGINE_CONFIG Then
		sPreviousPage = "Welcome.asp"
    Else
		 sPreviousPage = "AdminSummary.asp?section=" & (lAdminSection-1)
    End If

	getOverviewSettings = lErr
End Function


Function DisplayNextButtonSummary(lStatus, lAdminSection)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "displayNextButtonSummary"
	Dim bContinue

	bContinue = False

	'Based upon DB type, set variables

	If (lStatus = CONFIG_OK) Then
		bContinue = True
	Else
		Select Case lAdminSection
			Case SECTION_ENGINE_CONFIG:
					If ((lStatus And CONFIG_MISSING_MD) = 0)Then
						bContinue = True
					End If
			Case SECTION_PORTAL_MANAGEMENT:
					If ((lStatus And CONFIG_MISSING_AUREP) = 0) And ((lStatus And CONFIG_MISSING_SBREP) = 0) Then
						bContinue = True
					End If
			Case SECTION_SITE_MANAGEMENT:
					If (lStatus And CONFIG_MISSING_SITE) = 0 Then
						bContinue = True
					End If
		End Select
	End If

	displayNextButtonSummary1 = bContinue

End Function


Function GetSummarySettings(aPageInfo, lAdminSection, sRedirectPage, sPreviousPage, sErrorMessage)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "getSummarySettings"
	Dim lErr
	Dim sLink

	lErr = NO_ERR

	'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_NAME_PAGE) = "adminSummary.asp"

    If IsEmpty(lAdminSection) Or (lAdminSection = 0) Then
		lAdminSection = SECTION_ENGINE_CONFIG
    End If

    'Based upon DB type, set variables
    Select Case lAdminSection
		Case SECTION_ENGINE_CONFIG:
				sPreviousPage = "select_md.asp"
				aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_ENGINE_CONFIG
				aPageInfo(S_TITLE_PAGE) = STEP_SUMMARY_MD & " " & asDescriptors(331) 'Descriptor:Summary
		Case SECTION_PORTAL_MANAGEMENT:
				sPreviousPage = "select_site.asp"
				aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT
				aPageInfo(S_TITLE_PAGE) = SECTION_PORTAL_SUMMARY & " " & asDescriptors(331) 'Descriptor:Summary
		Case SECTION_SITE_MANAGEMENT:
				sPreviousPage = "preferences.asp"
				aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SITE_MANAGEMENT
				aPageInfo(S_TITLE_PAGE) = STEP_SUMMARY_SITE & " " & asDescriptors(331) 'Descriptor:Summary
		Case SECTION_SERVICES:
				sPreviousPage = "services_config.asp"
				aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
				aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_SUMMARY & " " & asDescriptors(331) 'Descriptor:Summary
    End Select

    If lAdminSection = SECTION_SERVICES Then
		sRedirectPage = "login.asp"
    Else
		 sRedirectPage = "adminOverview.asp?section=" & (lAdminSection + 1)
    End If

	getOverviewSettings = lErr
End Function


Function RenderSitePreferencesToolbar()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
		On Error Resume Next
		Dim sSiteLink

		sSiteLink = "adminOverview.asp?section=3"
		Call RenderAdminWizardStep(STEP_WELCOME, asDescriptors(709), 0, sSiteLink, (aPageInfo(S_NAME_PAGE) & "?section=3") = sSiteLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_MD) = 0)) 'Descriptor: Overview

        sSiteLink = "channels.asp"
        Call RenderAdminWizardStep(STEP_CHANNELS, asDescriptors(471), 0, sSiteLink, aPageInfo(S_NAME_PAGE) = sSiteLink, (lStatus = CONFIG_OK)) 'Descriptor:Channels

        If aPageInfo(S_NAME_PAGE) = "channels_wizard.asp" Then
            sSiteLink = "channels_wizard.asp?back=true&" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
            Call RenderAdminWizardStep(STEP_CHANNELS_FOLDER, asDescriptors(589), 1, sSiteLink, aWizardInfo(CHANNEL_STEP) = STEP_WIZARD_FOLDER, (lStatus = CONFIG_OK))  'Descriptor: Select the Channels Folder

            sSiteLink = "channels_wizard.asp?next=true&" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
            Call RenderAdminWizardStep(STEP_CHANNELS_NAME, asDescriptors(669), 1, sSiteLink, aWizardInfo(CHANNEL_STEP) = STEP_WIZARD_NAME, (lStatus = CONFIG_OK)) 'Descriptor:Name and description
        End If

        sSiteLink = "deviceTypes.asp"
        Call RenderAdminWizardStep(STEP_DEVICE_TYPES, asDescriptors(493), 0, sSiteLink, aPageInfo(S_NAME_PAGE) = sSiteLink, (lStatus = CONFIG_OK)) 'Descriptor"Device Types"

        If (aPageInfo(S_NAME_PAGE) = "editDeviceType.asp") Or _
           (aPageInfo(S_NAME_PAGE) = "deviceTypeFolders.asp") Then

            sSiteLink = "editDeviceType.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
            Call RenderAdminWizardStep(STEP_DEVICE_TYPES_EDIT, asDescriptors(516), 1, sSiteLink, aPageInfo(S_NAME_PAGE) = "editDeviceType.asp" , True)  'Descriptor: Definition

            sSiteLink = "deviceTypeFolders.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
            Call RenderAdminWizardStep(STEP_DEVICE_TYPES_FOLDER, asDescriptors(496), 1, sSiteLink, aPageInfo(S_NAME_PAGE) = "deviceTypeFolders.asp" , aDeviceTypeInfo(DEV_TYPE_ID) <> "")  'Descriptor: Device Folders

        End If


        sSiteLink = "devices_config.asp"
        Call RenderAdminWizardStep(STEP_SITE_DEVICES, asDescriptors(579), 0, sSiteLink, InStr(1, "devices_config.asp;select_devices.asp", aPageInfo(S_NAME_PAGE), vbTextCompare) > 0, (lStatus = CONFIG_OK)) 'Descriptor: Site Devices

        sSiteLink = "is_config.asp"
        Call RenderAdminWizardStep(STEP_IS, asDescriptors(568), 0,  sSiteLink, aPageInfo(S_NAME_PAGE) = sSiteLink, (lStatus = CONFIG_OK)) 'Descriptor: Information Sources

        sSiteLink = "preferences.asp"
        Call RenderAdminWizardStep(STEP_PREFERENCES, asDescriptors(286), 0,  sSiteLink, aPageInfo(S_NAME_PAGE) = sSiteLink, (lStatus = CONFIG_OK)) 'Descriptor:"Preferences"

		sSiteLink = "adminSummary.asp?section=3"
        Call RenderAdminWizardStep(STEP_SUMMARY_SITE, asDescriptors(331), 0, sSiteLink, (aPageInfo(S_NAME_PAGE)& "?section=3") = sSiteLink, ((lStatus = CONFIG_OK) Or (lStatus < CONFIG_MISSING_SBREP)))  'Need Descriptor: Summary

End Function


Function RenderEngineConfigToolbar()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	 On Error Resume Next
	 Dim sLink

	 sLink = "welcome.asp"
     Call RenderAdminWizardStep(STEP_WELCOME, asDescriptors(283), 0, sLink, aPageInfo(S_NAME_PAGE) = sLink, True)'Descriptor:Welcome

     sLink = "select_md.asp"
     Call RenderAdminWizardStep(STEP_SELECT_MD, asDescriptors(569), 0, sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_ENGINE) = 0))  'Descriptor:Metadata Connection

     If StrComp(aPageInfo(S_NAME_PAGE), "addDBAlias.asp", vbTextCompare) = 0 Then
		sLink = "addDBAlias.asp"
		Call RenderAdminWizardStep(STEP_SELECT_MD_DBALIAS, asDescriptors(621), 1, sLink, Strcomp(aPageInfo(S_NAME_PAGE), "addDBAlias.asp", vbTextCompare) = 0 , True)  'Descriptor: Add a new Database Alias
 	 End If

     If StrComp(aPageInfo(S_NAME_PAGE), "editDBAlias.asp", vbTextCompare) = 0 Then
		sLink = "editDBAlias.asp"
		Call RenderAdminWizardStep(STEP_SELECT_MD_DBALIAS, asDescriptors(891), 1, sLink, Strcomp(aPageInfo(S_NAME_PAGE), "editDBAlias.asp", vbTextCompare) = 0 , True)  'Descriptor: Edit a Database connection
 	 End If

     If StrComp(aPageInfo(S_NAME_PAGE), "delete_dbalias.asp", vbTextCompare) = 0 Then
		sLink = "delete_dbalias.asp"
		Call RenderAdminWizardStep(STEP_SELECT_MD_DBALIAS, asDescriptors(830), 1, sLink, Strcomp(aPageInfo(S_NAME_PAGE), "delete_dbalias.asp", vbTextCompare) = 0 , True)  'Descriptor: Add a new Database Alias
 	 End If

 	 sLink = "adminSummary.asp"
     Call RenderAdminWizardStep(STEP_SUMMARY_MD, asDescriptors(331), 0, sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_MD) = 0))  'Need Descriptor: Summary

End Function



Function RenderPortalManagementToolbar()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	 On Error Resume Next
	 Dim sLink
	 Dim sPages

	 sLink = "adminOverview.asp?section=2"
     Call RenderAdminWizardStep(STEP_WELCOME, asDescriptors(709), 0, sLink, (aPageInfo(S_NAME_PAGE)& "?section=2") = sLink, True)'Descriptor:Overview

     sPages = "select_portal.asp;delete_portal.asp"
     If InStr(1, sPages, aPageInfo(S_NAME_PAGE),vbTextCompare) > 0 Then
        sLink = ""
        Call RenderAdminWizardStep(STEP_SELECT_SITE, asDescriptors(623), 0, sLink, True, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_MD) = 0)) 'Descriptor:Site Definition

        'sLink = "select_portal.asp"
        'Call RenderAdminWizardStep(SECTION_PORTAL, asDescriptors(628), 1, sLink, aPageInfo(S_NAME_PAGE) = sLink, True)'Descriptor:Portal Management
     Else

        sLink = "select_site.asp"
        Call RenderAdminWizardStep(STEP_SELECT_SITE, asDescriptors(623), 0, sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_MD) = 0)) 'Descriptor:Site Definition

        sPages = "site_name.asp;select_aurep.asp;select_sbrep.asp;addDBAlias.asp;dbaliases.asp;delete_dbalias.asp;editDBAlias.asp"
        If InStr(1, sPages, aPageInfo(S_NAME_PAGE),vbTextCompare) > 0 Then
           sLink = "site_name.asp"
           Call RenderAdminWizardStep(STEP_SITE_NAME, asDescriptors(482), 1 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_SITE) = 0))  'Descriptor:"Name & Description"

           sLink = "select_aurep.asp"
           Call RenderAdminWizardStep(STEP_SITE_AUREP, asDescriptors(575), 1 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_SITE) = 0)) 'Descriptor:Project Repository"

	       If StrComp(aPageInfo(S_NAME_PAGE), "adddbalias.asp", vbTextCompare) = 0 Then
		    	sLink = "adddbalias.asp"
		    	If aPageInfo(N_ALIAS_PAGE) = 1 Then
		    		Call RenderAdminWizardStep(STEP_SITE_AUREP_DBALIAS, asDescriptors(621), 2 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_SITE) = 0)) 'Descriptor: Add a new Database Alias
		    	End If
		    End If

	       If StrComp(aPageInfo(S_NAME_PAGE), "editdbalias.asp", vbTextCompare) = 0 Then
		    	sLink = "editdbalias.asp"
		    	If aPageInfo(N_ALIAS_PAGE) = 1 Then
		    		Call RenderAdminWizardStep(STEP_SITE_AUREP_DBALIAS, "Edit a database alias", 2 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_SITE) = 0)) 'Descriptor: Add a new Database Alias
		    	End If
		    End If

           If StrComp(aPageInfo(S_NAME_PAGE), "delete_dbalias.asp", vbTextCompare) = 0 Then
		        sLink = "delete_dbalias.asp"
		    	If aPageInfo(N_ALIAS_PAGE) = 1 Then
		            Call RenderAdminWizardStep(STEP_SITE_AUREP_DBALIAS, asDescriptors(830), 1, sLink, Strcomp(aPageInfo(S_NAME_PAGE), "delete_dbalias.asp", vbTextCompare) = 0 , True)  'Descriptor: Delete a new Database Alias
		    	End If
 	       End If

           sLink = "select_sbrep.asp"
           Call RenderAdminWizardStep(STEP_SITE_SBREP, asDescriptors(581), 1 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_AUREP) = 0))  'Descriptor:"Subscription Book Repository"

           If StrComp(aPageInfo(S_NAME_PAGE), "addDBAlias.asp", vbTextCompare) = 0 Then
		    	sLink = "adddbalias.asp"
		    	If aPageInfo(N_ALIAS_PAGE) = 2 Then
		    		Call RenderAdminWizardStep(STEP_SITE_SBREP_DBALIAS, asDescriptors(621), 2 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_SITE) = 0)) 'Descriptor: Add a new Database Alias
		    	End If
		    End If

           If StrComp(aPageInfo(S_NAME_PAGE), "editDBAlias.asp", vbTextCompare) = 0 Then
		    	sLink = "editdbalias.asp"
		    	If aPageInfo(N_ALIAS_PAGE) = 2 Then
		    		Call RenderAdminWizardStep(STEP_SITE_SBREP_DBALIAS, asDescriptors(891), 2 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_SITE) = 0)) 'Descriptor: Edit a database connection
		    	End If
		    End If

           If StrComp(aPageInfo(S_NAME_PAGE), "delete_dbalias.asp", vbTextCompare) = 0 Then
		        sLink = "delete_dbalias.asp"
		    	If aPageInfo(N_ALIAS_PAGE) = 2 Then
		            Call RenderAdminWizardStep(STEP_SITE_SBREP_DBALIAS, asDescriptors(830), 1, sLink, Strcomp(aPageInfo(S_NAME_PAGE), "delete_dbalias.asp", vbTextCompare) = 0 , True)  'Descriptor: Delete a new Database Alias
		    	End If
 	       End If

        End If
    End If

    sLink = "adminSummary.asp?section=2"
    Call RenderAdminWizardStep(SECTION_PORTAL_SUMMARY, asDescriptors(331), 0, sLink, (aPageInfo(S_NAME_PAGE)& "?section=2") = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_SITE) = 0))  'Need Descriptor: Summary

End Function


Function proccessSiteManagement()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	 On Error Resume Next
	 Dim sLink

	 sLink = "select_site.asp"
     Call RenderAdminWizardStep(1, "Site Definition", 0, sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_MD) = 0))

     If (aPageInfo(S_NAME_PAGE) = "site_name.asp") Or _
        (aPageInfo(S_NAME_PAGE) = "select_aurep.asp") Or _
        (aPageInfo(S_NAME_PAGE) = "select_sbrep.asp") Or _
        (aPageInfo(S_NAME_PAGE) = "dbaliases.asp") Then
        sLink = "site_name.asp"
        Call RenderAdminWizardStep(STEP_SITE_NAME, asDescriptors(482), 1 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_SITE) = 0))  'Descriptor:"Name & Description"

        sLink = "select_aurep.asp"
        Call RenderAdminWizardStep(STEP_SITE_AUREP, asDescriptors(575), 1 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_SITE) = 0)) 'Descriptor:Project Repository"

        sLink = "select_sbrep.asp"
        Call RenderAdminWizardStep(STEP_SITE_SBREP, asDescriptors(581), 1 ,  sLink, aPageInfo(S_NAME_PAGE) = sLink, ((lStatus = CONFIG_OK) Or (lStatus AND CONFIG_MISSING_AUREP) = 0))  'Descriptor:"Subscription Book Repository"
     End If

End Function


Function RenderServicesConfigToolbar()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim sServicesLink
	Dim sServicesPages

    sServicesPages = "adminOverview.asp;finish.asp;services_overview.asp"
    If InStr(1, sServicesPages , aPageInfo(S_NAME_PAGE)) > 0 Then

	    sServicesLink = "adminOverview.asp?section=4"
        Call RenderAdminWizardStep(STEP_WELCOME, asDescriptors(709), 0, sServicesLink, (aPageInfo(S_NAME_PAGE)& "?section=4") = sServicesLink, True)'Descriptor:Overview

	    sServicesLink = "services_overview.asp"
        Call RenderAdminWizardStep(STEP_SERVICES, asDescriptors(362), 0, sServicesLink & "?next=Next", aPageInfo(S_NAME_PAGE) = sServicesLink, (lStatus = CONFIG_OK)) 'Descriptor"Services"

        sServicesLink = "finish.asp"
        Call RenderAdminWizardStep(STEP_FINISH, asDescriptors(331), 0,  sServicesLink, aPageInfo(S_NAME_PAGE) = sServicesLink, (lStatus = CONFIG_OK)) 'Descriptor:"Summary"

    Else

        If aSvcConfigInfo(SVCCFG_STEP) = "static" And aPageInfo(S_NAME_PAGE) <> "services_static.asp" Then

            sServicesPages = "services_map_tables.asp;services_map_columns.asp;adddbalias.asp"
            If InStr(1, sServicesPages , aPageInfo(S_NAME_PAGE), vbTextCompare) > 0 Then

                Call RenderAdminWizardStep("...", "", 0, "", False, False)

                sServicesLink = "services_select_map.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep(STEP_SERVICES_STATIC_SELECT_MAP, asDescriptors(782), 0, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_select_map.asp", Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0)  'Descriptor:Select Storage

                sServicesLink = "services_map_tables.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep("",  asDescriptors(783), 1, sServicesLink, (aPageInfo(S_NAME_PAGE) = "services_map_tables.asp") Or (aPageInfo(S_NAME_PAGE) = "adddbalias.asp"), True)  'Descriptor:Select Tables

                sServicesLink = "services_map_columns.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep("", asDescriptors(784), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_map_columns.asp", True)  'Descriptor:Select Columns

                Call RenderAdminWizardStep("...", "", 0, "", False, False)

            Else

                Call RenderAdminWizardStep("...", "", 0, "", False, False)

                sServicesLink = "services_subsset_modify.asp?back=true&" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep(STEP_SERVICES_STATIC, asDescriptors(722), 0, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_static.asp", Len(aSvcConfigInfo(SVCCFG_SVC_ID)) > 0)  'Static Subscriptions

                sServicesLink = "services_subsset.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep(STEP_SERVICES_STATIC_SUBSET, asDescriptors(785), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_subsset.asp", True)  'Descriptors"Configure Subscription Set"

                sServicesPages = "services_select_qo.asp;services_select_map.asp;services_map_tables.asp;services_map_columns.asp"
                If InStr(1, sServicesPages , aPageInfo(S_NAME_PAGE)) > 0 Then

                    sServicesLink = "services_select_qo.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                    Call RenderAdminWizardStep(STEP_SERVICES_STATIC_SELECT_QO, asDescriptors(786), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_select_qo.asp", True)   '"Select Question"

                    sServicesLink = "services_select_map.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                    Call RenderAdminWizardStep(STEP_SERVICES_STATIC_SELECT_MAP, asDescriptors(782), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_select_map.asp", Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0)  ' "Select Storage"

                End If

                Call RenderAdminWizardStep("...", "", 0, "", False, False)
            End If


        ElseIf aSvcConfigInfo(SVCCFG_STEP) = "dynamic" And aPageInfo(S_NAME_PAGE) <> "services_dynamic.asp" Then

            sServicesPages = "services_map_tables.asp;services_map_columns.asp;adddbalias.asp"
            If InStr(1, sServicesPages , aPageInfo(S_NAME_PAGE)) > 0 And Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then

                Call RenderAdminWizardStep("...", "", 0, "", False, False)

                sServicesLink = "services_select_map.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep(STEP_SERVICES_DYNAMIC_SELECT_MAP, asDescriptors(782), 0, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_select_map.asp", Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0)  '"Select Storage"

                sServicesLink = "services_map_tables.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep("", asDescriptors(783), 1, sServicesLink, (aPageInfo(S_NAME_PAGE) = "services_map_tables.asp") Or (aPageInfo(S_NAME_PAGE) = "adddbalias.asp"), True)  '"Select Tables"

                sServicesLink = "services_map_columns.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep("", asDescriptors(784), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_map_columns.asp", True)  '"Select Columns"

                Call RenderAdminWizardStep("...", "", 0, "", False, False)

            Else

                Call RenderAdminWizardStep("...", "", 0, "", False, False)

                sServicesLink = "services_subsset_modify.asp?back=true&" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep(STEP_SERVICES_DYNAMIC, asDescriptors(723), 0, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_static.asp", Len(aSvcConfigInfo(SVCCFG_SVC_ID)) > 0)  '"Dynamic Subscriptions"

                sServicesLink = "services_subsset.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep(STEP_SERVICES_DYNAMIC_SUBSET, asDescriptors(785), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_subsset.asp", True)  ' "Configure Subscription Set"

                sServicesPages = "services_select_qo.asp;services_select_map.asp"
                If InStr(1, sServicesPages , aPageInfo(S_NAME_PAGE)) > 0 Then

                    sServicesLink = "services_select_qo.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                    Call RenderAdminWizardStep(STEP_SERVICES_DYNAMIC_SELECT_QO, asDescriptors(786), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_select_qo.asp", True)  '"Select Question"

                    sServicesLink = "services_select_map.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                    Call RenderAdminWizardStep(STEP_SERVICES_DYNAMIC_SELECT_MAP, asDescriptors(782), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_select_map.asp", Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0)  '"Select Storage"

                End If

                sServicesLink = "services_map_tables.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep(STEP_SERVICES_DYNAMIC_TABLES, asDescriptors(783), 1, sServicesLink, (aPageInfo(S_NAME_PAGE) = "services_map_tables.asp") Or (aPageInfo(S_NAME_PAGE) = "adddbalias.asp"), Len(aSvcConfigInfo(SVCCFG_QO_ID)) = 0)  '"Select Tables"

                sServicesLink = "services_map_columns.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
                Call RenderAdminWizardStep(STEP_SERVICES_DYNAMIC_COLUMNS, asDescriptors(784), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_map_columns.asp", False)  '"Select Columns"


                Call RenderAdminWizardStep("...", "", 0, "", False, False)
            End If

        Else

            sServicesLink = "adminOverview.asp?section=4"
            Call RenderAdminWizardStep(STEP_WELCOME, asDescriptors(709), 0, sServicesLink, (aPageInfo(S_NAME_PAGE)& "?section=4") = sServicesLink, True)'Descriptor:Overview

            sServicesLink = "services_overview.asp"
            Call RenderAdminWizardStep(STEP_SERVICES, asDescriptors(362), 0, sServicesLink & "?next=Next", aPageInfo(S_NAME_PAGE) = sServicesLink, True) 'Descriptor"Services"

            sServicesLink = "services_select.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
            Call RenderAdminWizardStep(STEP_SERVICES_SELECT, asDescriptors(781), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_select.asp", Len(aSvcConfigInfo(SVCCFG_SVC_ID)) > 0)  'Descriptor: Select Service

            sServicesLink = "services_static.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
            Call RenderAdminWizardStep(STEP_SERVICES_STATIC, asDescriptors(722), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_static.asp", Len(aSvcConfigInfo(SVCCFG_SVC_ID)) > 0)  '"Static Subscriptions"

            sServicesLink = "services_dynamic.asp?" & aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)
            Call RenderAdminWizardStep(STEP_SERVICES_DYNAMIC, asDescriptors(723), 1, sServicesLink, aPageInfo(S_NAME_PAGE) = "services_dynamic.asp", Len(aSvcConfigInfo(SVCCFG_SVC_ID)) > 0)  '"Dynamic Subscriptions"

            sServicesLink = "finish.asp"
            Call RenderAdminWizardStep(STEP_FINISH, asDescriptors(331), 0,  sServicesLink, aPageInfo(S_NAME_PAGE) = sServicesLink, (lStatus = CONFIG_OK)) 'Descriptor:"Summary"

        End If
    End If


End Function

Function GetNewName(aNames, sDefaultName)
'********************************************************
'*Purpose: Returns the next valid New Name.
'           If sDefaultName already exist in the list of names, it returns
'           sDefaultName (n), where n is the next new name.
'*Inputs:  aName: list of current names, sDefaultName: the default name
'*Outputs: It returns the new Name
'********************************************************
Dim sNewName
Dim lCount
Dim bFound
Dim n, i

    If IsEmpty(aNames) Then
        sNewName = sDefaultName
    Else
        lCount = UBound(aNames)

        n = 1
        bFound = True
        sNewName = sDefaultName

        Do While bFound

            bFound = False

            For i = 0 to lCount
                If StrComp(aNames(i), sNewName, vbTextCompare) = 0 Then
                    bFound = True
                    Exit For
                End If
            Next

            If bFound Then
                sNewName = sDefaultName & " (" & n & ")"
                n = n + 1
            End If
        Loop

    End If

    GetNewName = sNewName
    Err.Clear

End Function


Function IsEngineFolderShared(sMachineName)
	On Error Resume Next
	Dim lErr
	Dim oShareEngineUtil
	Set oShareEngineUtil = Server.CreateObject(PROGID_NETSHARE)
	If Err.number <> NO_ERR Then
            lErr = Err.number
            Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "AdminCuLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_NETSHARE, LogLevelError)
	Else
	    lErr = oShareEngineUtil.IsSubEngineDriveShared(sMachineName)
	End if
	Set oShareEngineUtil = Nothing
	IsEngineFolderShared = lErr
End Function

Function ShareSubEngineFolder(sMachineName)
	On Error Resume Next
	Dim lErr
	Dim oShareEngineUtil
	Set oShareEngineUtil = Server.CreateObject(PROGID_NETSHARE)
	If Err.number <> NO_ERR Then
            lErr = Err.number
            Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "AdminCuLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_NETSHARE, LogLevelError)
    	Else
	    lErr = oShareEngineUtil.ShareSubEngineDrive(sMachineName)
	End If
	Set oShareEngineUtil = Nothing
	ShareSubEngineFolder = lErr
End Function

Function UnshareSubEngineFolder(sMachineName)
	On Error Resume Next
	Dim lErr
	Dim oShareEngineUtil
	Set oShareEngineUtil = Server.CreateObject(PROGID_NETSHARE)
	If Err.number <> NO_ERR Then
            lErr = Err.number
            Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "AdminCuLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_NETSHARE, LogLevelError)
    	Else
	    lErr = oShareEngineUtil.UnshareSubEngineDrive(sMachineName)
	End If
	Set oShareEngineUtil = Nothing
	ShareSubEngineFolder = lErr
End Function

Function GetEngineInstallDrive(sMachineName)
	On Error Resume Next
	Dim lErr
	Dim oShareEngineUtil
	Set oShareEngineUtil = Server.CreateObject(PROGID_NETSHARE)
	If Err.number <> NO_ERR Then
            lErr = Err.number
            Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "AdminCuLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_NETSHARE, LogLevelError)
    	Else
	    lErr = oShareEngineUtil.GetEngineInstallDrive(sMachineName)
	End If
	Set oShareEngineUtil = Nothing
	GetEngineInstallDrive = lErr
End Function

Function FillEngineAndDriveInfo(sMessage, sEngineName, sDriveName)
	sMessage = Replace(sMessage, "#1", sDriveName)
	sMessage = Replace(sMessage, "#2", sEngineName)
	FillEngineAndDriveInfo = sMessage
End Function
%>