<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!-- #include file="CustomLib/CommonLib.asp" -->
<!-- #include file="CustomLib/ConnectCuLib.asp"-->
<!-- #include file="CustomLib/commonHeaderCuLib.asp" -->

<%
    'Descriptors variables:
	Dim asDescriptors
	Dim aFontInfo()
	Redim aFontInfo(MAX_FONT_INFO)
	Dim aSourceInfo(2)

	Dim aPageInfo(12)  'MAX_PAGE_INFO



	aPageInfo(S_NAME_PAGE) = GetPageWithoutPath(Request.ServerVariables("PATH_INFO"))

	Dim bPromptPage
	bPromptPage = false
	If StrComp(aPageInfo(S_NAME_PAGE),"Prompt.asp",vbTextCompare) = 0 Then
		bPromptPage = True
	End If

	If Not bPromptPage Then
		Call SetLocaleInformation(asDescriptors, aFontInfo)
	Else
		Dim asWebDescriptors
		Dim asHydraDescriptors
		Call SetLocaleInformation(asHydraDescriptors, aFontInfo)
		Call SetLocaleInformationForPromptPage(asWebDescriptors, aFontInfo)
		asDescriptors = asHydraDescriptors
	End If



	'Error variables
	Dim lErr
	Dim lValidationError
	Dim sErrDesc
	Dim sErrorHeader
	Dim sErrorMessage
	sErrorHeader = asDescriptors(39) 'Descriptor: Error
	sErrorMessage = asDescriptors(375) 'Descriptor: An error has occurred on this page.
	lValidationError = NO_ERR

	'Connection Information
	Dim aConnectionInfo(13)  'MAX_CONNECTION_INFO
	Call SetConnectionInfo(oRequest, aConnectionInfo)

	'Request object:
	Dim oRequest
	If Request.Form.Count > 0 Then
		Set oRequest = Request.Form
	Else
		Set oRequest = Request.QueryString
	End If


    'CurrentChannel:
	Dim sChannel
	sChannel = Trim(CStr(oRequest("site")))
	If Len(sChannel) = 0 Then
		sChannel = GetCurrentChannel()
	Else
		aSourceInfo(0) = SITE_COOKIE
		aSourceInfo(1) = CURRENT_SITE
		aSourceInfo(2) = COOKIES_EXPIRATION_DATE
		Call WriteToSource(aConnectionInfo, CStr(sChannel), SOURCE_COOKIES, aSourceInfo)
		Call SetRootFolder(sChannel)
	End If


	'Global variables:
	Dim sSysAdminEmail
	Dim sSysAdminPhone
	Dim APP_ROOT_FOLDER
	Dim APP_CACHE_FOLDER
	Dim SITE_ID
	Dim ASP_VERSION
	lErr =  SetApplicationlVariables()
	SITE_ID            = Application.Value("SITE_ID")
	APP_CACHE_FOLDER   = Application("Cache_folder")
	APP_ROOT_FOLDER    = Application("Root_Folder")
	sSysAdminEmail     = Application("Admin_email")
	sSysAdminPhone     = Application("Admin_phone")
	'Eventually, this will not necessarily be the same as the SDK
	ASP_VERSION = GetSDKVersion()

	'General variables
	Dim asReservedChars
	Dim STYLE_BEIGE_BACKGROUND
	Dim sHomeStyle
	Dim sAddressesStyle
	Dim sReportsStyle
	Dim sOptionsStyle
	Dim sSubscriptionsStyle
	Dim iSubscribeWizardStep
	Dim iAddressWizardStep
	Dim nStart
	asReservedChars = Array("\", """", ">", "&", "#", "?", "'", "+", "<")
	STYLE_BEIGE_BACKGROUND = "background-image: url('images/bg_beige.gif'); background-repeat: repeat-x"
	sHomeStyle = STYLE_BEIGE_BACKGROUND
	sAddressesStyle = STYLE_BEIGE_BACKGROUND
	sReportsStyle = STYLE_BEIGE_BACKGROUND
	sOptionsStyle = STYLE_BEIGE_BACKGROUND
	sSubscriptionsStyle = STYLE_BEIGE_BACKGROUND
	iSubscribeWizardStep = 0

	nStart = -1
	If Len(CStr(oRequest("start"))) > 0 Then
		nStart = CInt(oRequest("start"))
	Else
		nStart = -1
	End If

	'''''''''''''''
	'Castor
	'''''''''''''''
	'Page Information

	Dim oSession
	Dim oObjServer

	'Variables for the toolbar
	Dim iHelpFileID
	Dim aiQuestions()
	Dim iIndexForQuestions

	'Error Handling variables
	Dim lErrNumber
	Dim sErrDescription
	Dim sTemp

	'** BEGIN: TOOLBAR SECTION
	aSourceInfo(0) = USER_COOKIE
    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	If Len(CStr(oRequest("toolbar"))) > 0 Then
		aSourceInfo(1) = "Toolbar"
		Call WriteToSource(aConnectionInfo, CStr(oRequest("toolbar")), Application.Value("iSourcePerm"), aSourceInfo)
	End If
	If Len(CStr(oRequest("showSearch"))) > 0 Then
		aSourceInfo(1) = "SearchSection"
		Call WriteToSource(aConnectionInfo, CStr(oRequest("showSearch")), Application.Value("iSourcePerm"), aSourceInfo)
	End If
	If Len(CStr(oRequest("showHelp"))) > 0 Then
		aSourceInfo(1) = "HelpSection"
		Call WriteToSource(aConnectionInfo, CStr(oRequest("showHelp")), Application.Value("iSourcePerm"), aSourceInfo)
	End If

	aSourceInfo(1) = "Toolbar"
	If StrComp(ReadFromSource(aConnectionInfo, Application.Value("iSourcePerm"), aSourceInfo), "0") <> 0 Then aPageInfo(N_TOOLBARS_PAGE) = aPageInfo(N_TOOLBARS_PAGE) + MAIN_TOOLBAR
	aSourceInfo(1) = "SearchSection"
	If StrComp(ReadFromSource(aConnectionInfo, Application.Value("iSourcePerm"), aSourceInfo), "0") <> 0 Then aPageInfo(N_TOOLBARS_PAGE) = aPageInfo(N_TOOLBARS_PAGE) + SEARCH_TOOLBAR
	aSourceInfo(1) = "HelpSection"
	If StrComp(ReadFromSource(aConnectionInfo, Application.Value("iSourcePerm"), aSourceInfo), "0") <> 0 Then aPageInfo(N_TOOLBARS_PAGE) = aPageInfo(N_TOOLBARS_PAGE) + HELP_TOOLBAR
	iHelpFileID = -1
	'** END: TOOLBAR SECTION

	'Variables added for Web - Hydra prompt integration
	Dim bMFCBlaster
	bMFCBlaster = false ' This variable is always false in Hydra


%>