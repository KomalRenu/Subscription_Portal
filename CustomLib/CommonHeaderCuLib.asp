<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%

Public Const WEBADMIN_PROGID = "MSIXMLLib.DSSXMLAdmin.1"

Private Const S_SERVER_NAME_CONNECTION = 0
Private Const S_PROJECT_CONNECTION = 1
Private Const N_PORT_CONNECTION = 2
Private Const S_UID_CONNECTION = 3
Private Const S_PWD_CONNECTION = 4
Private Const S_TOKEN_CONNECTION = 5
Private Const N_PRIVILEGES_CONNECTION = 6
Private Const S_IP_ADDRESS_CONNECTION = 7
Private Const S_PROJECT_URL_CONNECTION = 8
Private Const N_USER_MODE_CONNECTION = 9
Private Const S_PROJECT_ALIAS_CONNECTION = 10
Private Const S_UID_PLUGIN_CONNECTION = 11
Private Const S_PWD_PLUGIN_CONNECTION = 12
Private Const S_SITEID_CONNECTION = 13

Private Const STD_MODE_CONNECTION = 1
Private Const NT_MODE_CONNECTION = 2
Private Const GUEST_MODE_CONNECTION = 4

Private Const MAX_CONNECTION_INFO = 12

Private Const S_TITLE_PAGE = 0
Private Const N_CURRENT_OPTION_PAGE = 1
Private Const N_OPTIONS_WITH_LINKS_PAGE = 2
Private Const N_TOOLBARS_PAGE = 3
Private Const N_BG_COLOR_PAGE = 4
Private Const S_NAME_PAGE = 5
Private Const S_FOLDER_ID_PAGE = 6
Private Const N_ALIAS_PAGE = 7
Private Const S_ROOT_FOLDER_ID_PAGE = 8
Private Const S_ROOT_FOLDER_NAME_PAGE = 9
Private Const S_MY_REPORTS_PAGE = 10
Private Const S_SHARED_REPORTS_PAGE = 11
Private Const S_CREATE_REPORTS_PAGE = 12

Private Const MAX_PAGE_INFO = 12

Function GetPageWithoutPath(sPageNameWithPath)
'*******************************************************************************
'Purpose:	Get the path for the current page
'Inputs:	sPageNameWithPath
'Outputs:	Path for the current page
'*******************************************************************************
	On Error Resume Next
	Dim sEndOfPageName
	GetPageWithoutPath = LCase(Right(sPageNameWithPath, Len(sPageNameWithPath) - (InStrRev(sPageNameWithPath, "/"))))
	sEndOfPageName = InStr(1, GetPageWithoutPath, "?", vbBinaryCompare)
	If sEndOfPageName > 0 Then
		GetPageWithoutPath = LCase(Left(GetPageWithoutPath, sEndOfPageName))
	End If
	Err.Clear
End Function
%>