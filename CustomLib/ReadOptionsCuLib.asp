<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Private Const USER_OPTIONS_SCOPE = 1
Private Const ADMIN_OPTIONS_SCOPE = 2
Private Const ALL_OPTIONS_SCOPE = 3
Private Const HARDCODED_OPTIONS_SCOPE = 4

Private Const FORM_GET_METHOD = "1"
Private Const FORM_POST_METHOD = "2"

Private Const REFRESH_OPTION = "1"
Private Const REEXECUTE_OPTION = "2"

Private Const NO_CONNECTION_INFO = -2
Private Const NOT_FOUND = -1

Private Const ADVANCED_DRILL = 1
Private Const DRILL_DOWN_DRILL = 2
Private Const SIMPLE_DRILL = 3
Private Const NO_DRILL = 4
Private Const DESKTOP_DRILL = 5
Private Const DONT_KEEP_PARENT_DRILL = 0
Private Const KEEP_PARENT_DRILL = 1
Private Const DESKTOP_KEEP_PARENT_DRILL = 2

Private Const COMMA_DELIMITER = 1
Private Const TAB_DELIMITER = 2
Private Const SEMICOLON_DELIMITER = 3
Private Const SPACE_DELIMITER = 4

Private Const OPTION_USE_REPORT_STYLE = "1"
Private Const OPTION_USE_USER_STYLE = "2"

Private Const OPTION_SAVE_START_URL = 1
Private Const OPTION_SAVE_RETURN_URL = 2

'Total number of user options
Private Const OPTIONS = 111
Private Const MAX_GLOBAL_OPTIONS = 2

Private Const GRID_STYLE_OPTION = "aa"								'"gridStyle"
Private Const USE_DEFAULT_GRID_STYLE_OPTION = "ab"					'"useDefaultGridStyle"
Private Const GRID_ROWS_OPTION = "ac"								'"gridRows"
Private Const GRID_COLUMNS_OPTION = "ad"							'"gridColumns"
Private Const LOCALE_OPTION = "ae"									'"LOCALE_OPTION"
Private Const EXPORT_FORMAT_OPTION = "af"									'"EXPORT_FORMAT_OPTION"
Private Const USER_HEADER_OPTION = "ag"							'"userHeader"
Private Const USER_FOOTER_OPTION = "ah"							'"userFooter"
Private Const PRINT_ROWS_OPTION = "ai"								'"printRows"
Private Const PRINT_COLUMNS_OPTION = "aj"							'"printColumns"
Private Const MAX_PRINT_ROWS_OPTION = "ak"							'"maxPrintRows"
Private Const MAX_PRINT_COLUMNS_OPTION = "al"						'"maxPrintColumns"
Private Const START_PAGE_OPTION = "am"								'"startPage"
Private Const CANCEL_JOBS_OPTION = "an"							'"cancelJobs"
Private Const DELETE_JOBS_OPTION = "ao"							'"deleteJobs"
Private Const SAVE_PWD_OPTION = "ap"								'"savePWD"
Private Const AUTHENTICATION_OPTION = "aq"						'"NTAuthentication"
Private Const INBOX_EXPIRES_OPTION = "ar"							'"inboxExpires"
Private Const INBOX_MAX_JOBS_OPTION = "as"							'"inboxMaxJobs"
Private Const INBOX_MAX_SIZE_OPTION = "at"							'"inboxMaxSize"
Private Const GRAPH_SIZE_OPTION = "au"								'"graphSize"
Private Const GRAPH_HEIGHT_OPTION = "av"							'"graphHeight"
Private Const GRAPH_WIDTH_OPTION = "aw"							'"graphWidth"
Private Const OBJECT_BROWSING_OPTION = "ax"						'"objectBrowsing"
Private Const EXPORT_NEW_WINDOW_OPTION = "ay"						'"excelNewWindow"
Private Const PAGE_NUMBERING_OPTION = "az"							'"pageNumbering"
Private Const PAGE_LAYOUT_OPTION = "ba"							'"pageLayout"
Private Const ADMIN_HEADER_OPTION = "bb"							'"adminHeader"
Private Const ADMIN_FOOTER_OPTION = "bc"							'"adminFooter"
Private Const MAX_EXPORT_COLS_OPTION= "bd"							'"maxExportCols"
Private Const GRID_CSS_OPTION= "bf"								'"gridCSS"
Private Const MAX_EXPORT_ROWS_OPTION= "bg"							'"maxExportRows"
Private Const MAX_EXPORT_ROWS_TEXT_OPTION= "bh"					'"maxExportRowsText"
Private Const DEFAULT_AUTHENTICATION_OPTION = "bi"							'"allowGuest"
Private Const PIVOT_MODE_OPTION = "bk"								'"pivotMode"
Private Const INBOX_SORT_OPTION = "bm"								'"inboxSort"
Private Const CANCEL_JOBS_PROMPT_OPTION = "bn"						'"CancelJobsPrompt"
Private Const DELETE_JOBS_PROMPT_OPTION = "bo"						'"DeleteJobsPrompt"
Private Const USE_SECURITY_PLUGIN_OPTION = "bp"					'"UseSecurityPlugin"
Private Const SECURITY_PLUGIN_CLASS_OPTION = "bq"					'"SecurityPluginClass"
Private Const SECURITY_PLUGIN_FREQ_OPTION = "br"					'"SecurityPluginFreq"
Private Const PROMPT_ON_PRINT_OPTION = "bs"						'"PromptOnPrint"
Private Const PROMPT_ON_EXPORT_OPTION = "bt"						'"PromptOnExport"
Private Const EXECUTION_MODE_OPTION = "bu"							'"ExecutionMode"
Private Const EXECUTION_WAIT_TIME_OPTION = "bw"					'"ExecutionWaitTime"
Private Const PRINT_FILTER_DETAILS_OPTION = "by"					'"PrintFilterDetails"
Private Const EXPORT_FILTER_DETAILS_OPTION = "bz"					'"ExportFilterDetails"
Private Const ELE_PROMPT_BLOCK_COUNT_OPTION = "ca"					'"ElePromptBlockCount"
Private Const EXPORT_SECTION_OPTION = "cb"							'"ExportSection"
Private Const SHOW_DATA_OPTION = "cc"								'"ShowData"
Private Const SHOW_TOOLBAR_OPTION = "cd"							'"ShowToolbar"
Private Const SHOW_HELP_OPTION = "ce"								'"ShowHelp"
Private Const WAIT_TIME_IN_WAIT_PAGE_OPTION = "cf"					'"WaitTimeInWaitPage"
Private Const ADMIN_DELETE_JOBS_OPTION = "cg"						'"AdminDeleteJobs"
Private Const OBJ_PROMPT_BLOCK_COUNT_OPTION = "ch"					'"ObjPromptBlockCount"
Private Const FORM_METHOD_OPTION = "ci"							'"FormMethod"
Private Const DEFAULT_START_DOCUMENT_OPTION = "cj"					'"DefaultStartDocument"
Private Const ADMIN_CANCEL_JOBS_OPTION = "ck"						'"AdminCancelJobs"
Private Const ICON_VIEW_MODE_OPTION = "cl"							'"IconViewMode"
Private Const MAX_SEARCH_RESULTS_OPTION = "cm"						'"MaxSearchResults"
Private Const SEARCH_TIMEOUT_OPTION = "cn"							'"SearchTimeout"
Private Const DEFAULT_OBJECT_NAME_OPTION = "co"					'"DefaultObjectName"
Private Const MAX_PROJECT_NAME_OPTION = "cp"						'"MaxProjectName"
Private Const PROJECT_ALIAS_OPTION = "cq"							'"ProjectAlias"
Private Const MAX_PROJECT_TABS_OPTION = "cr"						'"MaxProjectTabs"
Private Const SHOW_TAB_OPTION = "cs"								'"ShowTab"
Private Const ADD_TO_MY_HISTORY_LIST_OPTION = "ct"					'"Add to my history list"
Private Const ORIENTATION_PREVIEW_OPTION = "cv"						'"Orientation Preview"
Private Const PAPER_SIZE_PREVIEW_OPTION = "cw"						'"Paper Size Preview"
Private Const ADMIN_CONTACT_INFO_OPTION = "cx"						'"Admin Contact Info"
Private Const START_PAGE_RADIO_OPTION = "cz"						'"Start Page Option"
Private Const DROPDOWN_START_PAGE_URL_OPTION = "da"					'"Dropdown Start Page URL"
Private Const CURRENT_START_PAGE_NAME_OPTION = "db"					'"Current Start Page Name"
Private Const CURRENT_START_PAGE_URL_OPTION = "dc"					'"Current Start Page URL"
Private Const NEW_START_PAGE_NAME_OPTION = "dd"						'"New Start Page Name"
Private Const NEW_START_PAGE_URL_OPTION = "de"						'"New Start Page URL"
Private Const ACTUAL_START_PAGE_URL_OPTION = "df"					'"Actual Start Page URL"
Private Const ACTUAL_START_PAGE_NAME_OPTION = "dg"					'"Actual Start Page Name"
Private Const PROMPT_ADMIN_CANCEL_JOBS_OPTION = "dh"					'"Prompt Admin Cancel Jobs Option"
Private Const PROMPT_ADMIN_DELETE_JOBS_OPTION = "di"					'"Prompt Admin Delete Jobs Option"
Private Const ALLOW_USER_EXPORT_TEMP_FILES_OPTION= "dj"				'
Private Const ALLOW_USER_USE_DHTML_OPTION = "dk"					'
Private Const KEEP_PARENT_OPTION = "dl"					'
Private Const USE_EXPORT_TEMP_FILES_OPTION = "dm"					'"Use Export Temp Files"
Private Const WORKING_SET_SIZE_OPTION = "do"					'
Private Const DRILL_PRIVILEGE_OPTION = "dp"					'
Private Const USE_DHTML_PROMPTS_OPTION = "dq"						'"Use Dynamic HTML Prompts"
Private Const USE_DHTML_VALUE_OPTION = "dr"
Private Const SEARCH_WORKING_SET_SIZE_OPTION = "ds"
Private Const GRAPH_FORMAT_OPTION = "dt"
Private Const PREVIEW_NEW_WINDOW_OPTION = "du"
Private Const EXPORT_VALUES_AS_TEXT_OPTION = "dv"
Private Const REFRESH_METHOD_OPTION = "dw"							'"Refresh Method - Refresh/Re-execute"
Private Const PRINT_GRID_AND_GRAPH_TOGETHER_OPTION = "dx"
Private Const REPEAT_COLHEADERS_OPTION = "dy"							'Repeat Column Headers YES/NO/From report definition
Private Const EXPORT_ROWS_FORMATTING_OPTION = "dz"
Private Const EXPORT_COLS_SPREADSHEET_OPTION = "ea"
Private Const EXPORT_GRAPH_FORMAT_OPTION = "eb"
Private Const EXPORT_GRAPH_AND_GRID_FORMAT_OPTION = "ec"
Private Const POP_UP_DRILL_OPTION = "ed"
Private Const PROMPTS_ON_ONE_PAGE_OPTION = "ef"
Private Const REQUIRED_PROMPTS_FIRST_OPTION = "eg"
Private Const GRAPH_CATEGORIES_AND_SERIES_FROM_DESKTOP_OPTION = "eh"
Private Const GRAPH_CATEGORIES_OPTION = "ei"
Private Const GRAPH_SERIES_OPTION = "ej"
Private Const EXPORT_PLAINTEXT_DELIMITER_OPTION = "ek"
Private Const SEARCH_OBJECTS_OPTION = "el"
Private Const DOCUMENT_EXPORT_FORMAT_OPTION = "em"
Private Const ALLOW_USER_SET_GRAPH_SETTINGS_OPTION = "en"
Private Const REPORT_TAB_OPTION = "eo"								'The tab that is opened in the report widget
Private Const REUSE_MESSAGE_FOR_SCHEDULED_REPORTS_OPTION = "ep"
Private Const ALLOWED_FILE_EXTENSION_OPTION = "et"
Private Const MAXIMUM_FILE_SIZE_TO_UPLOAD_OPTION = "eu"
Private Const MAX_ELEMENTS_TO_IMPORT_OPTION = "ey"
Private Const KEEP_WHITESPACE_IN_PROMPTS_OPTION = "it"
Private Const DEFAULT_PROMPT_MATCH_CASE_OPTION = "ip"
Private Const ACCESSIBILITY_OPTION = "fx"
'Error constants
Private Const ERROR_IN_OPTIONS = -1

'XML file
Private Const XML_SETTING_FILE = "AdminOptions.xml"

Function ReadUserOption(sOptionKey)
'*******************************************************************************
'Purpose:   Retrieves the value of a certain option from the cookie
'Inputs:	sOptionKey
'Outputs:
'*******************************************************************************
	On Error Resume Next



	Select Case(sOptionKey)

	Case REQUIRED_PROMPTS_FIRST_OPTION

		ReadUserOption = ""

	Case ALLOWED_FILE_EXTENSION_OPTION

		ReadUserOption = "txt,csv"

	Case MAX_ELEMENTS_TO_IMPORT_OPTION

		ReadUserOption = 300

	Case MAXIMUM_FILE_SIZE_TO_UPLOAD_OPTION

		ReadUserOption = "100"

	Case KEEP_WHITESPACE_IN_PROMPTS_OPTION

		ReadUserOption = ""

	Case ELE_PROMPT_BLOCK_COUNT_OPTION
			If Len(Application.Value("ELE_PROMPT_BLOCK_COUNT_OPTION")) > 0 Then
				ReadUserOption = CLng(Application.Value("ELE_PROMPT_BLOCK_COUNT_OPTION"))
			Else
				ReadUserOption = CONST_ELEPROMPT_BLOCKCOUNT
			End IF

	Case OBJ_PROMPT_BLOCK_COUNT_OPTION

			If Len(Application.Value("OBJ_PROMPT_BLOCK_COUNT_OPTION")) > 0 Then
				ReadUserOption = CLng(Application.Value("OBJ_PROMPT_BLOCK_COUNT_OPTION"))
			Else
				ReadUserOption = CONST_OBJPROMPT_BLOCKCOUNT
			End IF

	Case DEFAULT_PROMPT_MATCH_CASE_OPTION
			If Application.Value("PROMPT_MATCH_CASE_OPTION") = 0 Then
				ReadUserOption = "unchecked"
			Else
				ReadUserOption = "checked"
			End IF

	Case ACCESSIBILITY_OPTION

		ReadUserOption = ""

	Case Else

		aSourceInfo(0) = "CurrUsrOpt"
		aSourceInfo(1) = sOptionKey
		aSourceInfo(2) = ""
		ReadUserOption = ReadFromSource(aConnectionInfo, Application.Value("iSourcePerm"), aSourceInfo)

	End Select


	Err.Clear
End Function


%>