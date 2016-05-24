<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/OptionsCoLib.asp" -->
<%

	Function ParseRequestForOptions(oRequest, sSiteLanguage, sOptSection, sUseJavaScript, sStartPage, sLocale, sSummaryPage)
	'********************************************************
	'*Purpose:
	'*Inputs: oRequest
	'*Outputs: sSiteLanguage, sOptSection, sUseJavaScript, sStartPage
	'********************************************************
		On Error Resume Next
		Dim lErrNumber
		Dim sLocaleData
		Dim iDelimiter

		lErrNumber = NO_ERR

		sSiteLanguage = ""
		sOptSection = ""
		sUseJavaScript = ""
		sStartPage = ""
		sLocale = ""

		sOptSection = Trim(CStr(oRequest("optSection")))
		sUseJavaScript = Trim(CStr(oRequest("useJavaScript")))
		sStartPage = Trim(CStr(oRequest("startPage")))
		sSummaryPage = Trim(CStr(oRequest("summary")))

		sLocaleData = Trim(CStr(oRequest("Locale")))
		If Len(sLocaleData) > 0 Then
		    iDelimiter = Instr(1, sLocaleData, ";")
		    sLocale = Left(sLocaleData, iDelimiter - 1)
		    sSiteLanguage = Mid(sLocaleData, iDelimiter + 1)
		End If

		If Err.number <> NO_ERR Then
		    lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "ParseRequestForOptions", "", "Error setting variables equal to Request variables", LogLevelError)
		End If

		ParseRequestForOptions = lErrNumber
		Err.Clear
	End Function

	Function ParseRequestForChangePassword(oRequest, sOldPassword, sNewPassword, sConfirmNewPassword, sHint)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'********************************************************
		On Error Resume Next
		Dim lErrNumber
		lErrNumber = NO_ERR

		sOldPassword = ""
		sNewPassword = ""
		sConfirmNewPassword = ""
		sHint = ""

		sOldPassword = Trim(CStr(oRequest("oldPwd")))
		sNewPassword = Trim(CStr(oRequest("newPwd")))
		sConfirmNewPassword = Trim(CStr(oRequest("confirmNewPwd")))
		sHint = Trim(CStr(oRequest("Hint")))

		If Err.number <> NO_ERR Then
		    lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "ParseRequestForChangePassword", "", "Error setting variables equal to Request variables", LogLevelError)
		End If

		ParseRequestForChangePassword = lErrNumber
		Err.Clear
	End Function

	Function validate_ChangeUserPassword(sOldPassword, sNewPassword, sConfirmNewPassword, sHint)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'********************************************************
	    On Error Resume Next
	    Dim lErrNumber
	    lErrNumber = NO_ERR

	    If Len(sOldPassword) = 0 Or Len(sNewPassword) = 0 Then
			lErrNumber = lErrNumber + ERR_LOGIN_BLANKS
        End If

        If sNewPassword <> sConfirmNewPassword Then
			lErrNumber = lErrNumber + ERR_CONFIRM_PASSWORD
		End If

		If Len(sHint) = 0 Then
		    lErrNumber = lErrNumber + ERR_HINT_BLANK
		End If

	    validate_ChangeUserPassword = lErrNumber
	    Err.Clear
	End Function

Function cu_UpdateUserPassword(sOldPassword, sNewPassword)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_UpdateUserPassword"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    If lErrNumber = NO_ERR Then
        lErrNumber = co_UpdateUserPassword(sSessionID, sOldPassword, sNewPassword)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_UpdateUserPassword", LogLevelTrace)
        End If
    End If

	cu_UpdateUserPassword = lErrNumber
	Err.Clear
End Function

Function RenderLanguageOptions()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim sLng
	Dim lErrNumber

	lErrNumber = NO_ERR
	sLng = GetLng()

	Response.Write "<select name=""siteLanguage"" class=""pullDownClass"">"

	Response.Write "<option value=""1031"""
	If sLng = "1031" Then Response.Write " SELECTED"
	Response.Write ">" & asDescriptors(309) & "</option>" 'Descriptor: German

	Response.Write "<option value=""1033"""
	If sLng = "1033" Then Response.Write " SELECTED"
	Response.Write ">" & asDescriptors(490) & "</option>" 'Descriptor: English

	Response.Write "<option value=""1034"""
	If sLng = "1034" Then Response.Write " SELECTED"
	Response.Write ">" & asDescriptors(489) & "</option>" 'Descriptor: Spanish

	Response.Write "<option value=""1036"""
	If sLng = "1036" Then Response.Write " SELECTED"
	Response.Write ">" & asDescriptors(310) & "</option>" 'Descriptor: French

	Response.Write "<option value=""1040"""
	If sLng = "1040" Then Response.Write " SELECTED"
	Response.Write ">" & asDescriptors(311) & "</option>" 'Descriptor: Italian

	Response.Write "<option value=""1041"""
	If sLng = "1041" Then Response.Write " SELECTED"
	Response.Write ">" & asDescriptors(312) & "</option>" 'Descriptor: Japanese

	Response.Write "<option value=""1042"""
	If sLng = "1042" Then Response.Write " SELECTED"
	Response.Write ">" & asDescriptors(313) & "</option>" 'Descriptor: Korean

	Response.Write "<option value=""1046"""
	If sLng = "1046" Then Response.Write " SELECTED"
	Response.Write ">" & asDescriptors(332) & "</option>" 'Descriptor: Portuguese (Brazilian)

	Response.Write "<option value=""1053"""
	If sLng = "1053" Then Response.Write " SELECTED"
	Response.Write ">" & asDescriptors(314) & "</option>" 'Descriptor: Swedish

	Response.Write "</select>"

    If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "RenderLanguageOptions", "", "Error rendering language options", LogLevelError)
    End If

	RenderLanguageOptions = lErrNumber
	Err.Clear
End Function

Function ChangeLanguage(sLanguage, asDescriptors, aFontInfo)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber

	lErrNumber = NO_ERR

	If sLanguage <> "" Then
		If GetLng() <> CStr(sLanguage) Then
			aSourceInfo(0) = SITE_COOKIE
	        aSourceInfo(1) = "Lng"
	        aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	        Call WriteToSource(aConnectionInfo, CStr(sLanguage), SOURCE_COOKIES, aSourceInfo)
			If Err.number <> NO_ERR Then
				lErrNumber = Err.number
				Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "ChangeLanguage", "", "Error calling WriteToSource for Lng", LogLevelError)
			End If
		End If
	End If

    If lErrNumber = NO_ERR Then
        Call SetLocaleInformation(asDescriptors, aFontInfo)
    End If

	ChangeLanguage = lErrNumber
    Err.Clear
End Function

Function ChangeJavaScript(sUseJavaScript)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	lErrNumber = NO_ERR

	If sUseJavaScript <> "" Then
		If GetJavaScriptPreference() <> CStr(sUseJavaScript) Then
			aSourceInfo(0) = SITE_COOKIE
	        aSourceInfo(1) = PREF_JAVASCRIPT
	        aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	        Call WriteToSource(aConnectionInfo, CStr(sUseJavaScript), SOURCE_COOKIES, aSourceInfo)
			If Err.number <> NO_ERR Then
				lErrNumber = Err.number
				Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "ChangeJavaScript", "", "Error calling WriteToSource for USE_JAVASCRIPT", LogLevelError)
			Else
		        Select Case sUseJavaScript
		            Case "2"
			            aSourceInfo(0) = SITE_COOKIE
	                    aSourceInfo(1) = USE_JAVASCRIPT
	                    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	                    Call WriteToSource(aConnectionInfo, CStr(SupportsJavaScript()), SOURCE_COOKIES, aSourceInfo)
		            Case Else
			            aSourceInfo(0) = SITE_COOKIE
	                    aSourceInfo(1) = USE_JAVASCRIPT
	                    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
	                    Call WriteToSource(aConnectionInfo, CStr(sUseJavaScript), SOURCE_COOKIES, aSourceInfo)
		        End Select
			End If
		End If
	End If

	ChangeJavaScript = lErrNumber
	Err.Clear
End Function

Function ChangeStartPage(sStartPage)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    lErrNumber = NO_ERR

    If sStartPage <> "" Then
        If GetStartPage() <> CStr(sStartPage) Then
            Call SetStartPage(sStartPage)
            If Err.number <> NO_ERR Then
                lErrNumber = Err.number
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "ChangeStartPage", "", "Error calling SetStartPage", LogLevelError)
            End If
        End If
    End If

    ChangeStartPage = lErrNumber
    Err.Clear
End Function

Function ChangeLocale(sLocale)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim sGetUserPropertiesXML
    Dim oUserPropertiesDOM
    Dim sDefAddID 'Default Address ID
    Dim sHint   'Password hint
    Dim sCurrentLocale

    lErrNumber = NO_ERR

    lErrNumber = cu_GetUserProperties(sGetUserPropertiesXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "ChangeLocale", "", "Error calling cu_GetUserProperties", LogLevelTrace)
    Else
        Set oUserPropertiesDOM = Server.CreateObject("Microsoft.XMLDOM")
	    oUserPropertiesDOM.async = False
	    oUserPropertiesDOM.loadXML(sGetUserPropertiesXML)
	    sCurrentLocale = oUserPropertiesDOM.selectSingleNode("/mi/prs/pr[@n = 'MR_LOCALE_ID']").getAttribute("v")
	    sDefAddID = oUserPropertiesDOM.selectSingleNode("/mi/prs/pr[@n = 'MR_DEF_ADD_ID']").getAttribute("v")
	    sHint = oUserPropertiesDOM.selectSingleNode("/mi/prs/pr[@n = 'MR_PWD_HINT']").getAttribute("v")
    End If

    If (lErrNumber = NO_ERR) And (sLocale <> sCurrentLocale) Then
        lErrNumber = cu_UpdateUserProperties(sHint, sLocale, sDefAddID)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "ChangeLocale", "", "Error calling cu_UpdateUserProperties", LogLevelTrace)
        Else
		    aSourceInfo(0) = SITE_COOKIE
		    aSourceInfo(1) = SITE_LOCALE
		    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
		    Call WriteToSource(aConnectionInfo, CStr(sLocale), SOURCE_COOKIES, aSourceInfo)
		    If Err.number <> NO_ERR Then
		    	lErrNumber = Err.number
		    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "ChangeLocale", "", "Error calling WriteToSource for SITE_LOCALE", LogLevelError)
		    End If
        End If
    End If

    Set oUserPropertiesDOM = Nothing

    ChangeLocale = lErrNumber
    Err.Clear
End Function

Function ChangeSummaryPage(sSummaryPage)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    lErrNumber = NO_ERR

    If Len(sSummaryPage) > 0 Then
        If Strcomp(GetSummaryPageSetting(), CStr(sSummaryPage)) <> 0 Then
            Call SetSummaryPageSetting(sSummaryPage)
            If Err.number <> NO_ERR Then
                lErrNumber = Err.number
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "ChangeSummaryPage", "", "Error calling SetSummaryPage", LogLevelError)
            End If
        End If
    End If

    ChangeSummaryPage = lErrNumber
    Err.Clear
End Function

Function validate_ChangeAuthentications(oRequest)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim Item
    lErrNumber = NO_ERR

	For Each Item in oRequest
        If Left(Item, 8) = "AO_User_" Then
           If Len(CStr(oRequest(Item))) = 0 Then
                If CStr(oRequest("AO_Req_" & Right(Item, Len(Item) - 8))) = "1" Then
                    lErrNumber = lErrNumber + ERR_ISLOGIN_BLANK
                    Exit For
                ElseIf CStr(oRequest("AO_Password_" & Right(Item, Len(Item) - 8))) <> "" Then
                    lErrNumber = lErrNumber + ERR_ISLOGIN_ERROR
                    Exit For
                End If
            End If
        End If
    Next

    validate_ChangeAuthentications = lErrNumber
    Err.Clear
End Function

Function cu_UpdateUserAuthenticationObjects(sUserAuthenticationObjectsXML, oRequest)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Get Castor XML Structure
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim sSessionID
	Dim iArrayCounter
    Dim iArrayCounterForSave
    Dim iArrayCounterForUpdate
	Dim Item
	Dim sServerName
	Dim sProjectID
	Dim lPort
	Dim sUserName
	Dim sPwd
	Dim sUserID
	Dim sISID
	'Dim sUserAuthenticationObjectsXML
	Dim asInformationSourceIDForSave()
    Dim asInformationSourceIDForUpdate()
    Dim asXMLAuthenticationStringForSave()
    Dim asXMLAuthenticationStringForUpdate()
    Dim oUserAuth
    Dim oUserAuthDOM
    Dim iArraySize

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
    iArrayCounterForSave = 0
    iArrayCounterForUpdate = 0
    iArrayCounter = 0
    iArraySize = CInt(oRequest("AO_COUNT"))
	Redim asInformationSourceIDForSave(iArraySize)
    Redim asInformationSourceIDForUpdate(iArraySize)
    Redim asXMLAuthenticationStringForSave(iArraySize)
    Redim asXMLAuthenticationStringForUpdate(iArraySize)

    lErr = LoadXMLDOMFromString(aConnectionInfo, sUserAuthenticationObjectsXML, oUserAuthDOM)
	If lErr <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	End If

	For Each Item in oRequest
        If strcomp(Left(Item, 8), "AO_User_") = 0 Then

			sISID = Right(Item, Len(Item) - 8)
            sServerName = CStr(oRequest("AO_Serv_" & sISID))
            sProjectID = CStr(oRequest("AO_Proj_" & sISID))
            lPort = CLng(oRequest("AO_Port_" & sISID))
            sUserName = CStr(oRequest(Item))
            sPwd = CStr(oRequest("AO_Password_" & sISID))
            sUserID = ""

            Set oUserAuth = oUserAuthDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_INFORMATION_SOURCE & "' $and$ @id='" & sISID & "']")

			If Len(sUserName) > 0 Then
				'User Detail IS doesn't need validation
                If (Len(sServerName) > 0) And (Len(sProjectID) > 0) And (Len(CStr(lPort)) > 0) Then
                    lErrNumber = UserAuthenticate(sServerName, sProjectID, lPort, sUserName, sPwd, sUserID)
                    If lErrNumber <> NO_ERR Then
                        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "cu_SaveUserAuthentications", "", "Error calling UserAuthenticate", LogLevelTrace)
                        Exit For
                    ElseIf Len(sUserID) = 0 Then
                        lErrNumber = ERR_ISLOGIN_ERROR
                        Exit For
                    End If
                End If

                If oUserAuth is Nothing Then
					asInformationSourceIDForSave(iArrayCounterForSave) = CStr(sISID)
					Call BuildAuthenticationObject(sUserName, sPwd, sUserID, asXMLAuthenticationStringForSave(iArrayCounterForSave))
					iArrayCounterForSave = iArrayCounterForSave + 1
				Else
					asInformationSourceIDForUpdate(iArrayCounterForUpdate) = CStr(sISID)
					Call BuildAuthenticationObject(sUserName, sPwd, sUserID, asXMLAuthenticationStringForUpdate(iArrayCounterForUpdate))
					iArrayCounterForUpdate = iArrayCounterForUpdate + 1
				End If
            Else
                'If the user is clearing the IS Authentication, send a blank string
				asInformationSourceIDForUpdate(iArrayCounterForUpdate) = CStr(sISID)
				Call BuildAuthenticationObject(sUserName, sPwd, sUserID, asXMLAuthenticationStringForUpdate(iArrayCounterForUpdate) )
				iArrayCounterForUpdate = iArrayCounterForUpdate + 1
            End If

            iArrayCounter = iArrayCounter + 1
            If iArrayCounter = iArraySize Then
                Exit For
            End If
        End If
    Next

    If (lErrNumber = NO_ERR) Then
        If iArrayCounterForSave > 0 Then
			Redim Preserve asInformationSourceIDForSave(iArrayCounterForSave-1)
			Redim Preserve asXMLAuthenticationStringForSave(iArrayCounterForSave-1)
			lErrNumber = co_SaveUserAuthenticationObjects(sSessionID, asInformationSourceIDForSave, asXMLAuthenticationStringForSave)
			If lErrNumber <> NO_ERR Then
			    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "cu_UpdateUserAuthenticationObjects", "", "Error calling co_SaveUserAuthenticationObjects", LogLevelTrace)
			End If
			'Response.Write "Save" & iArrayCounterForSave
        End If
        If iArrayCounterForUpdate > 0 Then
			Redim Preserve asInformationSourceIDForUpdate(iArrayCounterForUpdate-1)
			Redim Preserve asXMLAuthenticationStringForUpdate(iArrayCounterForUpdate-1)
			lErrNumber = co_UpdateUserAuthenticationObjects(sSessionID, asInformationSourceIDForUpdate, asXMLAuthenticationStringForUpdate)
        	If lErrNumber <> NO_ERR Then
			    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "cu_UpdateUserAuthenticationObjects", "", "Error calling co_SaveUserAuthenticationObjects", LogLevelTrace)
			End If
			'Response.Write "Update" & iArrayCounterForUpdate
        End If
    End If

	Erase asInformationSourceIDForSave
    Erase asInformationSourceIDForUpdate
    Erase asXMLAuthenticationStringForSave
    Erase asXMLAuthenticationStringForUpdate
    Set oUserAuth = Nothing
    Set oUserAuthDOM = Nothing

	cu_UpdateUserAuthenticationObjects = lErrNumber
	Err.Clear
End Function

Function GetVariablesFromXML_ChangePassword(sGetUserPropertiesXML, sLocaleID, sDefAddID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oPropsDOM
    Dim oProperties

    lErrNumber = NO_ERR

	Set oPropsDOM = Server.CreateObject("Microsoft.XMLDOM")
	oPropsDOM.async = False
	If oPropsDOM.loadXML(sGetUserPropertiesXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "GetVariablesFromXML_ChangePassword", "", "Error loading sGetUserPropertiesXML", LogLevelError)
    Else
        Set oProperties = oPropsDOM.selectSingleNode("/mi/prs")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "OptionsCuLib.asp", "GetVariablesFromXML_ChangePassword", "", "Error retrieving prs node", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
        sLocaleID = CStr(oProperties.selectSingleNode("pr[@n = 'MR_LOCALE_ID']").getAttribute("v"))
        sDefAddID = CStr(oProperties.selectSingleNode("pr[@n = 'MR_DEF_ADD_ID']").getAttribute("v"))
    End If

    Set oPropsDOM = Nothing
    Set oProperties = Nothing

    GetVariablesFromXML_ChangePassword = lErrNumber
    Err.Clear
End Function

Function RenderSummaryPageChoices()
'********************************************************
'*Purpose:	Display 3 choices for summary page
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim sSelectedChoice

    sSelectedChoice = CStr(GetSummaryPageSetting())

	Response.Write "<INPUT NAME=""summary"" TYPE=""radio"" VALUE=""1"" "
	If Clng(sSelectedChoice) = SITE_PROPVALUE_SUMMARY_PAGE_ALWAYS Then
		Response.Write "CHECKED=""1"" "
	End If
	Response.Write "/><FONT FACE=" & aFontInfo(S_FAMILY_FONT) & " SIZE=" & aFontInfo(N_SMALL_FONT) & ">" & asDescriptors(35) & "</FONT>"	'Descriptor:Always

	Response.Write "<INPUT NAME=""summary"" TYPE=""radio"" VALUE=""2"" "
	If Clng(sSelectedChoice) = SITE_PROPVALUE_SUMMARY_PAGE_WHENMORETHANONEQO Then
		Response.Write "CHECKED=""1"" "
	End If
	Response.Write "/><FONT FACE=" & aFontInfo(S_FAMILY_FONT) & " SIZE=" & aFontInfo(N_SMALL_FONT) & ">" & "only when there are more than 1 question" & "</FONT>"	'Descriptor: XXXXXX asDescriptors(836)

	Response.Write "<INPUT NAME=""summary"" TYPE=""radio"" VALUE=""3"" "
	If Clng(sSelectedChoice) = SITE_PROPVALUE_SUMMARY_PAGE_NEVER Then
		Response.Write "CHECKED=""1"" "
	End If
	Response.Write "/><FONT FACE=" & aFontInfo(S_FAMILY_FONT) & " SIZE=" & aFontInfo(N_SMALL_FONT) & ">" & asDescriptors(36) & "</FONT>"	'Descriptor:Never

    RenderSummaryPageChoices = lErrNumber
	Err.Clear
End Function
%>