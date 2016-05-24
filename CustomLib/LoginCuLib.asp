<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/LoginCoLib.asp" -->
<!--#include file="../CustomLib/DeviceTypesCuLib.asp" -->
<%
	Function ParseRequestForLogin(oRequest, sUserName, sPassword, sSavePwd, sStatus, sNTUser)
	'********************************************************
	'*Purpose:
	'*Inputs: oRequest
	'*Outputs: sUserName, sPassword, sSavePwd, sReqLogin
	'********************************************************
		On Error Resume Next
        Dim lErrNumber

        lErrNumber = NO_ERR

		sUserName = ""
		sPassword = ""
		sSavePwd = ""
		sStatus = ""

		'sUserName = Server.HTMLEncode(Trim(CStr(oRequest("userName"))))
		'sPassword = Server.HTMLEncode(Trim(CStr(oRequest("Pwd"))))
		sUserName = Trim(CStr(oRequest("userName")))
		sPassword = Trim(CStr(oRequest("Pwd")))
		sSavePwd = Trim(CStr(oRequest("SavePwd")))
		sStatus = Trim(CStr(oRequest("status")))
		sNTUser = Trim(CStr(oRequest("NTUser")))

        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "ParseRequestForLogin", "", "Error setting variables equal to Request variables", LogLevelError)
        End If

		ParseRequestForLogin = lErrNumber
		Err.Clear
	End Function

	Function ParseRequestForNewUser(oRequest, sUserName, sPassword, sConfirmPassword, sHint, sDefEmail, sLocaleID, sLanguageID, sSavePwd)
	'********************************************************
	'*Purpose:
	'*Inputs: oRequest
	'*Outputs: sUserID, sUserName, sPassword, sConfirmPassword, sHint
	'********************************************************
		On Error Resume Next
		Dim sLocaleData
		Dim iDelimiter
        Dim lErrNumber

        lErrNumber = NO_ERR

		sUserName = ""
		sPassword = ""
		sConfirmPassword = ""
		sHint = ""
		sDefEmail = ""
		sLocaleID = ""
		sLanguageID = ""
		sSavePwd = ""

		sUserName = Trim(CStr(oRequest("userName")))
		sPassword = Trim(CStr(oRequest("Pwd")))
		sConfirmPassword = Trim(CStr(oRequest("confirmPwd")))
		sHint = Trim(CStr(oRequest("Hint")))
		sDefEmail = Trim(CStr(oRequest("defEmail")))
		sLocaleData = Trim(CStr(oRequest("Locale")))
		If sLocaleData <> "" Then
		    iDelimiter = Instr(1, sLocaleData, ";")
		    sLocaleID = Left(sLocaleData, iDelimiter - 1)
		    sLanguageID = Mid(sLocaleData, iDelimiter + 1)
		End If
		sSavePwd = Trim(CStr(oRequest("SavePwd")))

        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "ParseRequestForNewUser", "", "Error setting variables equal to Request variables", LogLevelError)
        End If

		ParseRequestForNewUser = lErrNumber
		Err.Clear
	End Function

	Function ParseRequestForPasswordHint(oRequest, sUserName)
	'********************************************************
	'*Purpose:
	'*Inputs: oRequest
	'*Outputs: sUserName, sReqPwdHint
	'********************************************************
		On Error Resume Next
		Dim lErrNumber

		lErrNumber = NO_ERR

		sUserName = ""

		sUserName = Server.HTMLEncode(Trim(CStr(oRequest("userName"))))

		If Err.number <> NO_ERR Then
		    lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "ParseRequestForPasswordHint", "", "Error setting variables equal to Request variables", LogLevelError)
		End If

		ParseRequestForPasswordHint = lErrNumber
		Err.Clear
	End Function

	Function ParseRequestForDeactivateUser(oRequest, sPassword)
	'********************************************************
	'*Purpose:
	'*Inputs: oRequest
	'*Outputs: sPassword, sReqDeact
	'********************************************************
	    On Error Resume Next
	    Dim lErrNumber

	    lErrNumber = NO_ERR

	    sPassword = ""

	    sPassword = Server.HTMLEncode(Trim(CStr(oRequest("deactPassword"))))

	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "ParseRequestForDeactivateUser", "", "Error setting variables equal to Request variables", LogLevelError)
	    End If

	    ParseRequestForDeactivateUser = lErrNumber
	    Err.Clear
	End Function

Function validate_CreateUser(oRequest, sUserName, sPassword, sConfirmPassword, sHint)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim Item
    Dim lErrNumber
    Dim tempError
    lErrNumber = NO_ERR
    tempError = NO_ERR

    If Len(sUserName) = 0 Or Len(sPassword) = 0 Then
		lErrNumber = lErrNumber + ERR_LOGIN_BLANKS
	End If

    If Len(sHint) = 0 Then
	    lErrNumber = lErrNumber + ERR_HINT_BLANK
	End If

    If sPassword <> sConfirmPassword Then
		lErrNumber = lErrNumber + ERR_CONFIRM_PASSWORD
	End If

    If Len(CStr(oRequest("defEmail"))) > 0 Then
        Select Case CStr(Application("Device_Validation"))
            Case S_DEVICE_VALIDATION_EMAIL
                tempError = ValidateEmailAddress(oRequest("defEmail"))
            Case S_DEVICE_VALIDATION_NUMBER
                tempError = ValidateNumberAddress(oRequest("defEmail"))
            Case S_DEVICE_VALIDATION_NONE
                'Do nothing
            Case Else
        End Select
        If tempError = -1 Then
            lErrNumber = lErrNumber + ERR_DEFAULT_ADDRESS_INVALID
        End If
    End If

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

    validate_CreateUser = lErrNumber
    Err.Clear
End Function

	Function SetSessionInfo(sChannel, sSessionID, sUserName, sSavePwd, sLocaleID, sLanguageID)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'*TO DO: Add error handling
	'*QUESTION: should current site be set in here?
	'********************************************************
		On Error Resume Next
		Dim lErrNumber
		Dim oPropertiesDOM
		Dim oLocalesDOM
		Dim dUserCookieExpiration
		Dim sSavePasswordFlag
		Dim sGetUserPropertiesXML
		Dim sGetLocalesForSiteXML
		Dim sEncryptedPassword
		Dim sSiteLocale
		Dim sUserLocaleData
		Dim iDelimiter

		lErrNumber = NO_ERR

		If sSavePwd = "1" Then
		    dUserCookieExpiration = COOKIES_EXPIRATION_DATE
		    sSavePasswordFlag = "1"
		Else
		    dUserCookieExpiration = ""
		    sSavePasswordFlag = "0"
        End If

		If Len(sChannel) > 0 Then
			aSourceInfo(0) = SITE_COOKIE
		    aSourceInfo(1) = CURRENT_SITE
		    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
		    Call WriteToSource(aConnectionInfo, CStr(sChannel), SOURCE_COOKIES, aSourceInfo)
			If Err.number <> NO_ERR Then
				lErrNumber = Err.number
				Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetSessionInfo", "", "Error calling WriteToSource for CURRENT_SITE", LogLevelError)
			End If
		End If

		If lErrNumber = NO_ERR Then
		    aSourceInfo(0) = USER_COOKIE
		    aSourceInfo(1) = "uname"
		    aSourceInfo(2) = dUserCookieExpiration
		    Call WriteToSource(aConnectionInfo, sUserName, SOURCE_COOKIES, aSourceInfo)
		    If Err.number <> NO_ERR Then
                lErrNumber = Err.number
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetSessionInfo", "", "Error calling WriteToSource for uname", LogLevelError)
		    End If
		End If

        If lErrNumber = NO_ERR Then
		    lErrNumber = cu_GetUserProperties(sGetUserPropertiesXML)
		    If lErrNumber <> NO_ERR Then
		        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetSessionInfo", "", "Error calling cu_GetUserProperties", LogLevelTrace)
		    Else
                lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetUserPropertiesXML, oPropertiesDOM)
                If lErrNumber <> NO_ERR Then
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetSessionInfo", "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
                Else
					sEncryptedPassword = CStr(oPropertiesDOM.selectSingleNode("mi/prs/pr[@n = 'MR_PASSWORD']").getAttribute("v"))
					If Len(sLocaleID) = 0 Then
						sLocaleID = CStr(oPropertiesDOM.selectSingleNode("mi/prs/pr[@n = 'MR_LOCALE_ID']").getAttribute("v"))
					'	sUserLocaleData = CStr(oPropertiesDOM.selectSingleNode("mi/prs/pr[@n = 'MR_LOCALE_ID']").getAttribute("v"))
					'	iDelimiter = Instr(1, sUserLocaleData, ";")
					'	sLocaleID = Left(sUserLocaleData, iDelimiter - 1)
					'	sLanguageID = Mid(sUserLocaleData, iDelimiter + 1)
					End If
				End If
    	    End If
		End If

        If lErrNumber = NO_ERR Then
    		If Len(sEncryptedPassword) > 0 Then
			    aSourceInfo(0) = USER_COOKIE
			    aSourceInfo(1) = USER_PASSWORD
			    aSourceInfo(2) = dUserCookieExpiration
			    Call WriteToSource(aConnectionInfo, CStr(sEncryptedPassword), SOURCE_COOKIES, aSourceInfo)
			    If Err.number <> NO_ERR Then
			    	lErrNumber = Err.number
			    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetSessionInfo", "", "Error calling WriteToSource for USER_PASSWORD", LogLevelError)
			    End If
			End If

			aSourceInfo(0) = USER_COOKIE
		    aSourceInfo(1) = SAVE_PASSWORD
		    aSourceInfo(2) = dUserCookieExpiration
		    Call WriteToSource(aConnectionInfo, sSavePasswordFlag, SOURCE_COOKIES, aSourceInfo)
		    If Err.number <> NO_ERR Then
		    	lErrNumber = Err.number
		    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetSessionInfo", "", "Error calling WriteToSource for SAVE_PASSWORD", LogLevelError)
		    End If
		End If

		If lErrNumber = NO_ERR Then
			If (StrComp(sLocaleID, CStr(GetSiteLocale()), vbBinaryCompare) = 0) And (Len(CStr(GetLanguageSetting())) > 0) Then
			    'assume the language cookie is fine and do nothing
			Else
		        Call SetSiteLocale(sLocaleID)

				lErrNumber = cu_GetLocalesForSite(sGetLocalesForSiteXML)
				lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetLocalesForSiteXML, oLocalesDOM)
				sLanguageID = CStr(oLocalesDOM.selectSingleNode("/mi/in/oi[@id = '" & sLocaleID & "']").getAttribute("plid"))

			    aSourceInfo(0) = SITE_COOKIE
			    aSourceInfo(1) = "Lng"
			    aSourceInfo(2) = COOKIES_EXPIRATION_DATE
			    Call WriteToSource(aConnectionInfo, CStr(sLanguageID), SOURCE_COOKIES, aSourceInfo)
			    If Err.number <> NO_ERR Then
			    	lErrNumber = Err.number
			    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetSessionInfo", "", "Error calling WriteToSource for Lng", LogLevelError)
			    End If
			End If
		End If

        If lErrNumber = NO_ERR Then
            If Len(GetJavaScriptPreference()) = 0 Then
                aSourceInfo(0) = SITE_COOKIE
                aSourceInfo(1) = PREF_JAVASCRIPT
                aSourceInfo(2) = COOKIES_EXPIRATION_DATE
                Call WriteToSource(aConnectionInfo, Application("Default_use_dhtml"), SOURCE_COOKIES, aSourceInfo)
		    	If Err.number <> NO_ERR Then
                    lErrNumber = Err.number
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetSessionInfo", "", "Error calling WriteToSource for PREF_JAVASCRIPT", LogLevelError)
		    	End If
            End If
        End If

	    If lErrNumber = NO_ERR Then
		    If Len(GetJavaScriptSetting()) = 0 Then
		        Dim sJSPref
		        sJSPref = GetJavaScriptPreference()
    		    Select Case sJSPref
			        Case "2"
				        aSourceInfo(0) = SITE_COOKIE
		                aSourceInfo(1) = USE_JAVASCRIPT
		                aSourceInfo(2) = COOKIES_EXPIRATION_DATE
		                Call WriteToSource(aConnectionInfo, CStr(SupportsJavaScript()), SOURCE_COOKIES, aSourceInfo)
			        Case Else
				        aSourceInfo(0) = SITE_COOKIE
		                aSourceInfo(1) = USE_JAVASCRIPT
		                aSourceInfo(2) = COOKIES_EXPIRATION_DATE
		                Call WriteToSource(aConnectionInfo, CStr(sJSPref), SOURCE_COOKIES, aSourceInfo)
			    End Select
		    	If Err.number <> NO_ERR Then
                    lErrNumber = Err.number
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetSessionInfo", "", "Error calling WriteToSource for USE_JAVASCRIPT", LogLevelError)
		    	End If
		    End If
		End If

		Set oPropertiesDOM = Nothing
		Set oLocalesDOM = Nothing

		SetSessionInfo = lErrNumber
		Err.Clear
	End Function

Function SetPortalAddress(sPortalID)
'********************************************************
'*Purpose:
'*Inputs: sPortalID
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
    Dim dCookieExpiration

    lErrNumber = NO_ERR

    If GetSavePasswordSetting() = "1" Then
        dCookieExpiration = COOKIES_EXPIRATION_DATE
    Else
        dCookieExpiration = ""
    End If

	aSourceInfo(0) = USER_COOKIE
	aSourceInfo(1) = PORTAL_ADDRESS
	aSourceInfo(2) = dCookieExpiration
	Call WriteToSource(aConnectionInfo, CStr(sPortalID), SOURCE_COOKIES, aSourceInfo)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "SetPortalAddress", "", "Error calling WriteToSource for PORTAL_ADDRESS", LogLevelError)
	End If

	SetPortalAddress = lErrNumber
	Err.Clear
End Function

Function cu_AddPortalAddress(sUserName)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_AddPortalAddress"
	Dim lErrNumber
	Dim sDevice
	Dim asAddressProperties()
	Redim asAddressProperties(MAX_ADDR_PROP)
	Dim sSessionID
	Dim sPortalAddressXML
	Dim sPortalID
	Dim bGenerateTransProps

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
    sDevice = Application("Portal_Device")
    bGenerateTransProps = True

	If (lErrNumber = NO_ERR) And (Len(sDevice) > 0) Then
	    sPortalID = ""
	    sPortalID = GetGUID()

	    asAddressProperties(ADDR_PROP_ADDRESS_ID) = sPortalID							'addressID
		asAddressProperties(ADDR_PROP_ADDRESS_NAME) = sUserName & " Web Page"						'addressName
		asAddressProperties(ADDR_PROP_PHYSICAL_ADDRESS) = sUserName & " - Web - My Web Page"   'physicalAddress
	    asAddressProperties(ADDR_PROP_ADDRESS_DISPLAY) = sUserName & " Web Page"						'addressDisplay
	    asAddressProperties(ADDR_PROP_DEVICE_ID) = sDevice							'deviceID
		asAddressProperties(ADDR_PROP_DELIVERY_WINDOW) = ""						'deliveryWindow
	    asAddressProperties(ADDR_PROP_TIMEZONE_ID) = GetDefaultTimeZone()								'DefaultTimezoneStdName
	    asAddressProperties(ADDR_PROP_STATUS) = "1"								'status
		asAddressProperties(ADDR_PROP_CREATED_BY) = ""								'createdBy
		asAddressProperties(ADDR_PROP_LAST_MODIFIED_BY) = ""								'lastModBy
	    asAddressProperties(ADDR_PROP_TRANSMISSION_PROPERTIES_ID) = "" 'transPropsID
	    asAddressProperties(ADDR_PROP_PIN) = ""								'PIN
	    asAddressProperties(ADDR_PROP_EXPIRATION_DATE) = ""		'expirationDate
	    asAddressProperties(ADDR_PROP_CREATED_DATE) = ""		'createdDate
	    asAddressProperties(ADDR_PROP_LAST_MODIFIED_DATE) = ""		'lastModDate

	    lErrNumber = co_AddAddress(sSessionID, asAddressProperties, bGenerateTransProps, sPortalAddressXML)
	    If lErrNumber <> NO_ERR Then
	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_AddAddress", LogLevelTrace)
	    End If
	End If

    If (lErrNumber = NO_ERR) And (sDevice <> "") Then
        Call SetPortalAddress(sPortalID)
    End If

	cu_AddPortalAddress = lErrNumber
	Err.Clear
End Function

Function AddDefaultAddress(sUserName, sDefAddID, sDefEmail)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "AddDefaultAddress"
	Dim lErrNumber
	Dim asAddressProperties()
	Redim asAddressProperties(MAX_ADDR_PROP)
	Dim sSessionID
	Dim sDefaultAddressXML
	Dim bGenerateTransProps

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
	bGenerateTransProps = True

	asAddressProperties(ADDR_PROP_ADDRESS_ID) = sDefAddID      'addressID
	asAddressProperties(ADDR_PROP_ADDRESS_NAME) = sDefEmail       'addressName
	asAddressProperties(ADDR_PROP_PHYSICAL_ADDRESS) = sDefEmail      'physicalAddress
	asAddressProperties(ADDR_PROP_ADDRESS_DISPLAY) = sDefEmail		'addressDisplay
	asAddressProperties(ADDR_PROP_DEVICE_ID) = Application("Default_Device")     'deviceID
	asAddressProperties(ADDR_PROP_DELIVERY_WINDOW) = ""						'deliveryWindow
	asAddressProperties(ADDR_PROP_TIMEZONE_ID) = GetDefaultTimeZone()								'DefaultTimezoneStdName
	asAddressProperties(ADDR_PROP_STATUS) = "1"								'status
	asAddressProperties(ADDR_PROP_CREATED_BY) = ""								'createdBy
	asAddressProperties(ADDR_PROP_LAST_MODIFIED_BY) = ""								'lastModBy
	asAddressProperties(ADDR_PROP_TRANSMISSION_PROPERTIES_ID) = "" 'transPropsID
	asAddressProperties(ADDR_PROP_PIN) = ""								'PIN
	asAddressProperties(ADDR_PROP_EXPIRATION_DATE) = ""		'expirationDate
	asAddressProperties(ADDR_PROP_CREATED_DATE) = ""		'createdDate
	asAddressProperties(ADDR_PROP_LAST_MODIFIED_DATE) = ""		'lastModDate

	lErrNumber = co_AddAddress(sSessionID, asAddressProperties, bGenerateTransProps, sDefaultAddressXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_AddAddress", LogLevelTrace)
	End If

	AddDefaultAddress = lErrNumber
	Err.Clear
End Function

Function CheckForPortalAddress(sGetUserAddressesXML, sUserName)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oOutputDOM
	Dim sPortalDeviceID
	Dim oPortalAddress
	Dim sPortalAddressId

	lErrNumber = NO_ERR

	sPortalDeviceID = CStr(Application("Portal_Device"))
	If Len(sPortalDeviceID) = 0 Then
	    'If not portal device, the user cannot have a portal address, set it to blank:
	    Call SetPortalAddress("")

	Else
        'Search for the portal device in the user addresses:
	    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetUserAddressesXML, oOutputDOM)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "CheckForPortalAddress", "", "Error loading sGetUserAddressesXML", LogLevelTrace)
	    Else

            Set oPortalAddress = oOutputDOM.selectSingleNode("//oi[@dvid = '" & sPortalDeviceID & "']")
	        If Not oPortalAddress Is Nothing Then
	            'If found, save it on cache:
	        	Call SetPortalAddress(oPortalAddress.getAttribute("id"))
	        Else
	            'If not found, create a new address:
                lErrNumber = cu_AddPortalAddress(sUserName)
                If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "CheckForPortalAddress", "", "Error loading cu_AddPortalAddress", LogLevelTrace)
	        End If

	    End If

	End If

	Set oOutputDOM = Nothing
	Set oPortalAddress = Nothing

	CheckForPortalAddress = lErrNumber
	Err.Clear

End Function

'***NEW FUNCTIONS

Function cu_CreateSession(sUserName, sPassword, bEncryptedFlag, sSessionID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_CreateSession"
	Dim lErrNumber
	Dim sSiteID
	Dim sCreateSessionXML
	Dim oOutputDOM
	Dim oError

	lErrNumber = NO_ERR
	sSiteID = SITE_ID
	sSessionID = ""

	If Len(sUserName) = 0 Or Len(sPassword) = 0 Then
		lErrNumber = ERR_LOGIN_BLANKS
	End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CreateSession(sSiteID, sUserName, sPassword, bEncryptedFlag, sCreateSessionXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateSession", LogLevelTrace)
        End If
    End If
    If lErrNumber = NO_ERR Then
	    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCreateSessionXML, oOutputDOM)
        Set oError = oOutputDOM.selectSingleNode("/mi/er")
        sSessionID = CStr(oError.getAttribute("des"))
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error retrieving er node", LogLevelError)
        End If
    End If

    Set oOutputDOM = Nothing
	Set oError = Nothing

    cu_CreateSession = lErrNumber
	Err.Clear
End Function

Function cu_CreateUser(sUsername, sPassword, sHint, sLocaleID, sLanguageID, sDefAddID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_CreateUser"
	Dim asUserProperties()
	Redim asUserProperties(MAX_USER_PROP)
	Dim lErrNumber
    Dim sSiteID
    Dim sExpireSetting
    Dim sExpireDate
    Dim sUserID

	lErrNumber = NO_ERR
    sSiteID = SITE_ID
    sUserID = GetGUID()
    If Len(sUserID) = 0 then
    	cu_CreateUser = ERR_EMPTY_GUID
    	Exit Function
    End If

    sDefAddID = GetGUID()
    If Len(sUserID) = 0 then
	    cu_CreateUser = ERR_EMPTY_GUID
    	Exit Function
    End If

    sExpireSetting = Application("Default_expire_value")

    'Convert expiration date to format: yyyy-mm-dd
    If IsDate(sExpireSetting) Then
        sExpireSetting = CDate(sExpireSetting)
    Else
        'Assume it is number of days from now
        sExpireSetting = DateAdd("d", sExpireSetting, Now)
    End If
    sExpireDate = Year(sExpireSetting) & "-"
    If Month(sExpireSetting) < 10 Then
        sExpireDate = sExpireDate & "0" & Month(sExpireSetting) & "-"
    Else
        sExpireDate = sExpireDate & Month(sExpireSetting) & "-"
    End If
    If Day(sExpireSetting) < 10 Then
        sExpireDate = sExpireDate & "0" & Day(sExpireSetting)
    Else
        sExpireDate = sExpireDate & Day(sExpireSetting)
    End If

	asUserProperties(USER_PROP_USER_ID) = sUserID         'UserID
	asUserProperties(USER_PROP_USER_NAME) = sUsername         'Username
	asUserProperties(USER_PROP_PASSWORD) = sPassword         'Password
	asUserProperties(USER_PROP_HINT) = sHint         'Hint
	asUserProperties(USER_PROP_LOCALE_ID) = sLocaleID '& ";" & sLanguageID      'LocaleID
	asUserProperties(USER_PROP_DEFAULT_ADDRESS_ID) = sDefAddID  'Default AddressID
	asUserProperties(USER_PROP_AGREEMENT_ID) = ""         'AgreementID
	asUserProperties(USER_PROP_ACCOUNT_ID) = ""         'AccountID
	asUserProperties(USER_PROP_STATUS) = "1"         'Status
	asUserProperties(USER_PROP_EXPIRATION_DATE) = sExpireDate       'Expiration Date
	asUserProperties(USER_PROP_CREATED_DATE) = ""         'Created Date
	asUserProperties(USER_PROP_CREATED_BY) = ""         'Created By
	asUserProperties(USER_PROP_LAST_MODIFIED_DATE) = ""         'Modification Date
	asUserProperties(USER_PROP_LAST_MODIFIED_BY) = ""         'Last Modified By

	lErrNumber = co_CreateUser(sSiteID, asUserProperties)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateUser", LogLevelTrace)
	End If

	cu_CreateUser = lErrNumber
	Err.Clear
End Function

Function cu_GetUserHint(sUsername, sPasswordHint)
'********************************************************
'*Purpose:
'*Inputs: sUsername
'*Outputs: sPasswordHint
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetUserHint"
	Dim lErrNumber
	Dim sSiteID
	Dim sGetUserHintXML
	Dim oOutputDOM
	Dim oError

	lErrNumber = NO_ERR
	sSiteID = SITE_ID
    sPasswordHint = ""

    If Len(sUsername) = 0 Then
        lErrNumber = ERR_LOGIN_BLANKS
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_GetUserHint(sSiteID, sUsername, sGetUserHintXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetUserHint", LogLevelTrace)
        End If
    End If

    If lErrNumber = NO_ERR Then
	    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetUserHintXML, oOutputDOM)
        Set oError = oOutputDOM.selectSingleNode("/mi/er")
        sPasswordHint = CStr(oError.getAttribute("des"))
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error retrieving er node", LogLevelError)
        End If
    End If

    Set oOutputDOM = Nothing
    Set oError = Nothing

	cu_GetUserHint = lErrNumber
	Err.Clear
End Function

Function cu_DeactivateUser(sPassword)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_DeactivateUser"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    If Len(sPassword) = 0 Then
        lErrNumber = ERR_LOGIN_BLANKS
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_DeactivateUser(sSessionID, sPassword)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_DeactivateUser", LogLevelTrace)
        End If
    End If

	cu_DeactivateUser = lErrNumber
	Err.Clear
End Function

Function cu_DeleteUser()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_DeleteUser"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    lErrNumber = co_DeleteUser(sSessionID)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_DeleteUser", LogLevelTrace)
    End If

	cu_DeleteUser = lErrNumber
	Err.Clear
End Function

Function cu_CloseSession()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CloseSession(sSessionID)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", "cu_CloseSession", "", "Error calling co_CloseSession", LogLevelTrace)
        End If
    End If

	cu_CloseSession = lErrNumber
	Err.Clear
End Function

Function cu_GetUserProperties(sGetUserPropertiesXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetUserProperties"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    lErrNumber = co_GetUserProperties(sSessionID, sGetUserPropertiesXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetUserProperties", LogLevelTrace)
    End If

	cu_GetUserProperties = lErrNumber
	Err.Clear
End Function

Function cu_SaveUserAuthenticationObjects(oRequest)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Get Castor XML structure
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_SaveUserAuthenticationObjects"
	Dim lErrNumber
	Dim sSessionID
	Dim asInformationSourceID()
	Dim asXMLAuthenticationString()
	Dim iArrayCounter
	Dim Item
	Dim sServerName
	Dim sProjectID
	Dim lPort
	Dim sUserName
	Dim sPwd
	Dim sUserID
	Dim iArraySize
	Dim sISID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
	iArrayCounter = 0

	iArraySize = 0
    For Each Item in oRequest
		If (Left(Item, 8) = "AO_User_") Then
            If Len(CStr(oRequest(Item))) > 0 Then
                sISID = CStr(Right(Item, Len(Item) - 8))
                IF (oRequest("AO_Req_" & sISIS) = "1") Then
					iArraySize = iArraySize + 1
				End IF
			End IF
		End IF
    Next

	Redim asInformationSourceID(iArraySize-1)
    Redim asXMLAuthenticationString(iArraySize-1)

    For Each Item in oRequest
		If (Left(Item, 8) = "AO_User_") Then
            If Len(CStr(oRequest(Item))) > 0 Then
                sISID = CStr(Right(Item, Len(Item) - 8))
                IF (oRequest("AO_Req_" & sISIS) = "1") Then
					sServerName = CStr(oRequest("AO_Serv_" & sISID))
					sProjectID = CStr(oRequest("AO_Proj_" & sISID))
					lPort = CLng(oRequest("AO_Port_" & sISID))
					sUserName = CStr(oRequest(Item))
					sPwd = CStr(oRequest("AO_Password_" & sISID))
					sUserID = ""

					If (Len(sServerName) > 0) And (Len(sProjectID) > 0) And (Len(CStr(lPort)) > 0) Then
					    lErrNumber = UserAuthenticate(sServerName, sProjectID, lPort, sUserName, sPwd, sUserID)
					    If lErrNumber <> NO_ERR Then
					        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling UserAuthenticate", LogLevelTrace)
					        Exit For
					    ElseIf Len(sUserID) = 0 Then
					        lErrNumber = ERR_ISLOGIN_ERROR
					        Exit For
					    End If
					End If

					asInformationSourceID(iArrayCounter) = sISID
					Call BuildAuthenticationObject(sUserName, sPwd, sUserID, asXMLAuthenticationString(iArrayCounter))
					iArrayCounter = iArrayCounter + 1
					If iArrayCounter = iArraySize Then
					    Exit For
					End If
				End If
            End If
        End If
    Next

    If (lErrNumber = NO_ERR) And (iArrayCounter > 0) Then
        lErrNumber = co_SaveUserAuthenticationObjects(sSessionID, asInformationSourceID, asXMLAuthenticationString)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SaveUserAuthenticationObjects", LogLevelTrace)
        End If
    End If

	cu_SaveUserAuthenticationObjects = lErrNumber
	Err.Clear
End Function

Function cu_GetInformationSourcesForSite(sGetInformationSourcesForSiteXML, bHasProjects)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetInformationSourcesForSite"
	Dim lErrNumber
	Dim sSiteID
	Dim oOutputDOM
	Dim oNodes

	lErrNumber = NO_ERR
	sSiteID = SITE_ID
	bHasProjects = False

    lErrNumber = co_GetInformationSourcesForSite(sSiteID, sGetInformationSourcesForSiteXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetInformationSourcesForSite", LogLevelTrace)
    End If

    If lErrNumber = NO_ERR Then
		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetInformationSourcesForSiteXML, oOutputDOM)
        Set oNodes = oOutputDOM.selectNodes("/mi/in/oi")
        If oNodes.length > 0 Then
            bHasProjects = True
        End If
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error retrieving er node", LogLevelError)
        End If
    End If

    Set oOutputDOM = Nothing
    Set oNodes = Nothing

	cu_GetInformationSourcesForSite = lErrNumber
	Err.Clear
End Function

Function cu_UpdateUserProperties(sHint, sLocaleID, sDefAddID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_UpdateUserProperties"
	Dim lErrNumber
	Dim sSessionID
	Dim asUserProperties()
	Redim asUserProperties(MAX_USER_PROP)

	lErrNumber = NO_ERR
    sSessionID = GetSessionID()

	asUserProperties(USER_PROP_USER_ID) = ""         'UserID
	asUserProperties(USER_PROP_USER_NAME) = GetUsername()         'Username
	asUserProperties(USER_PROP_PASSWORD) = ""         'Password
	asUserProperties(USER_PROP_HINT) = sHint         'Hint
	asUserProperties(USER_PROP_LOCALE_ID) = sLocaleID         'LocaleID
	asUserProperties(USER_PROP_DEFAULT_ADDRESS_ID) = sDefAddID  'Default AddressID
	asUserProperties(USER_PROP_AGREEMENT_ID) = ""         'AgreementID
	asUserProperties(USER_PROP_ACCOUNT_ID) = ""         'AccountID
	asUserProperties(USER_PROP_STATUS) = ""         'Status
	asUserProperties(USER_PROP_EXPIRATION_DATE) = ""         'Expiration Date
	asUserProperties(USER_PROP_CREATED_DATE) = ""         'Created Date
	asUserProperties(USER_PROP_CREATED_BY) = ""         'Created By
	asUserProperties(USER_PROP_LAST_MODIFIED_DATE) = ""         'Modification Date
	asUserProperties(USER_PROP_LAST_MODIFIED_BY) = ""         'Last Modified By

    If lErrNumber = NO_ERR Then
        lErrNumber = co_UpdateUserProperties(sSessionID, asUserProperties)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_UpdateUserProperties", LogLevelTrace)
        End If
    End If

	cu_UpdateUserProperties = lErrNumber
	Err.Clear
End Function

Function SiteHasLocales(sGetLocalesForSiteXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
Const PROCEDURE_NAME = "SiteHasLocales"
Dim lErrNumber
Dim oOutputDOM
Dim oLocales

    On Error Resume Next
    lErrNumber = NO_ERR

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetLocalesForSiteXML, oOutputDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "LoginCuLib.asp", PROCEDURE_NAME, "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
		SiteHasLocales = False
	Else
        Set oLocales = oOutputDOM.selectNodes("/mi/in/oi[@tp='" & TYPE_LOCALE & "' $and$ @plid!='']")
        SiteHasLocales = oLocales.length > 0
	End If

    Set oOutputDOM = Nothing
    Set oLocale = Nothing

    Err.Clear

End Function



Function DefaultPage()
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
Const PROCEDURE_NAME = "DefaultPage"
Dim sDefaultPage
Dim oChannels
Dim oChannelsDOM
Dim sChannelsXML
Dim lErr

	On Error Resume Next
	lErr = NO_ERR

	sDefaultPage = "default.asp"

    lErr = cu_GetChannels(sChannelsXML)
    If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_UpdateUserProperties", LogLevelTrace)

    If lErr = NO_ERR Then
	    lErr = LoadXMLDOMFromString(aConnectionInfo, sChannelsXML, oChannelsDOM)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_UpdateUserProperties", LogLevelTrace)
        Else
            Set oChannels = oChannelsDOM.selectNodes("//oi[prs/pr[@id='active' and @v='1']]")

            If oChannels.length = 1 Then
                sDefaultPage = GetStartPage() & ".asp?start=0&site=" & oChannels.item(0).getAttribute("id")
            End If
        End If
    End If

    DefaultPage = sDefaultPage
    Err.Clear

End Function

Function ProcessCreateNewUser(sNCSUserName,sCastorUserName, sCastorUserID, sCastorPassword, sAuthMode, sSessionID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
Const PROCEDURE_NAME = "ProcessCreateNewUser"

Dim sLanguage
Dim lErrNumber
Dim sGetLocalesForSiteXML
Dim sLocaleID
Dim sDefAddID
Dim oLocalesDOM
Dim oOutputDOM
Dim oNodes
Dim oNode
Dim sISourceArray()
Dim sAuthObjectArray()
Dim i,j,count
Dim aInformationSources
Dim asServerNames


	On Error Resume Next
	lErrNumber = NO_ERR
	sLocaleID = ""
	i = 0

	sLanguage = GetLng()

	lErrNumber = cu_GetLocalesForSite(sGetLocalesForSiteXML)

	If lErrNumber = NO_ERR Then
		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetLocalesForSiteXML, oLocalesDOM)
		sLocaleID = CStr(oLocalesDOM.selectSingleNode("/mi/in/oi[@plid = '" & sLanguage & "']").getAttribute("id"))
		If sLocaleID = "" Then
			sLocaleID = "FBBF7C1E37EC11D4887C00C04F48F8FD"   'Default to System locale ID
		End If
	End If

	If lErrNumber = NO_ERR Then
		lErrNumber = cu_CreateUser(sNCSUserName, sCastorUserID, sNCSUserName, sLocaleID, sLanguage, sDefAddID)
	End If

	If lErrNumber = NO_ERR Then
		lErrNumber = cu_CreateSession(sNCSUserName, sCastorUserID, bEncryptedFlag, sSessionID)
	End If


    'Get IS List:
    If lErrNumber = NO_ERR Then
        lErrNumber = getInfSourcesServerNames(aInformationSources)
    End If

	asServerNames = GetClusterNodeNames()

	Count = 0
	For i=0 To UBound(aInformationSources)
		If isIServerInCluster(aInformationSources(i,1), asServerNames) Then
			Count = count + 1
		End If
	Next

	If Count > 0 Then

		Redim sISourceArray(count-1)
		Redim sAuthObjectArray(count-1)

		j=0
		If lErrNumber = NO_ERR Then
			For i=0 To UBound(aInformationSources)
				If isIServerInCluster(aInformationSources(i,1), asServerNames) Then
					sISourceArray(j) = aInformationSources(i,0)
					sAuthObjectArray(j) = "AuthUserName=""" + sCastorUserName + """ AuthUserID=""" + sCastorUserID + """ AuthUserPwd=""" + Encrypt(sCastorPassword) + """ AuthMode=""" + sAuthMode + """"
					j = j + 1
				End If
			Next

		    If Err.number <> NO_ERR Then
		        lErrNumber = Err.number
		        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error retrieving information sources", LogLevelError)
		    End If
		End If

		If lErrNumber = NO_ERR Then
		    lErrNumber = co_SaveUserAuthenticationObjects(sSessionID, sISourceArray, sAuthObjectArray)
		    If lErrNumber <> NO_ERR Then
		        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SaveUserAuthenticationObjects", LogLevelTrace)
		    End If
		End If

	End IF

    Set oOutputDOM = Nothing
    Set oNodes = Nothing
    Set oNode = Nothing

    ProcessCreateNewUser = lErrNumber
    Err.Clear

End Function


Function UpdateAuthenticationObject(sSessionID, sCastorUserName, sCastorUserID, sCastorPassword, sAuthMode)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
Const PROCEDURE_NAME = "UpdateAuthenticationObject"

Dim lErrNumber
Dim sISourceArray()
Dim sAuthObjectArray()
Dim i,j,count
Dim aInformationSources
Dim sUserAuthenticationObjectsXML
Dim oUserAuthDOM
Dim oUserAuth
Dim sAuthName
Dim sAuthPwd
Dim sAuthID
Dim asServerNames

	On Error Resume Next

    'Get IS List:
    If lErrNumber = NO_ERR Then
        lErrNumber = getInfSourcesServerNames(aInformationSources)
    End If

	lErr = co_GetUserAuthenticationObjects(sSessionID, sUserAuthenticationObjectsXML)
	If lErr = NO_ERR Then
		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sUserAuthenticationObjectsXML, oUserAuthDOM)
	End If

	asServerNames = GetClusterNodeNames()

	Count = 0
	For i=0 To UBound(aInformationSources)
		'If Strcomp(aInformationSources(i,1), getIServerName(), vbTextCompare) = 0 Then
		If isIServerInCluster(aInformationSources(i,1), asServerNames) Then
			Set oUserAuth = oUserAuthDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_INFORMATION_SOURCE & "' $and$ @id='" & aInformationSources(i,0) & "']")
			If Not oUserAuth is Nothing Then
				sAuthID = ""
				lErrNumber = ParseAuthenticationObject(oUserAuth.getAttribute("v"), sAuthName, sAuthPwd, sAuthID)
				If sAuthID = sCastorUserID Then
					Count = Count + 1
				End If
			End If
		End If
	Next


	If Count > 0 Then

		Redim sISourceArray(count-1)
		Redim sAuthObjectArray(count-1)

		j=0
		If lErrNumber = NO_ERR Then
			For i=0 To UBound(aInformationSources)
				'If StrComp(aInformationSources(i,1), getIServerName(), vbTextCompare) = 0  Then
				If isIServerInCluster(aInformationSources(i,1), asServerNames) Then
					Set oUserAuth = oUserAuthDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_INFORMATION_SOURCE & "' $and$ @id='" & aInformationSources(i,0) & "']")
					If Not oUserAuth is Nothing Then
						sAuthID = ""
						lErrNumber = ParseAuthenticationObject(oUserAuth.getAttribute("v"), sAuthName, sAuthPwd, sAuthID)
						If sAuthID = sCastorUserID Then
							sISourceArray(j) = aInformationSources(i,0)
							sAuthObjectArray(j) = "AuthUserName=""" + sCastorUserName + """ AuthUserID=""" + sCastorUserID + """ AuthUserPwd=""" + Encrypt(sCastorPassword) + """ AuthMode=""" + sAuthMode + """"
							j = j + 1
						End If
					End If
				End If
			Next

		    If Err.number <> NO_ERR Then
		        lErrNumber = Err.number
		        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error retrieving information sources", LogLevelError)
		    End If
		End If

		If lErrNumber = NO_ERR Then
		    lErrNumber = co_SaveUserAuthenticationObjects(sSessionID, sISourceArray, sAuthObjectArray)
		    If lErrNumber <> NO_ERR Then
		        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "LoginCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SaveUserAuthenticationObjects", LogLevelTrace)
		    End If
		End If

	End IF

    Set oUserAuthDOM = Nothing
	Set oUserAuth = Nothing

    UpdateAuthenticationObject = lErrNumber
    Err.Clear

End Function

Function getClusterNodeNames()
	Dim oAdminAPI
	Dim sServer
	Dim asServerNames
	Dim oClusterNodes
	Dim lServerCount
	Dim i

	sServer = getIServerName()
	Set oAdminAPI = CreateObject(XMLADMIN_PROGID)
	oAdminAPI.Connect sServer
	Set oClusterNodes = oAdminAPI.getCluster(sServer)
	lServerCount = oClusterNodes.Count

	Redim asServerNames(lServerCount)

	For i = 1 to lServerCount
		asServerNames(i) = oClusterNodes.Item(i).NodeName
	Next

	getClusterNodeNames = asServerNames

End Function

Function isIServerInCluster(sServerName, asServerNames)
	Dim i
	isIServerInCluster = false
	If len(sServerName) > 0 Then
		For i = 1 to UBound(asServerNames)
			If UCase(asServerNames(i)) = UCase(sServerName) Then
				isIServerInCluster = true
				Exit For
			End If
		Next
	End If
End Function


Function getInfSourcesServerNames(aInformationSources)
'********************************************************
'*Purpose:  Returns IS server names
'*Inputs:
'********************************************************
CONST PROCEDURE_NAME = "getInfSourcesServerNames"
Dim lErr
Dim sErr

Dim oSiteInfo
Dim sISXML
Dim sAllISXML

Dim oDOM
Dim oAllDOM
Dim oAllISs
Dim oIS
Dim i

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    If lErr = NO_ERR Then
        sAllISXML = oSiteInfo.getAllInformationSources(Application.Value("SITE_ID"))
        lErr = checkReturnValue(sAllISXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "getAllInformationSources", "Error calling getAllInformationSources", LogLevelError)
    End If

    If lErr = NO_ERR Then
        Set oAllDOM = Server.CreateObject("Microsoft.XMLDOM")
        oAllDOM.async = False
        If oAllDOM.loadXML(sAllISXML) = False Then
            lErr = ERR_XML_LOAD_FAILED
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "loadXML", "Error loading sAllISXML", LogLevelError)
        End If
    End If

    dim oAllNames

    If lErr = NO_ERR Then

        Set oAllISs = Nothing
        Set oAllISs = oAllDOM.selectNodes("//oi[@tp='6']")

        If oAllISs.length > 0 Then
            Redim aInformationSources(oAllISs.length - 1,1)
            For i = 0 To oAllISs.length - 1
                aInformationSources(i, 0) = oAllISs.item(i).getAttribute("id")
                Set oAllNames = oAllISs.item(i).selectSingleNode("prs/pr[@id='IS_ServerName']")
                aInformationSources(i, 1) = oAllNames.getAttribute("v")
            Next
		End If
    End If

    Set oSiteInfo = Nothing
    Set oAllDOM = Nothing
    Set oDOM = Nothing
    Set oIS = Nothing
    Set oAllISs = Nothing
    Set oAllNamess = Nothing

    getISList = lErr
    Err.Clear

End Function

%>