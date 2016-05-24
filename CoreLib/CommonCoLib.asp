<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%

Function co_GetAllPortals(sPortalXML)
'********************************************************
'*Purpose:  Returns a list of all the Portals found on the current machine
'*Inputs:   None
'*Outputs:  Array of Portals containing Portal Name, Associated Site Name and Default Property
'           Count of Portals.
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_GetAllPortals"
    Dim lErr
    Dim sErr
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sPortalXML = oAdmin.getVirtualDirectories()
        lErr = checkReturnValue(sPortalXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "getVirtualDirectories", "Error calling getVirtualDirectories", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_GetAllPortals = lErr
    Err.Clear
End Function

Function co_GetAvailableSubscriptions(sSessionID, asSubscriptionGUIDS, sGetAvailableSubscriptionsXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, asSubscriptionGUIDS
'*Outputs: sGetAvailableSubscriptionsXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetAvailableSubscriptions"
    Dim oDocRepository
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oDocRepository = Server.CreateObject(PROGID_DOC_REPOSITORY)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_DOC_REPOSITORY, LogLevelError)
    Else
        sGetAvailableSubscriptionsXML = oDocRepository.getAvailableSubscriptions(sSessionID, asSubscriptionGUIDS)
        lErrNumber = checkReturnValue(sGetAvailableSubscriptionsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErr), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "DocRepository.getAvailableSubscriptions", "Error calling getAvailableSubscriptions", LogLevelError)
        End If
    End If

    Set oDocRepository = Nothing

    co_GetAvailableSubscriptions = lErrNumber
    Err.Clear
End Function

Function co_GetChannelsForSite(sSiteId, sChannelsXML)
'********************************************************
'*Purpose: Returns the Channels for the given site
'*Inputs:  sSiteId
'*Outputs: sChannelsXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_GetChannelsForSite"
    Dim lErr
    Dim sErr
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.number), Err.source, "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sChannelsXML = oSiteInfo.getChannelsForSite(sSiteId)
        lErr = checkReturnValue(sChannelsXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErr), sErr, Err.source, "CommonCoLib.asp", PROCEDURE_NAME, "getChannelsForSite", "Error calling getChannelsForSite: SiteID=" & sSiteID, LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_GetChannelsForSite = lErr
    Err.Clear
End Function

Function co_GetLocalesForSite(sSiteID, sGetLocalesForSiteXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID
'*Outputs: sGetLocalesForSiteXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetLocalesForSite"
    Dim oSiteInfo
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR
    
    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sGetLocalesForSiteXML = co_GetLocalesForSiteInt(oSiteInfo, sSiteID)

        lErrNumber = checkReturnValue(sGetLocalesForSiteXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErr), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "SiteInfo.getLocalesForSite", "Error calling getLocalesForSite", LogLevelError)
        End If
    End If
    

    Set oSiteInfo = Nothing

    co_GetLocalesForSite = lErrNumber
    Err.Clear
End Function

Function co_GetLocalesForSiteInt(oSiteInfo, sSiteID)
	Dim sLngObjID

	sLngObjID = GetSiteLocale()
	If sLngObjID <> "" Then
		co_GetLocalesForSiteInt = oSiteInfo.getAllLocalesByLocaleID(sSiteID, sLngObjID)
	Else
		co_GetLocalesForSiteInt = oSiteInfo.getAllLocales(sSiteID)
	End If
end Function

Function co_GetPreferenceObjects(sSessionID, asPreferenceObjectID, sGetPreferenceObjectsXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, asPreferenceObjectID
'*Outputs: sGetPreferenceObjectsXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetPreferenceObjects"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        sGetPreferenceObjectsXML = oPersonalizationInfo.getPreferenceObjects(sSessionID, asPreferenceObjectID)
        lErrNumber = checkReturnValue(sGetPreferenceObjectsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErr), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.getPreferenceObjects", "Error while calling getPreferenceObjects", LogLevelError)
        End If
	End If

	Set oPersonalizationInfo = Nothing

	co_GetPreferenceObjects = lErrNumber
	Err.Clear
End Function

Function co_GetProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID, sGetProfileXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID, sPreferenceObjectID, sQuestionObjectID
'*Outputs: sGetProfileXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetProfile"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        sGetProfileXML = oPersonalizationInfo.getProfile(sSessionID, sPreferenceObjectID, sQuestionObjectID)
        lErrNumber = checkReturnValue(sGetProfileXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErr), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.getProfile", "Error while calling getProfile", LogLevelError)
        End If
	End If

	Set oPersonalizationInfo = Nothing

	co_GetProfile = lErrNumber
	Err.Clear
End Function

Function co_GetSiteProperties(sSiteId, sSitePropsXML)
'********************************************************
'*Purpose: Returns the Site properties stored in MD
'*Inputs:  none
'*Outputs: aSiteProperties
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetSiteProperties"
    Dim lErr
    Dim sErr
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sSitePropsXML = oSiteInfo.getSiteProperties(sSiteId)
        lErr = checkReturnValue(sSitePropsXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "getSiteProperties", "Error calling getSiteProperties", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_GetSiteProperties = lErr
    Err.Clear
End Function

Function co_SetSiteProperties(sSiteId, sConfigXML)
'********************************************************
'*Purpose:  Set the Properties of a site
'*Inputs:   A properites array, the Flags indicate which elements of the array
'           have valid information, using the FLAGS_PROP constants
'*Outputs:  none
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_SetSiteProperties"
    Dim lErrNumber
    Dim sErr
    Dim sReturnXML
    Dim oSiteInfo

    lErrNumber = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), Err.source, "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturnXML = oSiteInfo.updateSiteProperties(sSiteId, sConfigXML)
        lErrNumber = checkReturnValue(sReturnXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, sErr, Err.source, "CommonCoLib.asp", PROCEDURE_NAME, "updateSiteProperties", "Error calling updateSiteProperties", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_SetSiteProperties = lErrNumber
    Err.Clear
End Function

Function co_CreateSiteProperties(sSiteId, sConfigXML)
'********************************************************
'*Purpose:  Creates the Properties for a site
'*Inputs:   Properties XML and SiteID
'*Outputs:  none
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_CreateSiteProperties"
    Dim lErrNumber
    Dim sErr
    Dim sReturnXML
    Dim oSiteInfo

    lErrNumber = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), Err.source, "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturnXML = oSiteInfo.insertSiteProperties(sSiteId, sConfigXML)
        lErrNumber = checkReturnValue(sReturnXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, sErr, Err.source, "CommonCoLib.asp", PROCEDURE_NAME, "insertSiteProperties", "Error calling insertSiteProperties", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_CreateSiteProperties = lErrNumber
    Err.Clear
End Function

Function co_GetUserAuthenticationObjects(sSessionID, sGetUserAuthenticationObjectsXML)
'********************************************************
'*Purpose:
'*Inputs: sSessionID
'*Outputs: sGetUserAuthenticationObjectsXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetUserAuthenticationObjects"
    Dim oUser
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oUser = Server.CreateObject(PROGID_USER)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_USER, LogLevelError)
    Else
        sGetUserAuthenticationObjectsXML = oUser.getUserAuthenticationObjects(sSessionID)
        lErrNumber = checkReturnValue(sGetUserAuthenticationObjectsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErr), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "User.getUserAuthenticationObjects", "Error calling getUserAuthenticationObjects", LogLevelError)
        End If
    End If

    Set oUser = Nothing

    co_GetUserAuthenticationObjects = lErrNumber
    Err.Clear
End Function

Function co_GetSharedPropertyManager(sGroupName, sValue)
'********************************************************
'*Purpose: Get the shared property value for given Portal
'*Inputs: Portal Name
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetSharedPropertyManager"
    Dim oSPM
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR
    Set oSPM = Server.CreateObject(PROGID_PORTAL_VB_ADMIN)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PORTAL_VB_ADMIN, LogLevelError)
    Else
        lErrNumber = oSPM.GetSpm(sGroupName, sValue, sErr)
        If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErrNumber, sErr, CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "GetSpm", "Error calling GetSpm for Group: " & sGroupName, LogLevelError)
    End If

    Set oSPM = Nothing

    co_GetSharedPropertyManager = lErrNumber
    Err.Clear
End Function

Function co_SetSharedPropertyManager(sGroupName, sValue)
'********************************************************
'*Purpose: Set the shared property value for given portal
'*Inputs: sValue to be set and portal name
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_SetSharedPropertyManager"
    Dim oSPM
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR
    Set oSPM = Server.CreateObject(PROGID_PORTAL_VB_ADMIN)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PORTAL_VB_ADMIN, LogLevelError)
    Else
        lErr = oSPM.SetSpm(sGroupName, sValue, sErr)
        If lErr <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "SetSpm", "Error calling SetSpm for Group: " & sGroupName, LogLevelError)
        End If
    End If

    Set oSPM = Nothing

    co_SetSharedPropertyManager = lErrNumber
    Err.Clear
End Function


Function co_CreateTransmissionProperties(sSiteID, asTransPropID, asTransProperty, sCreateTransmissionPropertiesXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID, asTransPropID, asTransProperty
'*Outputs: sCreateTransmissionPropertiesXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_CreateTransmissionProperties"
    Dim oSiteInfo
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sCreateTransmissionPropertiesXML = oSiteInfo.createTransmissionProperties(sSiteID, asTransPropID, asTransProperty)
        lErrNumber = checkReturnValue(sCreateTransmissionPropertiesXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErr), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "SiteInfo.createTransmissionProperties", "Error calling createTransmissionProperties", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_CreateTransmissionProperties = lErrNumber
    Err.Clear
End Function

Function co_DeleteTransmissionProperties(sSiteID, asTransPropID, sDeleteTransmissionPropertiesXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID, asTransPropID,
'*Outputs: sCreateTransmissionPropertiesXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_DeleteTransmissionProperties"
    Dim oSiteInfo
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sDeleteTransmissionPropertiesXML = oSiteInfo.deleteTransmissionProperties(sSiteID, asTransPropID)
        lErrNumber = checkReturnValue(sDeleteTransmissionPropertiesXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErr), CStr(Err.source), "CommonCoLib.asp", PROCEDURE_NAME, "SiteInfo.deleteTransmissionProperties", "Error calling deleteTransmissionProperties", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_DeleteTransmissionProperties = lErrNumber
    Err.Clear
End Function

%>