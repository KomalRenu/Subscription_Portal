<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_CreateDeviceType(sSiteID, sDeviceTypePropertiesXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID, sDeviceTypePropertiesXML
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_CreateDeviceType"
	Dim oSiteInfo
	Dim lErrNumber
	Dim sErr
	Dim sCreateDeviceTypeXML

	lErrNumber = NO_ERR

	Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sCreateDeviceTypeXML = oSiteInfo.createDeviceType(sSiteID, sDeviceTypePropertiesXML)
        lErrNumber = checkReturnValue(sCreateDeviceTypeXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "SiteInfo.createDeviceType", "Error while calling createDeviceType", LogLevelError)
        End If
	End If

	Set oSiteInfo = Nothing

	co_CreateDeviceType = lErrNumber
	Err.Clear
End Function

Function co_CreateDeviceTypeDefinitions(sSiteID, asDeviceTypeID, asDeviceTypeDefinitionXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID, asDeviceTypeID, asDeviceTypeDefinitionXML
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_CreateDeviceTypeDefinitions"
	Dim oSiteInfo
	Dim lErrNumber
	Dim sErr
	Dim sCreateDeviceTypeDefinitionsXML

	lErrNumber = NO_ERR

	Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sCreateDeviceTypeDefinitionsXML = oSiteInfo.createDeviceTypeDefinitions(sSiteID, asDeviceTypeID, asDeviceTypeDefinitionXML)
        lErrNumber = checkReturnValue(sCreateDeviceTypeDefinitionsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "SiteInfo.createDeviceTypeDefinitions", "Error while calling createDeviceTypeDefinitions", LogLevelError)
        End If
	End If

	Set oSiteInfo = Nothing

	co_CreateDeviceTypeDefinitions = lErrNumber
	Err.Clear
End Function

Function co_DeleteDeviceType(sSiteID, sDeviceTypeID)
'********************************************************
'*Purpose:
'*Inputs: sSiteID, sDeviceTypeID
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_DeleteDeviceType"
	Dim oSiteInfo
	Dim lErrNumber
	Dim sErr
	Dim sDeleteDeviceTypeXML

	lErrNumber = NO_ERR

	Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sDeleteDeviceTypeXML = oSiteInfo.deleteDeviceType(sSiteID, sDeviceTypeID)
        lErrNumber = checkReturnValue(sDeleteDeviceTypeXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "SiteInfo.deleteDeviceType", "Error while calling deleteDeviceType", LogLevelError)
        End If
	End If

	Set oSiteInfo = Nothing

	co_DeleteDeviceType = lErrNumber
	Err.Clear
End Function

Function co_GetDeviceTypes(sSiteID, sGetDeviceTypesXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID
'*Outputs: sGetDeviceTypesXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetDeviceTypes"
	Dim oSiteInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sGetDeviceTypesXML = oSiteInfo.getDeviceTypes(sSiteID)
        lErrNumber = checkReturnValue(sGetDeviceTypesXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "SiteInfo.getDeviceTypes", "Error while calling getDeviceTypes", LogLevelError)
        End If
	End If

	Set oSiteInfo = Nothing

	co_GetDeviceTypes = lErrNumber
	Err.Clear
End Function

Function co_GetDeviceTypeDefinitions(sSiteID, asDeviceTypeID, sGetDeviceTypeDefinitionsXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID, asDeviceTypeID
'*Outputs: sGetDeviceTypeDefinitionsXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetDeviceTypeDefinitions"
	Dim oSiteInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sGetDeviceTypeDefinitionsXML = oSiteInfo.getDeviceTypeDefinitions(sSiteID, asDeviceTypeID)
        lErrNumber = checkReturnValue(sGetDeviceTypeDefinitionsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "SiteInfo.getDeviceTypeDefinitions", "Error while calling getDeviceTypeDefinitions", LogLevelError)
        End If
	End If

	Set oSiteInfo = Nothing

	co_GetDeviceTypeDefinitions = lErrNumber
	Err.Clear
End Function

Function co_UpdateDeviceTypeDefinitions(sSiteID, asDeviceTypeID, asDeviceTypeDefinitionXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID, asDeviceTypeID, asDeviceTypeDefinitionXML
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_UpdateDeviceTypeDefinitions"
	Dim oSiteInfo
	Dim lErrNumber
	Dim sErr
	Dim sUpdateDeviceTypeDefinitionsXML

	lErrNumber = NO_ERR

	Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sUpdateDeviceTypeDefinitionsXML = oSiteInfo.updateDeviceTypeDefinitions(sSiteID, asDeviceTypeID, asDeviceTypeDefinitionXML)
        lErrNumber = checkReturnValue(sUpdateDeviceTypeDefinitionsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "SiteInfo.updateDeviceTypeDefinitions", "Error while calling updateDeviceTypeDefinitions", LogLevelError)
        End If
	End If

	Set oSiteInfo = Nothing

	co_UpdateDeviceTypeDefinitions = lErrNumber
	Err.Clear
End Function

Function co_GetFolderContents(sSiteID, sFolderID, iDefaultFolder, sGetFolderContentsXML)
'********************************************************
'*Purpose:
'*Inputs: sSiteID, sFolderID, iDefaultFolder
'*Outputs: sGetFolderContentsXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetFolderContents"
	Dim oSiteInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sGetFolderContentsXML = oSiteInfo.getFolderContents(sSiteID, sFolderID, iDefaultFolder)
        lErrNumber = checkReturnValue(sGetFolderContentsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "DeviceTypesCoLib.asp", PROCEDURE_NAME, "SiteInfo.getFolderContents", "Error while calling getFolderContents", LogLevelError)
        End If
	End If

	Set oSiteInfo = Nothing

	co_GetFolderContents = lErrNumber
	Err.Clear
End Function
%>