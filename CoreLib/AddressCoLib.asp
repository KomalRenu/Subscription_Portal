<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function co_AddAddress(sSessionID, asAddressProperties, bGenerateTransProps, sAddAddressXML)
'********************************************************
'*Purpose: Creates a user address.  If bGenerateTransProps = True, then
'          the call to cu_CreateTransmissionProperties will return a new
'          sTransPropsID.
'*Inputs: sSessionID, asAddressProperties, bGenerateTransProps
'*Outputs: sAddAddressXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_AddAddress"
	Dim oAddresses
	Dim lErrNumber
	Dim sTransPropsID
	Dim sErr

	lErrNumber = NO_ERR

    If bGenerateTransProps = False Then
        sTransPropsID = asAddressProperties(ADDR_PROP_TRANSMISSION_PROPERTIES_ID)
    Else
        sTransPropsID = ""
    End If

    lErrNumber = cu_CreateTransmissionProperties(sTransPropsID)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "", "Error calling cu_CreateTransmissionProperties", LogLevelTrace)
    Else
        If bGenerateTransProps = True Then
            asAddressProperties(ADDR_PROP_TRANSMISSION_PROPERTIES_ID) = sTransPropsID
        End If
    End If

    If lErrNumber = NO_ERR Then
	    Set oAddresses = Server.CreateObject(PROGID_ADDRESS)
	    If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADDRESS, LogLevelError)
        Else
            sAddAddressXML = oAddresses.addAddress(sSessionID, asAddressProperties)
            lErrNumber = checkReturnValue(sAddAddressXML, sErr)
            If lErrNumber <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "Addresses.addAddress", "Error while calling addAddress", LogLevelError)
            End If
	    End If
	End If

	Set oAddresses = Nothing

	co_AddAddress = lErrNumber
	Err.Clear
End Function

Function co_DeleteAddresses(sSessionID, asAddressID)
'********************************************************
'*Purpose: Deletes one or more user addresses.
'*Inputs: sSessionID, asAddressID (an array of addressIDs)
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_DeleteAddresses"
	Dim oAddresses
	Dim lErrNumber
	Dim sErr
	Dim sDeleteAddressXML

	lErrNumber = NO_ERR

	Set oAddresses = Server.CreateObject(PROGID_ADDRESS)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADDRESS, LogLevelError)
    Else
        sDeleteAddressXML = oAddresses.deleteAddresses(sSessionID, asAddressID)
	    lErrNumber = checkReturnValue(sDeleteAddressXML, sErr)
	    If lErrNumber <> NO_ERR Then
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "Addresses.deleteAddresses", "Error while calling deleteAddresses", LogLevelError)
	    End If
	End If

	Set oAddresses = Nothing

	co_DeleteAddresses = lErrNumber
	Err.Clear
End Function

Function co_EditAddress(sSessionID, asAddressProperties)
'********************************************************
'*Purpose: Edits a user address.
'*Inputs: sSessionID, asAddressProperties
'*Outputs: sEditAddressXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_EditAddress"
	Dim oAddresses
	Dim lErrNumber
	Dim sErr
	Dim sEditAddressXML

	lErrNumber = NO_ERR

	Set oAddresses = Server.CreateObject(PROGID_ADDRESS)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADDRESS, LogLevelError)
    Else
        sEditAddressXML = oAddresses.editAddress(sSessionID, asAddressProperties)
        lErrNumber = checkReturnValue(sEditAddressXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "Addresses.editAddress", "Error while calling editAddress", LogLevelError)
        End If
	End If

	Set oAddresses = Nothing

	co_EditAddress = lErrNumber
	Err.Clear
End Function

Function co_GetDevicesInFolders(sSessionID, asFolderID, iCommentType, bBrowseSubFolders, sGetDevicesInFoldersXML)
'********************************************************
'*Purpose: Given one or more folderIDs, returns all devices in those folders.
'          If bBrowseSubFolders = True, then also returns devices in any
'          sub-folders of the given folders.  iCommentType specifies what
'          kind of description to return for devices (0 = noDesc, 1 = textDesc,
'          2 = htmlDesc, 3 = textDesc+htmlDesc)
'*Inputs: sSessionID, asFolderID, iCommentType, bBrowseSubFolders
'*Outputs: sGetDevicesInFoldersXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetDevicesInFolders"
	Dim oSystemInfo
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oSystemInfo = Server.CreateObject(PROGID_SYSTEM_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SYSTEM_INFO, LogLevelError)
    Else
        sGetDevicesInFoldersXML = oSystemInfo.getDevicesInFolders(sSessionID, asFolderID, iCommentType, bBrowseSubFolders)
        lErrNumber = checkReturnValue(sGetDevicesInFoldersXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "SystemInfo.getDevicesInFolders", "Error while calling getDevicesInFolders", LogLevelError)
        End If
	End If

	Set oSystemInfo = Nothing

	co_GetDevicesInFolders = lErrNumber
	Err.Clear
End Function

Function co_GetUserAddresses(sSessionID, sGetUserAddressesXML)
'********************************************************
'*Purpose: Retrieves addresses for a given user.
'*Inputs: sSessionID
'*Outputs: sGetUserAddressesXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "co_GetUserAddresses"
	Dim oAddresses
	Dim lErrNumber
	Dim sErr

	lErrNumber = NO_ERR

	Set oAddresses = Server.CreateObject(PROGID_ADDRESS)
	If Err.number <> NO_ERR Then
		lErrNumber = Err.number
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADDRESS, LogLevelError)
    Else
        sGetUserAddressesXML = oAddresses.getUserAddresses(sSessionID)
        lErrNumber = checkReturnValue(sGetUserAddressesXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "AddressCoLib.asp", PROCEDURE_NAME, "Addresses.getUserAddresses", "Error while calling getUserAddresses", LogLevelError)
        End If
	End If

	Set oAddresses = Nothing

	co_GetUserAddresses = lErrNumber
	Err.Clear
End Function

%>