<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/AddressCoLib.asp" -->
<%
	'Transaction Indices
	Const INDEX_ADDRESS_ACTION = 0

	Function ParseRequestForModifyAddress(oRequest, sAction, sDeviceTypeID, sAddressName, sPhysicalAddress, sDevice, sPIN, sDeliveryWindow, sDelAddrID, sEditAddID, sAddressNameVld, sPhysicalAddressVld, sPINVld, sTransPropsID)
	'********************************************************
	'*Purpose:
	'*Inputs: oRequest
	'*Outputs: sAction, sDeviceTypeID, sAddressName, sPhysicalAddress, sDevice, sPIN, sDeliveryWindow, sDelAddrID, sEditAddID, sAddressNameVld, sPhysicalAddressVld, sPINVld
	'********************************************************
		On Error Resume Next
        Dim lErrNumber

        lErrNumber = NO_ERR

		sAction = ""
		sDeviceTypeID = ""
		sAddressName = ""
		sPhysicalAddress = ""
		sDevice = ""
		sPIN = ""
		sDeliveryWindow = ""
		sDelAddrID = ""
		sEditAddID = ""
		sAddressNameVld = ""
		sPhysicalAddressVld = ""
		sPINVld = ""
		sTransPropsID = ""

		sAction = Trim(CStr(oRequest("action")))
		sDeviceTypeID = Trim(CStr(oRequest("deviceTypeID")))
		sAddressName = Trim(CStr(oRequest("AddressName")))
		sPhysicalAddress = Trim(CStr(oRequest("PhysicalAddress")))
		sDevice = Trim(CStr(oRequest("Device")))
		sPIN = Trim(CStr(oRequest("PIN")))
		sDeliveryWindow = Trim(CStr(oRequest("DeliveryWindowStart"))) & Trim(CStr(oRequest("DeliveryWindowEnd")))
		sDelAddrID = Trim(CStr(oRequest("delAddrID")))
		sEditAddID = Trim(CStr(oRequest("editAddID")))
		sAddressNameVld = Trim(CStr(oRequest("AddressNameVld")))
		sPhysicalAddressVld = Trim(CStr(oRequest("PhysicalAddressVld")))
		sPINVld = Trim(CStr(oRequest("PINVld")))
		sTransPropsID = Trim(CStr(oRequest("transPropsID")))

		If Err.number <> NO_ERR Then
		    lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifyAddressCuLib.asp", "ParseRequestForModifyAddress", "", "Error setting variables equal to Request variables", LogLevelError)
		Else
		    If Len(sAction) = 0 Then
		    	lErrNumber = URL_MISSING_PARAMETER
		    End If
		End If

		ParseRequestForModifyAddress = lErrNumber
		Err.Clear
	End Function

Function validate_AddressFields(sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim tempErr
    lErrNumber = NO_ERR
    tempErr = NO_ERR

    If Len(sAddressName) = 0 Or Len(sPhysicalAddress) = 0 Then
		lErrNumber = lErrNumber + ERR_ADDRESS_BLANKS
	End If

    If sPhysicalAddressVld <> "" Then
		Select Case sPhysicalAddressVld
			Case "email"
				tempErr = ValidateEmailAddress(sPhysicalAddress)
				If tempErr = -1 Then
					lErrNumber = lErrNumber + ERR_EMAIL_ADDR_INVALID
				End If
            Case "number"
                tempErr = ValidateNumberAddress(sPhysicalAddress)
                If tempErr = -1 Then
                    lErrNumber = lErrNumber + ERR_NUMBER_ADDR_INVALID
                End If
			Case Else
		End Select
	End If

    validate_AddressFields = lErrNumber
    Err.Clear
End Function

Function cu_AddAddress(sAddressName, sPhysicalAddress, sDevice, sPIN, sDeliveryWindow)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_AddAddress"
	Dim lErrNumber
	Dim asAddressProperties()
	Redim asAddressProperties(MAX_ADDR_PROP)
	Dim sSessionID
	Dim sAddAddressXML
	Dim bGenerateTransProps

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
	bGenerateTransProps = True
	If Len(sDeliveryWindow) = 0 Then
	    sDeliveryWindow = "0000"
	End If

	asAddressProperties(ADDR_PROP_ADDRESS_ID) = GetGUID()							'addressID
	asAddressProperties(ADDR_PROP_ADDRESS_NAME) = sAddressName						'addressName
	asAddressProperties(ADDR_PROP_PHYSICAL_ADDRESS) = sPhysicalAddress					'physicalAddress
	asAddressProperties(ADDR_PROP_ADDRESS_DISPLAY) = sAddressName						'addressDisplay
	asAddressProperties(ADDR_PROP_DEVICE_ID) = sDevice							'deviceID
	asAddressProperties(ADDR_PROP_DELIVERY_WINDOW) = sDeliveryWindow						'deliveryWindow
	asAddressProperties(ADDR_PROP_TIMEZONE_ID) = GetDefaultTimeZone()								'DefaultTimezoneStdName
	asAddressProperties(ADDR_PROP_STATUS) = "1"								'status
	asAddressProperties(ADDR_PROP_CREATED_BY) = ""								'createdBy
	asAddressProperties(ADDR_PROP_LAST_MODIFIED_BY) = ""								'lastModBy
	asAddressProperties(ADDR_PROP_TRANSMISSION_PROPERTIES_ID) = ""    'transPropsID
	asAddressProperties(ADDR_PROP_PIN) = sPIN								'PIN
	asAddressProperties(ADDR_PROP_EXPIRATION_DATE) = ""		'expirationDate
	asAddressProperties(ADDR_PROP_CREATED_DATE) = ""		'createdDate
	asAddressProperties(ADDR_PROP_LAST_MODIFIED_DATE) = ""		'lastModDate

	lErrNumber = co_AddAddress(sSessionID, asAddressProperties, bGenerateTransProps, sAddAddressXML)
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifyAddressCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_AddAddress", LogLevelTrace)
	End If

	cu_AddAddress = lErrNumber
	Err.Clear
End Function

Function cu_EditAddress(sEditAddID, sAddressName, sPhysicalAddress, sDevice, sDeliveryWindow, sPIN, sTransPropsID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_EditAddress"
	Dim lErrNumber
	Dim asAddressProperties()
	Redim asAddressProperties(MAX_ADDR_PROP)
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
	If Len(sDeliveryWindow) = 0 Then
	    sDeliveryWindow = "0000"
	End If

	asAddressProperties(ADDR_PROP_ADDRESS_ID) = sEditAddID		'addressID
	asAddressProperties(ADDR_PROP_ADDRESS_NAME) = sAddressName		'addressName
	asAddressProperties(ADDR_PROP_PHYSICAL_ADDRESS) = sPhysicalAddress	'physicalAddress
	asAddressProperties(ADDR_PROP_ADDRESS_DISPLAY) = sAddressName		'addressDisplay
	asAddressProperties(ADDR_PROP_DEVICE_ID) = sDevice			'deviceID
	asAddressProperties(ADDR_PROP_DELIVERY_WINDOW) = sDeliveryWindow		'deliveryWindow
    asAddressProperties(ADDR_PROP_TIMEZONE_ID) = "21"								'timezoneID
    asAddressProperties(ADDR_PROP_STATUS) = "1"								'status
    asAddressProperties(ADDR_PROP_CREATED_BY) = ""								'createdBy
    asAddressProperties(ADDR_PROP_LAST_MODIFIED_BY) = ""								'lastModBy
    asAddressProperties(ADDR_PROP_TRANSMISSION_PROPERTIES_ID) = sTransPropsID    'transPropsID
    asAddressProperties(ADDR_PROP_PIN) = sPIN								'PIN
    asAddressProperties(ADDR_PROP_EXPIRATION_DATE) = ""		'expirationDate
    asAddressProperties(ADDR_PROP_CREATED_DATE) = ""		'createdDate
    asAddressProperties(ADDR_PROP_LAST_MODIFIED_DATE) = ""		'lastModDate

	lErrNumber = co_EditAddress(sSessionID, asAddressProperties)
	If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifyAddressCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_EditAddress", LogLevelTrace)
	End If

	cu_EditAddress = lErrNumber
	Err.Clear
End Function

Function cu_DeleteAddress(sDelAddrID, sGetUserAddressesXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_DeleteAddress"
	Dim lErrNumber
	Dim asAddressID(0)
	Dim sSessionID
	Dim oAddressDOM
	Dim oCurrAddr
	Dim asTransPropsID(0)

	lErrNumber = NO_ERR
	asAddressID(0) = CStr(sDelAddrID)
	sSessionID = GetSessionID()

	lErrNumber = co_DeleteAddresses(sSessionID, asAddressID)
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifyAddressCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_DeleteAddresses", LogLevelTrace)
	End If

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetUserAddressesXML, oAddressDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "GetQuestionProperty", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		Set oCurrAddr = oAddressDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_ADDRESS & "' $and$ @id='" & sDelAddrID & "']")
		asTransPropsID(0) = oCurrAddr.getAttribute("trps")
	End If

	lErrNumber = cu_DeleteTransmissionProperties(asTransPropsID)
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ModifyAddressCuLib.asp", PROCEDURE_NAME, "", "Error while calling cu_DeleteTransmissionProperties", LogLevelTrace)
	End If

	cu_DeleteAddress = lErrNumber
	Err.Clear
End Function

Function ValidateAddressName(sAddressName)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim i
	Dim bFound

	bFound = False

	For i = 0 to Ubound(asReservedChars,1)
		If InStr(sAddressName, asReservedChars(i)) > 0 Then
			bFound = True
			Exit For
		End If
	Next

	If bFound = True Then ValidateAddressName = -1
End Function

%>
