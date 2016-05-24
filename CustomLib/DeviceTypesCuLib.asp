<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/DeviceTypesCoLib.asp" -->
<%
'Information about a device type:
Const DEV_TYPE_ID    = 0
Const DEV_TYPE_NAME  = 1
Const DEV_TYPE_LARGE_ICON = 2
Const DEV_TYPE_SMALL_ICON = 3
Const DEV_TYPE_ADDR_FORMAT = 4
Const DEV_TYPE_DISPLAY_ADDR_NAME = 5
Const DEV_TYPE_DISPLAY_ADDR_VALUE = 6
Const DEV_TYPE_DISPLAY_STYLE = 7
Const DEV_TYPE_DISPLAY_DELIVERY_WINDOW = 8
Const DEV_TYPE_EDIT_PIN = 9
Const DEV_TYPE_EDIT_DELIVERY_WINDOW = 10
Const DEV_TYPE_ACTION = 11
Const MAX_DEV_TYPE_INFO = 11

'ID For the display fields:
Const IDS_NAME = 508        'asDescriptors(508):Name
Const IDS_ADDRESS = 367     'asDescriptors(367):Address
Const IDS_STYLE = 510       'asDescriptors(510):Style
Const IDS_PIN = 511         'asDescriptors(511):PIN
Const IDS_DELIVERY_WINDOW = 512   'asDescriptors(512):DeliveryWindow
Const IDS_PASSWORD = 527    'asDescriptors(527):Confirm Password

Const IMAGE_EMPTY = "images/1ptrans.gif"


Function ParseRequestForDTFolders(oRequest, aDeviceTypeInfo, sFolderID)
'********************************************************
'*Purpose: Reads the request for browsing the DT folders.
'*Inputs:  oRequest, The request object;
'*Outputs: aDeviceTypeInfo: the information about the device type from the request object
'********************************************************

	On Error Resume Next
    Dim lErrNumber

    lErrNumber = NO_ERR

    lErrNumber = ParseRequestForDeviceType(oRequest, aDeviceTypeInfo)
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "ParseRequestForDTFolders", "", "Error setting variables to Request variables", LogLevelError)
    Else
        aDeviceTypeInfo(DEV_TYPE_ACTION) = "edit"
	    sFolderID = Trim(CStr(oRequest("folderID")))
	End If

	ParseRequestForDTFolders = lErrNumber
	Err.Clear
End Function

Function ParseRequestForDeviceType(oRequest, aDeviceTypeInfo)
'********************************************************
'*Purpose: Reads the device information from the request object
'*Inputs:  oRequest, The request object;
'*Outputs: aDeviceTypeInfo: the information about the device type from the request object
'********************************************************
	Dim lErrNumber

    On Error Resume Next
    lErrNumber = NO_ERR

    Redim aDeviceTypeInfo(MAX_DEV_TYPE_INFO)

    aDeviceTypeInfo(DEV_TYPE_ID) = Trim(CStr(oRequest("dtID")))
    aDeviceTypeInfo(DEV_TYPE_NAME) = Trim(CStr(oRequest("DTName")))
    aDeviceTypeInfo(DEV_TYPE_LARGE_ICON) = Trim(CStr(oRequest("DTLargeIcon")))
    aDeviceTypeInfo(DEV_TYPE_SMALL_ICON) = Trim(CStr(oRequest("DTSmallIcon")))
    aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = Trim(CStr(oRequest("addrFormat")))
    aDeviceTypeInfo(DEV_TYPE_ACTION) = Trim(CStr(oRequest("action")))

	aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME) = Trim(CStr(oRequest("DAddrName")))
	aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE) = CStr(oRequest("DAddrValue"))
	aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE) = CStr(oRequest("DStyle"))
	aDeviceTypeInfo(DEV_TYPE_DISPLAY_DELIVERY_WINDOW) = CStr(oRequest("DDeliveryWindow"))

	aDeviceTypeInfo(DEV_TYPE_EDIT_PIN) = CStr(oRequest("EPin"))
	aDeviceTypeInfo(DEV_TYPE_EDIT_DELIVERY_WINDOW) = CStr(oRequest("EDeliveryWindow"))

	ParseRequestForDeviceType = lErrNumber
	Err.Clear

End Function


Function CreateRequestForDeviceType(aDeviceTypeInfo)
'********************************************************
'*Purpose: Creates a request for a DeviceType based on the DeviceTypeInfo
'*Inputs:  aDeviceTypeInfo: The information about the device type.
'*Outputs: The request for a link to a device type.
'********************************************************
	Dim sRequest

    On Error Resume Next

    If Len(aDeviceTypeInfo(DEV_TYPE_ID)) > 0 Then sRequest =  sRequest & "&dtID=" & aDeviceTypeInfo(DEV_TYPE_ID)
    If Len(aDeviceTypeInfo(DEV_TYPE_ACTION)) > 0 Then sRequest =  sRequest & "&action=" & aDeviceTypeInfo(DEV_TYPE_ACTION)
    If Len(aDeviceTypeInfo(DEV_TYPE_NAME)) > 0 Then sRequest =  sRequest & "&DTName=" & Server.URLEncode(aDeviceTypeInfo(DEV_TYPE_NAME))
    If Len(aDeviceTypeInfo(DEV_TYPE_LARGE_ICON)) > 0 Then sRequest =  sRequest & "&DTLargeIcon=" & Server.URLEncode(aDeviceTypeInfo(DEV_TYPE_LARGE_ICON))
    If Len(aDeviceTypeInfo(DEV_TYPE_SMALL_ICON)) > 0 Then sRequest =  sRequest & "&DTSmallIcon=" & Server.URLEncode(aDeviceTypeInfo(DEV_TYPE_SMALL_ICON))
    If Len(aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT)) > 0 Then sRequest =  sRequest & "&addrFormat=" & aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT)

	If Len(aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME)) > 0 Then sRequest =  sRequest & "&DAddrName=" & aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME)
	If Len(aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE)) > 0 Then sRequest =  sRequest & "&DAddrValue=" & aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE)
	If Len(aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE)) > 0 Then sRequest =  sRequest & "&DStyle=" & aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE)
	If Len(aDeviceTypeInfo(DEV_TYPE_DISPLAY_DELIVERY_WINDOW)) > 0 Then sRequest =  sRequest & "&DDeliveryWindow=" & aDeviceTypeInfo(DEV_TYPE_DISPLAY_DELIVERY_WINDOW)

	If Len(aDeviceTypeInfo(DEV_TYPE_EDIT_PIN)) > 0 Then sRequest =  sRequest & "&EPin=" & aDeviceTypeInfo(DEV_TYPE_EDIT_PIN)
	If Len(aDeviceTypeInfo(DEV_TYPE_EDIT_DELIVERY_WINDOW)) > 0 Then sRequest =  sRequest & "&EDeliveryWindow=" & aDeviceTypeInfo(DEV_TYPE_EDIT_DELIVERY_WINDOW)

	If Len(sRequest) > 0 Then sRequest = Mid(sRequest, 2)

	CreateRequestForDeviceType = sRequest
	Err.Clear

End Function

Function EditDeviceType(aDeviceTypeInfo, sDeviceTypesXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oDeviceType
	Dim oTempNode
	Dim oValueNode
	Dim oValueNode2

	lErrNumber = NO_ERR

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "EditDeviceType", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & aDeviceTypeInfo(DEV_TYPE_ID) & "']")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "EditDeviceType", "", "Error loading devicetype node", LogLevelError)
        End If
    End If

	If lErrNumber = NO_ERR Then
	    If Len(aDeviceTypeInfo(DEV_TYPE_SMALL_ICON)) > 0 Then
	        oDeviceType.selectSingleNode("icon").text = aDeviceTypeInfo(DEV_TYPE_SMALL_ICON)
	    Else
	        oDeviceType.selectSingleNode("icon").text = IMAGE_EMPTY
	    End If

	    If Len(aDeviceTypeInfo(DEV_TYPE_LARGE_ICON)) > 0 Then
	        oDeviceType.selectSingleNode("largeicon").text = aDeviceTypeInfo(DEV_TYPE_LARGE_ICON)
	    Else
	        oDeviceType.selectSingleNode("largeicon").text = IMAGE_EMPTY
	    End If

	    'Display fields
	    Set oTempNode = oDeviceType.selectSingleNode("displayfields")

	    Set oValueNode = oTempNode.selectSingleNode("field[@col = 'n']")
	    If aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME) = "1" Then
	        If (oValueNode Is Nothing) Then
	            Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_NAME
	            oValueNode2.setAttribute "n", "Name"
	            oValueNode2.setAttribute "col", "n"
	        End If
	    Else
	        If Not (oValueNode Is Nothing) Then
	            oTempNode.removeChild(oValueNode)
	        End If
	    End If

	    Set oValueNode = oTempNode.selectSingleNode("field[@col = 'v']")
	    If aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE) = "1" Then
	        If (oValueNode Is Nothing) Then
	            Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_ADDRESS
	            oValueNode2.setAttribute "n", "Address"
	            oValueNode2.setAttribute "col", "v"
	        End If
	    Else
	        If Not (oValueNode Is Nothing) Then
	            oTempNode.removeChild(oValueNode)
	        End If
	    End If

	    Set oValueNode = oTempNode.selectSingleNode("field[@col = 'dvid']")
	    If aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE) = "1" Then
	        If (oValueNode Is Nothing) Then
	            Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_STYLE
	            oValueNode2.setAttribute "n", "Style"
	            oValueNode2.setAttribute "col", "dvid"
	        End If
	    Else
	        If Not (oValueNode Is Nothing) Then
	            oTempNode.removeChild(oValueNode)
	        End If
	    End If

	    Set oValueNode = oTempNode.selectSingleNode("field[@col = 'cb']")
	    If aDeviceTypeInfo(DEV_TYPE_DISPLAY_DELIVERY_WINDOW) = "1" Then
	        If (oValueNode Is Nothing) Then
	            Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_DELIVERY_WINDOW
	            oValueNode2.setAttribute "n", "DeliveryWindow"
	            oValueNode2.setAttribute "col", "cb"
	        End If
	    Else
	        If Not (oValueNode Is Nothing) Then
	            oTempNode.removeChild(oValueNode)
	        End If
	    End If

        'Edit fields
        Set oTempNode = oDeviceType.selectSingleNode("editfields")

	    Set oValueNode = oTempNode.selectSingleNode("field[@col = 'n']")
	    If (oValueNode Is Nothing) Then
	        Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
	        oValueNode2.setAttribute "di", IDS_NAME
	        oValueNode2.setAttribute "n", "AddressName"
	        oValueNode2.setAttribute "col", "n"
	        oValueNode2.setAttribute "size", "15"
	        oValueNode2.setAttribute "type", "text"
	        oValueNode2.setAttribute "vld", "name"
	    End If

	    Set oValueNode = oTempNode.selectSingleNode("field[@col = 'v']")
	    If (oValueNode Is Nothing) Then
	        Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
	        oValueNode2.setAttribute "di", IDS_ADDRESS
	        oValueNode2.setAttribute "n", "PhysicalAddress"
	        oValueNode2.setAttribute "col", "v"
	        oValueNode2.setAttribute "size", "15"
	        oValueNode2.setAttribute "type", "text"
	        If aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_EMAIL Then
	            oValueNode2.setAttribute "vld", "email"
	        ElseIf aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_NUMBER Then
	            oValueNode2.setAttribute "vld", "number"
            Else
                oValueNode2.setAttribute "vld", "none"
	        End If
	    Else
	        If aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_EMAIL Then
	            oValueNode.setAttribute "vld", "email"
	        ElseIf aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_NUMBER Then
	            oValueNode.setAttribute "vld", "number"
	        Else
	            oValueNode.setAttribute "vld", "none"
	        End If
	    End If

	    Set oValueNode = oTempNode.selectSingleNode("field[@col = 'dvid']")
	    If (oValueNode Is Nothing) Then
	        Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
	        oValueNode2.setAttribute "di", IDS_STYLE
	        oValueNode2.setAttribute "n", "Device"
	        oValueNode2.setAttribute "col", "dvid"
	    End If

	    Set oValueNode = oTempNode.selectSingleNode("field[@col = 'cb']")
	    If aDeviceTypeInfo(DEV_TYPE_EDIT_DELIVERY_WINDOW) = "1" Then
	        If (oValueNode Is Nothing) Then
	            Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_DELIVERY_WINDOW
	            oValueNode2.setAttribute "n", "DeliveryWindow"
	            oValueNode2.setAttribute "col", "cb"
	        End If
	    Else
	        If Not (oValueNode Is Nothing) Then
	            oTempNode.removeChild(oValueNode)
	        End If
	    End If

	    Set oValueNode = oTempNode.selectSingleNode("field[@col = 'pwd']")
	    If aDeviceTypeInfo(DEV_TYPE_EDIT_PIN) = "1" Then
	        If (oValueNode Is Nothing) Then
	            Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_PIN
	            oValueNode2.setAttribute "n", "PIN"
	            oValueNode2.setAttribute "col", "pwd"
	            oValueNode2.setAttribute "size", "15"
	            oValueNode2.setAttribute "type", "password"
	            oValueNode2.setAttribute "vld", "number"

	            Set oValueNode2 = oTempNode.appendChild(oDeviceTypesDOM.createElement("field"))
                oValueNode2.setAttribute "di", IDS_PASSWORD
                oValueNode2.setAttribute "n", "ConfirmPIN"
	            oValueNode2.setAttribute "col", "cpwd"
	            oValueNode2.setAttribute "size", "15"
	            oValueNode2.setAttribute "type", "password"
	            oValueNode2.setAttribute "vld", "number"
	        End If
	    Else
	        If Not (oValueNode Is Nothing) Then
	            oTempNode.removeChild(oValueNode)
	        End If
	        Set oValueNode = oTempNode.selectSingleNode("field[@col = 'cpwd']")
            If Not (oValueNode Is Nothing) Then
	            oTempNode.removeChild(oValueNode)
	        End If
	    End If
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = cu_UpdateDeviceTypeDefinitions(aDeviceTypeInfo(DEV_TYPE_ID), oDeviceType.xml)
    End If

    If lErrNumber = NO_ERR Then
        sDeviceTypesXML = oDeviceTypesDOM.xml
    End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing
	Set oTempNode = Nothing
	Set oValueNode = Nothing
	Set oValueNode2 = Nothing

	EditDeviceType = lErrNumber
	Err.Clear
End Function

Function RemoveFoldersFromDT(oRequest, sDeviceTypeID, sDeviceTypesXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oDeviceType
	Dim oTempNode
	Dim i
	Dim oFolderNode

	lErrNumber = NO_ERR

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RemoveFoldersFromDT", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & sDeviceTypeID & "']")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RemoveFoldersFromDT", "", "Error loading devicetype node", LogLevelError)
	    End If
	End If

    If lErrNumber = NO_ERR Then
        If oRequest("SDTFolder").Count > 0 Then
            Set oTempNode = oDeviceType.selectSingleNode("temp/dfs")
            For i=1 to oRequest("SDTFolder").Count
                oTempNode.removeChild(oTempNode.selectSingleNode("f[@id = '" & Left(CStr(oRequest("SDTFolder")(i)), Instr(1, CStr(oRequest("SDTFolder")(i)), ";") - 1) & "']"))
            Next
        End If
    End If

    If lErrNumber = NO_ERR Then
        sDeviceTypesXML = oDeviceTypesDOM.xml
    End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing
	Set oTempNode = Nothing
	Set oFolderNode = Nothing

	RemoveFoldersFromDT = lErrNumber
	Err.Clear
End Function

Function AddFoldersToDT(oRequest, sDeviceTypeID, sDeviceTypesXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oDeviceType
	Dim oTempNode
	Dim i
	Dim oFolderNode

	lErrNumber = NO_ERR

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "AddFoldersToDT", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & sDeviceTypeID & "']")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "AddFoldersToDT", "", "Error loading devicetype node", LogLevelError)
	    End If
	End If

    If lErrNumber = NO_ERR Then
        If oRequest("ADTFolder").Count > 0 Then
            Set oTempNode = oDeviceType.selectSingleNode("temp")
            If (oTempNode Is Nothing) Then
                Set oTempNode = oDeviceType.appendChild(oDeviceTypesDOM.createElement("temp"))
                Set oTempNode = oTempNode.appendChild(oDeviceTypesDOM.createElement("dfs"))
            Else
                Set oTempNode = oTempNode.selectSingleNode("dfs")
            End If
            For i=1 to oRequest("ADTFolder").Count
                Set oFolderNode = oTempNode.appendChild(oDeviceTypesDOM.createElement("f"))
                oFolderNode.setAttribute "id", Left(CStr(oRequest("ADTFolder")(i)), Instr(1, CStr(oRequest("ADTFolder")(i)), ";") - 1)
                oFolderNode.setAttribute "n", Right(CStr(oRequest("ADTFolder")(i)), Len(CStr(oRequest("ADTFolder")(i))) - Instr(1, CStr(oRequest("ADTFolder")(i)), ";"))
            Next
        End If
    End If

    If lErrNumber = NO_ERR Then
        sDeviceTypesXML = oDeviceTypesDOM.xml
    End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing
	Set oTempNode = Nothing
	Set oFolderNode = Nothing

	AddFoldersToDT = lErrNumber
	Err.Clear
End Function

Function SaveDTFolders(sDeviceTypeID, sDeviceTypesXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oDeviceType
	Dim oDFSNode

	lErrNumber = NO_ERR

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "SaveDTFolders", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & sDeviceTypeID & "']")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "SaveDTFolders", "", "Error retrieving devicetype node", LogLevelError)
	    End If
	End If

    If lErrNumber = NO_ERR Then
        Set oDFSNode = oDeviceType.selectSingleNode("dfs")
        If Not (oDFSNode Is Nothing) Then
            oDeviceType.removeChild(oDFSNode)
        End If
        oDeviceType.appendChild(oDeviceType.selectSingleNode("temp/dfs"))
        oDeviceType.removeChild(oDeviceType.selectSingleNode("temp"))
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = cu_UpdateDeviceTypeDefinitions(sDeviceTypeID, oDeviceType.xml)
    End If

    If lErrNumber = NO_ERR Then
        sDeviceTypesXML = oDeviceTypesDOM.xml
    End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing
	Set oDFSNode = Nothing

	SaveDTFolders = lErrNumber
	Err.Clear
End Function

Function LoadTempDFS(sDeviceTypesXML, sDeviceTypeID)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oDeviceType
	Dim oDFSNode
	Dim oTempNode

	lErrNumber = NO_ERR

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "LoadTempDFS", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & sDeviceTypeID & "']")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "LoadTempDFS", "", "Error retrieving devicetype node", LogLevelError)
	    End If
	End If

    If lErrNumber = NO_ERR Then
        Set oDFSNode = oDeviceType.selectSingleNode("dfs")
        If Not (oDFSNode Is Nothing) Then
            Set oTempNode = oDeviceType.selectSingleNode("temp")
            If (oTempNode Is Nothing) Then
                Set oTempNode = oDeviceType.appendChild(oDeviceTypesDOM.createElement("temp"))
                oTempNode.appendChild(oDFSNode.cloneNode(True))
                'oDeviceType.appendChild(oDFSNode)
                sDeviceTypesXML = oDeviceTypesDOM.xml
            End If
        End If
    End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing
	Set oDFSNode = Nothing
	Set oTempNode = Nothing

	LoadTempDFS = lErrNumber
	Err.Clear
End Function

Function DeleteTempDFS(sDeviceTypesXML, sDeviceTypeID)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oDeviceType
	Dim oTempNode

	lErrNumber = NO_ERR

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "DeleteTempDFS", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & sDeviceTypeID & "']")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "DeleteTempDFS", "", "Error retrieving devicetype node", LogLevelError)
	    End If
	End If

    If lErrNumber = NO_ERR Then
        Set oTempNode = oDeviceType.selectSingleNode("temp")
        If Not (oTempNode Is Nothing) Then
            oDeviceType.removeChild(oTempNode)
        End If
    End If

	If lErrNumber = NO_ERR Then
	    sDeviceTypesXML = oDeviceTypesDOM.xml
	End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing
	Set oTempNode = Nothing

	DeleteTempDFS = lErrNumber
	Err.Clear
End Function

Function GetDeviceTypeNames(sDeviceTypesXML, aNames)
'*******************************************************************************
'Purpose: Returns the list of current device types Names
'Inputs:  sDeviceTypesXML: The XML of the device types.
'Outputs: aNames: the list of names
'*******************************************************************************
Const PROCEDURE_NAME = "GetDeviceTypeNames"
Dim lErrNumber
Dim oDeviceTypes
Dim oDeviceTypesDOM
Dim i, lCount

    On Error Resume Next
    lErrNumber = NO_ERR

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sDeviceTypesXML, oDeviceTypesDOM)
    If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sDeviceTypesXML", LogLevelTrace)

    If lErrNumber = NO_ERR Then
	    Set oDeviceTypes = oDeviceTypesDOM.selectNodes("/devicetypes/devicetype")
	    lCount = oDeviceTypes.length

	    If lCount > 0 Then
	        ReDim aNames(lCount - 1)

            For i = 0 To lCount - 1
                aNames(i) = oDeviceTypes(i).selectSingleNode("name").text
            Next

        End If
    End If

    Set oDeviceTypes = Nothing
    Set oDeviceTypesDOM = Nothing

    GetDeviceTypeNames = lErrNumber
    Err.Clear

End Function

Function AddNewDeviceType(aDeviceTypeInfo, sDeviceTypesXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oTempNode
	Dim oValueNode
	Dim oValueNode2
	Dim aNames
	Dim sNewName
	Dim sSiteId

	lErrNumber = NO_ERR
	aDeviceTypeInfo(DEV_TYPE_ID) = GetGUID()
	sSiteId = Application.Value("SITE_ID")

	Call GetDeviceTypeNames(sDeviceTypesXML, aNames)
	sNewName = GetNewName(aNames, aDeviceTypeInfo(DEV_TYPE_NAME))

	'The name already existed:
	If sNewName <> aDeviceTypeInfo(DEV_TYPE_NAME) Then
	    lErrNumber = ERR_INVALID_NAME
    End If

    If lErrNumber = NO_ERR Then
	    lErrNumber = cu_CreateDeviceType(sSiteId, aDeviceTypeInfo(DEV_TYPE_ID), aDeviceTypeInfo(DEV_TYPE_NAME))
	    If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "AddNewDeviceType", "", "Error calling cu_CreateDeviceType", LogLevelTrace)
	End If

	If lErrNumber = NO_ERR Then
        Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	    oDeviceTypesDOM.async = False
	    If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	        lErrNumber = ERR_XML_LOAD_FAILED
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "AddNewDeviceType", "", "Error loading sDeviceTypesXML", LogLevelError)
	    Else
	        Set oTempNode = oDeviceTypesDOM.selectSingleNode("/devicetypes/temp")
	        If Not (oTempNode Is Nothing) Then
	            oDeviceTypesDOM.selectSingleNode("/devicetypes").removeChild(oTempNode)
	        End If

	        Set oTempNode = oDeviceTypesDOM.selectSingleNode("/devicetypes").appendChild(oDeviceTypesDOM.createElement("temp"))
	        Set oTempNode = oTempNode.appendChild(oDeviceTypesDOM.createElement("devicetype"))

	        Set oValueNode = oTempNode.appendChild(oDeviceTypesDOM.createElement("icon"))
	        If Len(aDeviceTypeInfo(DEV_TYPE_SMALL_ICON)) > 0 Then
	            oValueNode.text = aDeviceTypeInfo(DEV_TYPE_SMALL_ICON)
	        Else
	            oValueNode.text = IMAGE_EMPTY
	        End If
	        oValueNode.setAttribute "height", "20"
	        oValueNode.setAttribute "width", "20"

	        Set oValueNode = oTempNode.appendChild(oDeviceTypesDOM.createElement("largeicon"))
	        If Len(aDeviceTypeInfo(DEV_TYPE_LARGE_ICON)) > 0 Then
	            oValueNode.text = aDeviceTypeInfo(DEV_TYPE_LARGE_ICON)
	        Else
	            oValueNode.text = IMAGE_EMPTY
	        End If
	        oValueNode.setAttribute "height", "50"
	        oValueNode.setAttribute "width", "50"

	        Set oValueNode = oTempNode.appendChild(oDeviceTypesDOM.createElement("devicetypeID"))
	        oValueNode.text = aDeviceTypeInfo(DEV_TYPE_ID)

	        Set oValueNode = oTempNode.appendChild(oDeviceTypesDOM.createElement("name"))
	        oValueNode.text = aDeviceTypeInfo(DEV_TYPE_NAME)

	        oTempNode.appendChild(oDeviceTypesDOM.createElement("devices"))

	        'Display fields
	        Set oValueNode = oTempNode.appendChild(oDeviceTypesDOM.createElement("displayfields"))
	        If aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME) = "1" Then
	            Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_NAME
	            oValueNode2.setAttribute "n", "Name"
	            oValueNode2.setAttribute "col", "n"
	        End If
	        If aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE) = "1" Then
	            Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_ADDRESS
	            oValueNode2.setAttribute "n", "Address"
	            oValueNode2.setAttribute "col", "v"
	        End If
	        If aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE) = "1" Then
	            Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_STYLE
	            oValueNode2.setAttribute "n", "Style"
	            oValueNode2.setAttribute "col", "dvid"
	        End If
	        If aDeviceTypeInfo(DEV_TYPE_DISPLAY_DELIVERY_WINDOW) = "1" Then
	            Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_DELIVERY_WINDOW
	            oValueNode2.setAttribute "n", "DeliveryWindow"
	            oValueNode2.setAttribute "col", "cb"
	        End If

            'Edit fields
	        Set oValueNode = oTempNode.appendChild(oDeviceTypesDOM.createElement("editfields"))
	        Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	        oValueNode2.setAttribute "di", IDS_NAME
	        oValueNode2.setAttribute "n", "AddressName"
	        oValueNode2.setAttribute "col", "n"
	        oValueNode2.setAttribute "size", "15"
	        oValueNode2.setAttribute "type", "text"
	        oValueNode2.setAttribute "vld", "name"

	        Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	        oValueNode2.setAttribute "di", IDS_ADDRESS
	        oValueNode2.setAttribute "n", "PhysicalAddress"
	        oValueNode2.setAttribute "col", "v"
	        oValueNode2.setAttribute "size", "15"
	        oValueNode2.setAttribute "type", "text"
	        If aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_EMAIL Then
	            oValueNode2.setAttribute "vld", "email"
	        ElseIf aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_NUMBER Then
	            oValueNode2.setAttribute "vld", "number"
	        Else
	            oValueNode2.setAttribute "vld", "none"
	        End If

	        Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	        oValueNode2.setAttribute "di", IDS_STYLE
	        oValueNode2.setAttribute "n", "Device"
	        oValueNode2.setAttribute "col", "dvid"

	        If aDeviceTypeInfo(DEV_TYPE_EDIT_DELIVERY_WINDOW) = "1" Then
	            Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_DELIVERY_WINDOW
	            oValueNode2.setAttribute "n", "DeliveryWindow"
	            oValueNode2.setAttribute "col", "cb"
	        End If
	        If aDeviceTypeInfo(DEV_TYPE_EDIT_PIN) = "1" Then
	            Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_PIN
	            oValueNode2.setAttribute "n", "PIN"
	            oValueNode2.setAttribute "col", "pwd"
	            oValueNode2.setAttribute "size", "15"
	            oValueNode2.setAttribute "type", "password"
	            oValueNode2.setAttribute "vld", "number"
	            Set oValueNode2 = oValueNode.appendChild(oDeviceTypesDOM.createElement("field"))
	            oValueNode2.setAttribute "di", IDS_PASSWORD
	            oValueNode2.setAttribute "n", "ConfirmPIN"
	            oValueNode2.setAttribute "col", "cpwd"
	            oValueNode2.setAttribute "size", "15"
                oValueNode2.setAttribute "type", "password"
	            oValueNode2.setAttribute "vld", "number"
	        End If
        End If
    End If

    If lErrNumber = NO_ERR Then
        lErrNumber = cu_CreateDeviceTypeDefinitions(sSiteId, aDeviceTypeInfo(DEV_TYPE_ID), oDeviceTypesDOM.selectSingleNode("/devicetypes/temp/devicetype").xml)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "AddNewDeviceType", "", "Error calling cu_CreateDeviceTypeDefinitions", LogLevelTrace)
        Else
            oDeviceTypesDOM.selectSingleNode("/devicetypes").appendChild(oDeviceTypesDOM.selectSingleNode("/devicetypes/temp/devicetype"))
	        Set oTempNode = oDeviceTypesDOM.selectSingleNode("/devicetypes/temp")
	        If Not (oTempNode Is Nothing) Then
	            oDeviceTypesDOM.selectSingleNode("/devicetypes").removeChild(oTempNode)
	        End If
	    End If
    End If

    If lErrNumber = NO_ERR Then
        sDeviceTypesXML = oDeviceTypesDOM.xml
    End If

    Set oDeviceTypesDOM = Nothing
	Set oTempNode = Nothing
	Set oValueNode = Nothing
	Set oValueNode2 = Nothing

	AddNewDeviceType = lErrNumber
	Err.Clear
End Function

Function DeleteDeviceType(sDeviceTypesXML, sDTID)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM

	lErrNumber = NO_ERR

    lErrNumber = cu_DeleteDeviceType(sDTID)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "DeleteDeviceType", "", "Error calling cu_DeleteDeviceType", LogLevelTrace)
    End If

    If lErrNumber = NO_ERR Then
        Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	    oDeviceTypesDOM.async = False
	    If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	        lErrNumber = ERR_XML_LOAD_FAILED
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "DeleteDeviceType", "", "Error loading sDeviceTypesXML", LogLevelError)
	    Else
	        oDeviceTypesDOM.selectSingleNode("/devicetypes").removeChild(oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & sDTID & "']"))
	        If Err.number <> NO_ERR Then
	            lErrNumber = Err.number
	            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "DeleteDeviceType", "", "Error deleting devicetype node", LogLevelError)
	        End If
	    End If
	End If

    If lErrNumber = NO_ERR Then
        sDeviceTypesXML = oDeviceTypesDOM.xml
    End If

	Set oDeviceTypesDOM = Nothing

	DeleteDeviceType = lErrNumber
	Err.Clear
End Function

Function GetVariablesFromXML_DTFolders(aDeviceTypeInfo, sDeviceTypesXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oDeviceType

	lErrNumber = NO_ERR

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "GetVariablesFromXML_DTFolders", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & aDeviceTypeInfo(DEV_TYPE_ID) & "']")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "GetVariablesFromXML_DTFolders", "", "Error retrieving devicetype node", LogLevelError)
	    End If
	End If

    If lErrNumber = NO_ERR Then
        aDeviceTypeInfo(DEV_TYPE_NAME) = oDeviceType.selectSingleNode("name").text
    End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing

	GetVariablesFromXML_DTFolders = lErrNumber
	Err.Clear
End Function


Function GetNewDeviceTypeInfo(sDeviceTypesXML, aDeviceTypeInfo)
'********************************************************
'*Purpose: Reads the default information of a device type
'*Inputs:
'*Outputs: aDeviceTypeInfo: the information with the default values for a new device type
'********************************************************
Const PROCEDURE_NAME = "GetNewDeviceTypeInfo"
Dim lErrNumber
Dim aNames

    On Error Resume Next
    lErrNumber = NO_ERR

    Call GetDeviceTypeNames(sDeviceTypesXML, aNames)

    Redim aDeviceTypeInfo(MAX_DEV_TYPE_INFO)
    aDeviceTypeInfo(DEV_TYPE_NAME) = GetNewName(aNames, "New Device Type")
    aDeviceTypeInfo(DEV_TYPE_LARGE_ICON) = ""
    aDeviceTypeInfo(DEV_TYPE_SMALL_ICON) = ""
    aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_EMAIL
    aDeviceTypeInfo(DEV_TYPE_ACTION) = "new"

	aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME) = "1"
	aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE) = "1"
	aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE) = "1"
	aDeviceTypeInfo(DEV_TYPE_DISPLAY_DELIVERY_WINDOW) = "0"

	aDeviceTypeInfo(DEV_TYPE_EDIT_PIN) = "0"
	aDeviceTypeInfo(DEV_TYPE_EDIT_DELIVERY_WINDOW) = "0"


    Set oDeviceTypes = Nothing
    Erase aNames

	GetNewDeviceTypeInfo = lErrNumber
	Err.Clear

End Function

Function GetDeviceTypeInfo(aDeviceTypeInfo, sDeviceTypesXML)
'*******************************************************************************
'Purpose: Sets the DeviceType Inform from the device types XML
'Inputs:  sDeviceTypesXML
'Outputs: aDeviceTypeInfo
'*******************************************************************************
Const PROCEDURE_NAME = "GetDeviceTypeInfo"
Dim lErrNumber
Dim oDeviceTypesDOM
Dim oDeviceType

	On Error Resume Next
	lErrNumber = NO_ERR

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sDeviceTypesXML, oDeviceTypesDOM)
    If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sDeviceTypesXML", LogLevelTrace)

    If lErrNumber = NO_ERR Then
	    Set oDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & aDeviceTypeInfo(DEV_TYPE_ID) & "']")

        If Not oDeviceType Is Nothing Then
            aDeviceTypeInfo(DEV_TYPE_NAME) =oDeviceType.selectSingleNode("name").text
            aDeviceTypeInfo(DEV_TYPE_LARGE_ICON) = oDeviceType.selectSingleNode("largeicon").text
            aDeviceTypeInfo(DEV_TYPE_SMALL_ICON) = oDeviceType.selectSingleNode("icon").text

            If oDeviceType.selectSingleNode("editfields/field[@col = 'v']").getAttribute("vld") = "email" Then
                aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_EMAIL
            ElseIf oDeviceType.selectSingleNode("editfields/field[@col = 'v']").getAttribute("vld") = "number" Then
                aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_NUMBER
            Else
                aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_NONE
            End If

            If Not (oDeviceType.selectSingleNode("displayfields/field[@col = 'n']") Is Nothing) Then
                aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME) = "1"
            Else
                aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME) = "0"
            End If

            If Not (oDeviceType.selectSingleNode("displayfields/field[@col = 'v']") Is Nothing) Then
                aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE) = "1"
            Else
                aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE) = "0"
            End If

            If Not (oDeviceType.selectSingleNode("displayfields/field[@col = 'dvid']") Is Nothing) Then
                aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE) = "1"
            Else
                aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE) = "0"
            End If

            If Not (oDeviceType.selectSingleNode("displayfields/field[@col = 'cb']") Is Nothing) Then
                aDeviceTypeInfo(DEV_TYPE_DISPLAY_DELIVERY_WINDOW) = "1"
            Else
                aDeviceTypeInfo(DEV_TYPE_DISPLAY_DELIVERY_WINDOW) = "0"
            End If

            If Not (oDeviceType.selectSingleNode("editfields/field[@col = 'pwd']") Is Nothing) Then
                aDeviceTypeInfo(DEV_TYPE_EDIT_PIN) = "1"
            Else
                aDeviceTypeInfo(DEV_TYPE_EDIT_PIN) = "0"
            End If

            If Not (oDeviceType.selectSingleNode("editfields/field[@col = 'cb']") Is Nothing) Then
                aDeviceTypeInfo(DEV_TYPE_EDIT_DELIVERY_WINDOW) = "1"
            Else
                aDeviceTypeInfo(DEV_TYPE_EDIT_DELIVERY_WINDOW) = "0"
            End If
        End If
    End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing

	GetVariablesFromXML_Edit = lErrNumber
	Err.Clear
End Function

Function ReadDeviceTypesXML(sDeviceTypesXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
Const PROCEDURE_NAME = "ReadDeviceTypesXML"
Dim lErrNumber
Dim oDeviceTypesDOM
Dim oDTDOM
Dim oDTs
Dim oCurrentDT
Dim oDTD_DOM
Dim sGetDeviceTypesXML
Dim sGetDeviceTypeDefinitionsXML
Dim sFilePath
Dim sSiteID

	On Error Resume Next
	lErrNumber = NO_ERR

	sDeviceTypesXML = ""
	sFilePath = GetFilePath()
	sSiteId = Application.Value("SITE_ID")

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False

	'Try to read it from cache file, if not found, read it from MD:
	If oDeviceTypesDOM.load(sFilePath & "deviceTypes_" & sSiteID & ".xml") = False Then

        lErrNumber = cu_GetDeviceTypes(sGetDeviceTypesXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error retrieving sDeviceTypesXML", LogLevelTrace)
        Else
            lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetDeviceTypesXML, oDTDOM)
            If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sGetDeviceTypesXML", LogLevelTrace)
        End If

        If lErrNumber = NO_ERR Then
            lErrNumber = cu_GetDeviceTypeDefinitions(sGetDeviceTypesXML, sGetDeviceTypeDefinitionsXML)
            If lErrNumber <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error retrieving sGetDeviceTypeDefinitionsXML", LogLevelTrace)
            Else
                lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetDeviceTypeDefinitionsXML, oDTD_DOM)
                If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sGetDeviceTypeDefinitionsXML", LogLevelTrace)
            End If
        End If

        If lErr = NO_ERR Then
            oDeviceTypesDOM.loadXML("<devicetypes></devicetypes>")
            Set oDTs = oDTDOM.selectNodes("/mi/in/oi[@tp = '" & TYPE_DEVICE_TYPE & "']")
            For Each oCurrentDT in oDTs
                oDeviceTypesDOM.selectSingleNode("/devicetypes").appendChild(oDTD_DOM.selectSingleNode("//oi[@id = '" & oCurrentDT.getAttribute("id") & "']/prs/pr/devicetype"))
            Next
        End If

        If lErrNumber = NO_ERR Then
            oDeviceTypesDOM.save(sFilePath & "deviceTypes_" & sSiteId & ".xml")
        End If

	End If

	If lErrNumber = NO_ERR Then
	    sDeviceTypesXML = oDeviceTypesDOM.xml
	End If

	Set oDeviceTypesDOM = Nothing
	Set oDTDOM = Nothing
	Set oDTs = Nothing
	Set oCurrentDT = Nothing
	Set oDTD_DOM = Nothing

	ReadDeviceTypesXML = lErrNumber
	Err.Clear

End Function

Function WriteDeviceTypesXML(sDeviceTypesXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
Dim lErrNumber
Dim oDeviceTypesDOM
Dim sFilePath
Dim sSiteId

	On Error Resume Next
	lErrNumber = NO_ERR

	sFilePath = GetFilePath()
	sSiteId = Application.Value("SITE_ID")

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "WriteDeviceTypesXML", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
        oDeviceTypesDOM.save(sFilePath & "deviceTypes_" & sSiteId & ".xml")
    End If

	Set oDeviceTypesDOM = Nothing

	WriteDeviceTypesXML = lErrNumber
	Err.Clear
End Function

Function RenderExistingDeviceTypes(sDeviceTypesXML, sAction, sDeviceTypeID)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oDeviceTypes
	Dim oCurrentDevice
	Dim aNewDeviceTypeInfo

	lErrNumber = NO_ERR

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderExistingDeviceTypes", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oDeviceTypes = oDeviceTypesDOM.selectNodes("/devicetypes/devicetype")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderExistingDeviceTypes", "", "Error retrieving devicetype nodes", LogLevelError)
	    End If
	End If

    If lErr = NO_ERR Then
        lErr = GetNewDeviceTypeInfo(sDeviceTypesXML, aNewDeviceTypeInfo)
    End If

    If lErrNumber = NO_ERR Then
        If oDeviceTypes.length > 0 Then
        For Each oCurrentDevice in oDeviceTypes
            Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
            Response.Write "<TR>"
            Response.Write "<TD WIDTH=""1%"" VALIGN=TOP>"

            If (sAction = "rename") And (sDeviceTypeID = oCurrentDevice.selectSingleNode("devicetypeID").text) Then
                Response.Write "<IMG SRC=""../images/redtri.gif"" HEIGHT=""11"" WIDTH=""17"" ALT="""" BORDER=""0"" />"
            Else
                Response.Write "<IMG SRC=""../images/arrow_right.gif"" HEIGHT=""13"" WIDTH=""13"" ALT="""" BORDER=""0"" />"
            End If

            Response.Write "</TD>"
            Response.Write "<TD WIDTH=""1%"" VALIGN=TOP>"
                Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
                Response.Write "<TR>"

                If (sAction = "rename") And (sDeviceTypeID = oCurrentDevice.selectSingleNode("devicetypeID").text) Then
                    Response.Write "<TD BGCOLOR=""#cccccc"" COLSPAN=2>"
                        Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
                        Response.Write "<form METHOD=""POST"" NAME=""FormDevName"" ACTION=""deviceTypes.asp"">"
                        Response.Write "<INPUT TYPE=HIDDEN NAME=""dtID"" VALUE=""" & sDeviceTypeID & """ >"
                        Response.Write "<TR>"
                        Response.Write "<TD BGCOLOR=""#ffffff"" NOWRAP>"
                        Response.Write "<INPUT CLASS=""textBoxClass"" NAME=""DTName"" TYPE=""TEXT"" VALUE=""" & Server.HTMLEncode(oCurrentDevice.selectSingleNode("name").text) & """ /><BR />"
                        Response.Write "<INPUT CLASS=""buttonClass"" TYPE=SUBMIT NAME=""RenameDT"" onClick=""return validateForm();"" VALUE=""" & asDescriptors(498) & """> <INPUT CLASS=""buttonClass"" NAME=""CANCEL"" TYPE=SUBMIT VALUE=""" & asDescriptors(120) & """>" 'Descriptor: Rename, Cancel
                        Response.Write "</TD>"
                        Response.Write "</TR>"
                        Response.Write "</form>"
                        Response.Write "</TABLE>"
                    Response.Write "</TD>"
                Else
                    Response.Write "<TD COLSPAN=2 VALIGN=TOP><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & Server.HTMLEncode(oCurrentDevice.selectSingleNode("name").text) & "</b></font></TD>"
                End If

                Response.Write "</TR>"
                Response.Write "<TR>"
                Response.Write "<TD>"
                Response.Write "<IMG SRC="""
                If Left(oCurrentDevice.selectSingleNode("largeicon").text, 4) <> "http" Then
                    Response.Write "../"
                End If
                Response.Write Server.HTMLEncode(oCurrentDevice.selectSingleNode("largeicon").text) & """ HEIGHT=""" & oCurrentDevice.selectSingleNode("largeicon").getAttribute("height") & """ WIDTH=""" & oCurrentDevice.selectSingleNode("largeicon").getAttribute("width") & """ ALT="""" BORDER=""0"" />"
                Response.Write "</TD>"
                Response.Write "<TD VALIGN=TOP>"
                    Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
                    Response.Write "<TR>"
                    Response.Write "<TD WIDTH=""1%"">"
                    Response.Write "<IMG SRC=""../images/bullet.gif"" HEIGHT=""8"" WIDTH=""3"" ALT="""" BORDER=""0"" />"
                    Response.Write "</TD>"
                    Response.Write "<TD WIDTH=""99%"">"
                    Response.Write "<A HREF=""deviceTypes.asp?dtID=" & oCurrentDevice.selectSingleNode("devicetypeID").text & "&action=rename""><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(494) & "</font></A>" 'Descriptor: rename ...
                    Response.Write "</TD>"
                    Response.Write "</TR>"
                    Response.Write "<TR>"
                    Response.Write "<TD WIDTH=""1%"">"
                    Response.Write "<IMG SRC=""../images/bullet.gif"" HEIGHT=""8"" WIDTH=""3"" ALT="""" BORDER=""0"" />"
                    Response.Write "</TD>"
                    Response.Write "<TD WIDTH=""99%"">"
                    Response.Write "<A HREF=""editDeviceType.asp?action=edit&dtID=" & oCurrentDevice.selectSingleNode("devicetypeID").text & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(495) & "</font></A>" 'Descriptor: definition ...
                    Response.Write "</TD>"
                    Response.Write "</TR>"
                    Response.Write "<TR>"
                    Response.Write "<TD WIDTH=""1%"">"
                    Response.Write "<IMG SRC=""../images/bullet.gif"" HEIGHT=""8"" WIDTH=""3"" ALT="""" BORDER=""0"" />"
                    Response.Write "</TD>"
                    Response.Write "<TD WIDTH=""99%"">"
                    Response.Write "<A HREF=""deleteDeviceType.asp?n=" & Server.URLEncode(oCurrentDevice.selectSingleNode("name").text) & "&id=" & oCurrentDevice.selectSingleNode("devicetypeID").text& "&tp=" & TYPE_DEVICE_TYPE  & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(249) & "</font></A>" 'Descriptor: Delete
                    Response.Write "</TD>"
                    Response.Write "</TR>"
                    Response.Write "</TABLE>"
                Response.Write "</TD>"
                Response.Write "</TR>"
                Response.Write "</TABLE>"
            Response.Write "</TD>"
            Response.Write "<TD WIDTH=""98%"" VALIGN=TOP>"
                Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
                Response.Write "<TR>"
                Response.Write "<TD BGCOLOR=""#000000"">"
                    Response.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH=""100%"">"
                    Response.Write "<TR>"
                    Response.Write "<TD BGCOLOR=""#cccccc"">"
                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(496) & "</b></font>" 'Descriptor: Device Folders
                    Response.Write "</TD>"
                    Response.Write "</TR>"
                    Response.Write "</TABLE>"
                Response.Write "</TD>"
                Response.Write "</TR>"
                Response.Write "<TR>"
                Response.Write "<TD>"
                    Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
                    Response.Write "<TR>"
                    Response.Write "<TD WIDTH=""99%"">"

                    Call RenderDeviceFolders(oDeviceTypesDOM, oCurrentDevice.selectSingleNode("devicetypeID").text)

                    Response.Write "</TD>"
                    Response.Write "<TD WIDTH=""1%"" VALIGN=TOP NOWRAP>"
                    Response.Write "<A HREF=""deviceTypeFolders.asp?dtID=" & oCurrentDevice.selectSingleNode("devicetypeID").text & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(497) & "</font></A>" 'Descriptor: edit device folders ...
                    Response.Write "</TD>"
                    Response.Write "</TR>"
                    Response.Write "</TABLE>"
                Response.Write "</TD>"
                Response.Write "</TR>"
                Response.Write "</TABLE>"
            Response.Write "</TD>"
            Response.Write "</TR>"
            Response.Write "</TABLE>"
        Next
        Else
            Response.Write "<BR />"
            Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
            Response.Write "<TR>"
            Response.Write "<TD><IMG SRC=""../images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""13"" ALT="""" BORDER=""0"" /></TD>"
            Response.Write "<TD>"
            Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """ color=""#cc0000""><b>" & asDescriptors(526) & "</b></font>" 'Descriptor: No device types have been created.
            Response.Write "</TD>"
            Response.Write "</TR>"
            Response.Write "</TABLE>"
        End If

        Response.Write "<BR />"
        Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
        Response.Write "<TR>"
        Response.Write "<TD WIDTH=""1%"">"
        Response.Write "<IMG SRC=""../images/arrow_right.gif"" HEIGHT=""13"" WIDTH=""13"" ALT="""" BORDER=""0"" />"
        Response.Write "</TD>"
        Response.Write "<TD WIDTH=""99%"">"
        Response.Write "<A HREF=""editDeviceType.asp?" & CreateRequestForDeviceType(aNewDeviceTypeInfo) & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & asDescriptors(523) & "</b></font></A>" 'Descriptor: Create new Device Type
        Response.Write "</TD>"
        Response.Write "</TR>"
        Response.Write "</TABLE>"
    End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceTypes = Nothing
	Set oCurrentDevice = Nothing

	RenderExistingDeviceTypes = lErrNumber
	Err.Clear
End Function

Function RenderDeviceFolders(oDeviceTypesDOM, sDeviceTypeID)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'TO DO: add error handling!
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oFolders
	Dim iNumFolders
	Dim i

	lErrNumber = NO_ERR

	Set oFolders = oDeviceTypesDOM.selectNodes("/devicetypes/devicetype[devicetypeID = '" & sDeviceTypeID & "']/dfs/f")
	If oFolders.length > 0 Then
	    iNumFolders = oFolders.length
	    For i=0 to (iNumFolders - 1)
	        If (i+1) Mod 2 = 1 Then
	    		Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
	    		Response.Write "<TR><TD VALIGN=TOP WIDTH=""50%"">"
	    	Else
	    		Response.Write "<TD VALIGN=TOP WIDTH=""50%"">"
	    	End If

	    	Response.Write "<IMG SRC=""../images/folder.gif"" HEIGHT=""16"" WIDTH=""16"" ALT="""" BORDER=""0"" />"
	    	Response.Write "&nbsp;"
	    	Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oFolders.item(i).getAttribute("n") & "</font>"

	    	If (i+1) Mod 2 = 1 Then
	    		Response.Write "</TD>"
	    		If i = (iNumFolders-1) Then
	    			Response.Write "<TD></TD></TR></TABLE>"
	    		End If
	    	Else
	    		Response.Write "</TD></TR></TABLE>"
	    	End If
        Next
	End If

	Set oFolders = Nothing

	RenderDeviceFolders = lErrNumber
	Err.Clear
End Function

Function RenameDeviceType(sRenameDeviceTypeID, sRenameDeviceTypeName, sDeviceTypesXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'TO DO: add error handling!
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oRenameDT
	Dim aNames
	Dim sNewName

	lErrNumber = NO_ERR

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenameDeviceType", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oRenameDT = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & sRenameDeviceTypeID & "']")
	    If Not (oRenameDT Is Nothing) Then
	        If oRenameDT.selectSingleNode("name").text <> sRenameDeviceTypeName Then

	            Call GetDeviceTypeNames(sDeviceTypesXML, aNames)
	            sNewName = GetNewName(aNames, sRenameDeviceTypeName)

	            'The name already existed:
	            If sNewName <> sRenameDeviceTypeName Then
	                lErrNumber = ERR_INVALID_NAME
                End If

                If lErrNumber = NO_ERR Then
                    oRenameDT.selectSingleNode("name").text = sRenameDeviceTypeName
                    lErrNumber = cu_UpdateDeviceTypeDefinitions(sRenameDeviceTypeID, oRenameDT.xml)
                End If
            End If
	    End If
	End If

	sDeviceTypesXML = oDeviceTypesDOM.xml

	Set oDeviceTypesDOM = Nothing
	Set oRenameDT = Nothing

	RenameDeviceType = lErrNumber
    Err.Clear
End Function

Function RenderDeviceTypeEditor(aDeviceTypeInfo)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'TO DO: add error handling!
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber

	lErrNumber = NO_ERR

	Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
	Response.Write "<TR>"
	Response.Write "<TD BGCOLOR=""#cccccc"">"
	    Response.Write "<TABLE BORDER=0 CELLPADDING=10 CELLSPACING=0>"
	    Response.Write "<TR>"
	    Response.Write "<TD BGCOLOR=""#ffffff"" VALIGN=TOP>"
	        Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
	        Response.Write "<TR>"
	        Response.Write "<TD VALIGN=TOP>"
	            Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
	            If aDeviceTypeInfo(DEV_TYPE_ACTION) = "new" Then
	                Response.Write "<TR>"
	                Response.Write "<TD COLSPAN=2>"
	                Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(513) & "</font><BR />" 'Descriptor: Device Type Name:
	                Response.Write "<INPUT TYPE=""TEXT"" NAME=""DTName"" SIZE=""20"" CLASS=""textBoxClass"" VALUE=""" & Server.HTMLEncode(aDeviceTypeInfo(DEV_TYPE_NAME)) & """ />"
	                Response.Write "<BR /><BR /></TD>"
	                Response.Write "</TR>"
	            End If
	            Response.Write "<TR>"
	            Response.Write "<TD>"
	            Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(499) & "</font><BR />" 'Descriptor: Large icon URL:
	            Response.Write "<INPUT TYPE=""TEXT"" SIZE=""30"" NAME=""DTLargeIcon"" CLASS=""textBoxClass"" "
	            If aDeviceTypeInfo(DEV_TYPE_LARGE_ICON) <> "" And StrComp(aDeviceTypeInfo(DEV_TYPE_LARGE_ICON), IMAGE_EMPTY) <> 0 Then
	                Response.Write "VALUE=""" & Server.HTMLEncode(aDeviceTypeInfo(DEV_TYPE_LARGE_ICON)) & """ "
	            End If
	            Response.Write "/>"
	            Response.Write "</TD>"
	            Response.Write "<TD ALIGN=CENTER>"
	            If aDeviceTypeInfo(DEV_TYPE_ACTION) = "edit" Then
	                Response.Write "<IMG SRC="""
	                If Left(aDeviceTypeInfo(DEV_TYPE_LARGE_ICON), 4) <> "http" Then
	                    Response.Write "../"
	                End If
	                Response.Write Server.HTMLEncode(aDeviceTypeInfo(DEV_TYPE_LARGE_ICON)) & """ ALT="""" BORDER=""0"" />"
	            End If
	            Response.Write "</TD>"
	            Response.Write "</TR>"
	            Response.Write "<TR>"
	            Response.Write "<TD>"
	            Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(500) & "</font><BR />" 'Descriptor: Small icon URL:
	            Response.Write "<INPUT TYPE=""TEXT"" SIZE=""30"" NAME=""DTSmallIcon"" CLASS=""textBoxClass"" "
	            If Len(aDeviceTypeInfo(DEV_TYPE_SMALL_ICON)) > 0 And StrComp(aDeviceTypeInfo(DEV_TYPE_SMALL_ICON), IMAGE_EMPTY) <> 0 Then
	                Response.Write "VALUE=""" & Server.HTMLEncode(aDeviceTypeInfo(DEV_TYPE_SMALL_ICON)) & """ "
	            End If
	            Response.Write "/>"
	            Response.Write "</TD>"
	            Response.Write "<TD ALIGN=CENTER>"
	            If aDeviceTypeInfo(DEV_TYPE_ACTION) = "edit" Then
	                Response.Write "<IMG SRC="""
	                If Left(aDeviceTypeInfo(DEV_TYPE_SMALL_ICON), 4) <> "http" Then
	                    Response.Write "../"
	                End If
	                Response.Write Server.HTMLEncode(aDeviceTypeInfo(DEV_TYPE_SMALL_ICON)) & """ ALT="""" BORDER=""0"" />"
	            End If
	            Response.Write "</TD>"
	            Response.Write "</TR>"
	            Response.Write "</TABLE>"

	            Response.Write "<BR />"
	            Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(501) & "</font><BR />" 'Descriptor: Address format:
	            Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
	            Response.Write "<TR>"
	            Response.Write "<TD BGCOLOR=""#cccccc"">"
	                Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
	                Response.Write "<TR>"
	                Response.Write "<TD BGCOLOR=""#ffffff"">"
	                    Response.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD VALIGN=TOP>"
	                    Response.Write "<INPUT TYPE=""RADIO"" NAME=""addrFormat"" VALUE=""" & S_DEVICE_VALIDATION_EMAIL & """ "
	                    If StrComp(aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT), S_DEVICE_VALIDATION_EMAIL, vbBinaryCompare) = 0 Or Len(aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT)) = 0 Then
	                        Response.Write "CHECKED "
	                    End If
	                    Response.Write "/>"
	                    Response.Write "</TD>"
	                    Response.Write "<TD VALIGN=TOP>"
	                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"
	                    Response.Write "<b>" & asDescriptors(502) & "</b><BR />" & asDescriptors(503) & "<BR />" & asDescriptors(504) 'Descriptor: E-mail, This is the standard format for internet e-mail addresses., Format: xxxx@xxxxxx.xxx
	                    Response.Write "</font>"
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD VALIGN=TOP>"
	                    Response.Write "<INPUT NAME=""addrFormat"" VALUE=""" & S_DEVICE_VALIDATION_NUMBER & """ TYPE=""RADIO"" "
	                    If aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_NUMBER Then
	                        Response.Write "CHECKED "
	                    End If
	                    Response.Write "/>"
	                    Response.Write "</TD>"
	                    Response.Write "<TD VALIGN=TOP>"
	                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"
	                    Response.Write "<b>" & asDescriptors(505) & "</b><BR />" & asDescriptors(524) & "<BR />" & asDescriptors(525) 'Descriptor: Numeric, Use this format for a string of numbers only., Format: ########
	                    Response.Write "</font>"
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD VALIGN=TOP>"
	                    Response.Write "<INPUT NAME=""addrFormat"" VALUE=""" & S_DEVICE_VALIDATION_NONE & """ TYPE=""RADIO"" "
	                    If aDeviceTypeInfo(DEV_TYPE_ADDR_FORMAT) = S_DEVICE_VALIDATION_NONE Then
	                        Response.Write "CHECKED "
	                    End If
	                    Response.Write "/>"
	                    Response.Write "</TD>"
	                    Response.Write "<TD VALIGN=TOP>"
	                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"
	                    Response.Write "<b>" & asDescriptors(611) & "</b><BR />" & asDescriptors(612) & "<BR />" 'Descriptor: No validation, Address value can be any text string
	                    Response.Write "</font>"
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "</TABLE>"
	                Response.Write "</TD>"
	                Response.Write "</TR>"
	                Response.Write "</TABLE>"
	            Response.Write "</TD>"
	            Response.Write "</TR>"
	            Response.Write "</TABLE>"
	        Response.Write "</TD>"
	        Response.Write "<TD><IMG SRC=""../images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""20"" ALT="""" BORDER=""0"" /></TD>"
	        Response.Write "<TD VALIGN=TOP>"
	            Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
	            Response.Write "<TR>"
	            Response.Write "<TD BGCOLOR=""#cccccc"">"
	                Response.Write "<TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0>"
	                Response.Write "<TR>"
	                Response.Write "<TD BGCOLOR=""#ffffff"">"
	                    Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD COLSPAN=2>"
	                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(506) & "</font>" 'Descriptor: Display fields:
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""DAddrName"" VALUE=""1"" "
	                    If StrComp(aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME), "1", vbBinaryCompare) = 0 Or Len(aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_NAME)) = 0 Then
	                        Response.Write "CHECKED "
	                    'ElseIf aDeviceTypeInfo(DEV_TYPE_ACTION) = "new" Then
	                    '    Response.Write "CHECKED "
	                    End If
	                    Response.Write "/></TD>"
	                    Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(508) & "</font>" 'Descriptor: Address Name
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""DAddrValue"" VALUE=""1"" "
	                    If StrComp(aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE), "1", vbBinaryCompare) = 0 Or Len(aDeviceTypeInfo(DEV_TYPE_DISPLAY_ADDR_VALUE)) = 0 Then
	                        Response.Write "CHECKED "
	                    'ElseIf aDeviceTypeInfo(DEV_TYPE_ACTION) = "new" Then
	                    '    Response.Write "CHECKED "
	                    End If
	                    Response.Write "/></TD>"
	                    Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(509) & "</font>" 'Descriptor: Address Value
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""DStyle"" VALUE=""1"" "
	                    If StrComp(aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE), "1", vbBinaryCompare) = 0 Or Len(aDeviceTypeInfo(DEV_TYPE_DISPLAY_STYLE)) = 0 Then
	                        Response.Write "CHECKED "
	                    'ElseIf aDeviceTypeInfo(DEV_TYPE_ACTION) = "new" Then
	                    '    Response.Write "CHECKED "
	                    End If
	                    Response.Write "/></TD>"
	                    Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(510) & "</font>" 'Descriptor: Style
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""DDeliveryWindow"" VALUE=""1"" "
	                    If aDeviceTypeInfo(DEV_TYPE_DISPLAY_DELIVERY_WINDOW) = "1" Then
	                        Response.Write "CHECKED "
	                    End If
	                    Response.Write "/></TD>"
	                    Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(512) & "</font>" 'Descriptor: Delivery Window
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR><TD COLSPAN=2><BR /></TD></TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD COLSPAN=2>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(507) & "</font>" 'Descriptor: Edit fields:
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD ALIGN=CENTER><IMG SRC=""../images/check.gif"" BORDER=""0"" ALT=""""/></TD>"
	                    Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(508) & "</font>" 'Descriptor: Address Name
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD ALIGN=CENTER><IMG SRC=""../images/check.gif"" BORDER=""0"" ALT=""""/></TD>"
	                    Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(509) & "</font>" 'Descriptor: Address Value
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD ALIGN=CENTER><IMG SRC=""../images/check.gif"" BORDER=""0"" ALT=""""/></TD>"
	                    Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(510) & "</font>" 'Descriptor: Style
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""EPin"" VALUE=""1"" "
	                    If aDeviceTypeInfo(DEV_TYPE_EDIT_PIN) = "1" Then
	                        Response.Write "CHECKED "
	                    End If
	                    Response.Write "/></TD>"
	                    Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(511) & "</font>" 'Descriptor: PIN
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "<TR>"
	                    Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""EDeliveryWindow"" VALUE=""1"" "
	                    If aDeviceTypeInfo(DEV_TYPE_EDIT_DELIVERY_WINDOW) = "1" Then
	                        Response.Write "CHECKED "
	                    End If
	                    Response.Write "/></TD>"
	                    Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(512) & "</font>" 'Descriptor: Delivery Window
	                    Response.Write "</TD>"
	                    Response.Write "</TR>"
	                    Response.Write "</TABLE>"
	                Response.Write "</TD>"
	                Response.Write "</TR>"
	                Response.Write "</TABLE>"
	            Response.Write "</TD>"
	            Response.Write "</TR>"
	            Response.Write "</TABLE>"
	        Response.Write "</TD>"
	        Response.Write "</TR>"
	        Response.Write "</TR>"
	        Response.Write "</TABLE>"
	    Response.Write "</TD>"
	    Response.Write "</TR>"
	    Response.Write "</TABLE>"
	Response.Write "</TD>"
	Response.Write "</TR>"
	Response.Write "</TABLE>"

	RenderDeviceTypeEditor = lErrNumber
	Err.Clear
End Function

Function RenderSelectedDTFolders(sDeviceTypesXML, sDeviceTypeID)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oFolders
	Dim oCurrentFolder

	lErrNumber = NO_ERR

	Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderSelectedDTFolders", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oFolders = oDeviceTypesDOM.selectNodes("/devicetypes/devicetype[devicetypeID = '" & sDeviceTypeID & "']/temp/dfs/f")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderSelectedDTFolders", "", "Error loading f nodes", LogLevelError)
	    End If
	End If

    If lErrNumber = NO_ERR Then
        If oFolders.length > 0 Then
            Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"
            For Each oCurrentFolder in oFolders
                Response.Write "<TR>"
                Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""SDTFolder"" VALUE=""" & oCurrentFolder.getAttribute("id") & ";" & oCurrentFolder.getAttribute("n") & """ /></TD>"
                Response.Write "<TD>"
                Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oCurrentFolder.getAttribute("n") & "</font>"
                Response.Write "</TD>"
                Response.Write "</TR>"
            Next
            Response.Write "</TABLE>"
        Else
            Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000"">" & asDescriptors(519) & "</font>" 'Descriptor: No folders selected.
        End If
    End If

	Set oDeviceTypesDOM = Nothing
	Set oFolders = Nothing
	Set oCurrentFolder = Nothing

	RenderSelectedDTFolders = lErrNumber
	Err.Clear
End Function

Function RenderAvailableDTFolders(sDeviceTypesXML, sDeviceTypeID, sFoldersXML)
'*******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'*******************************************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oFoldersDOM
	Dim oFolders
    Dim oCurrentFolder
    Dim oDeviceTypesDOM
    Dim oDeviceType
	'Dim sFoldersXML

	lErrNumber = NO_ERR
	'sFoldersXML = "<fs><f id=""1"" n=""Folder one"" /><f id=""2"" n=""Folder two"" /><f id=""3"" n=""Folder three"" /><f id=""4"" n=""Folder four"" /></fs>"

	Set oFoldersDOM = Server.CreateObject("Microsoft.XMLDOM")
	oFoldersDOM.async = False
	If oFoldersDOM.loadXML(sFoldersXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderAvailableDTFolders", "", "Error loading sFoldersXML", LogLevelError)
	Else
	    'Set oFolders = oFoldersDOM.selectNodes("//f")
	    Set oFolders = oFoldersDOM.selectNodes("//oi[@tp = '" & TYPE_FOLDER & "']")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderAvailableDTFolders", "", "Error retrieving f nodes", LogLevelError)
	    End If
	End If

	If lErrNumber = NO_ERR Then
	    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	    oDeviceTypesDOM.async = False
	    If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
	        lErrNumber = ERR_XML_LOAD_FAILED
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderAvailableDTFolders", "", "Error loading sDeviceTypesXML", LogLevelError)
	    Else
	        Set oDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & sDeviceTypeID & "']")
	        If Err.number <> NO_ERR Then
	            lErrNumber = Err.number
	            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderAvailableDTFolders", "", "Error loading devicetype node", LogLevelError)
	        End If
	    End If
	End If

    If lErrNumber = NO_ERR Then
        If oFolders.length > 0 Then
            Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"
            For Each oCurrentFolder in oFolders
                'Make sure folder is not already selected
                If (oDeviceType.selectSingleNode("temp/dfs/f[@id = '" & oCurrentFolder.getAttribute("id") & "']") Is Nothing) Then
                    'Make sure folder is not a part of another devicetype dfs, unless it's this one
                    If (oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID != '" & sDeviceTypeID & "']/dfs/f[@id = '" & oCurrentFolder.getAttribute("id") & "']") Is Nothing) Then
                        Response.Write "<TR>"
                        Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""ADTFolder"" VALUE=""" & oCurrentFolder.getAttribute("id") & ";" & oCurrentFolder.getAttribute("n") & """ /></TD>"
                        Response.Write "<TD>"
                        Response.Write "<A HREF=""deviceTypeFolders.asp?folderID=" & oCurrentFolder.getAttribute("id") & "&dtID=" & sDeviceTypeID & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oCurrentFolder.getAttribute("n") & "</font></A>"
                        Response.Write "</TD>"
                        Response.Write "</TR>"
                    End If
                End If
            Next
            Response.Write "</TABLE>"
        Else

        End If
    End If

	Set oFoldersDOM = Nothing
	Set oFolders = Nothing
	Set oCurrentFolder = Nothing
	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing

	RenderAvailableDTFolders = lErrNumber
	Err.Clear
End Function

Function cu_CreateDeviceType(sSiteID, sDeviceTypeID, sDeviceTypeName)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_CreateDeviceType"
	Dim lErrNumber
	Dim sDeviceTypePropertiesXML

	lErrNumber = NO_ERR
	sDeviceTypePropertiesXML = "<mi><in><oi tp=""" & TYPE_DEVICE_TYPE & """ id=""" & sDeviceTypeID & """><prs></prs></oi></in></mi>"

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CreateDeviceType(sSiteID, sDeviceTypePropertiesXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateDeviceType", LogLevelTrace)
        End If
    End If

	cu_CreateDeviceType = lErrNumber
	Err.Clear
End Function

Function cu_CreateDeviceTypeDefinitions(sSiteID, sDeviceTypeID, sDeviceTypeXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_CreateDeviceTypeDefinitions"
	Dim lErrNumber
	Dim asDeviceTypeID(0)
	Dim asDeviceTypeDefinitionXML(0)

	lErrNumber = NO_ERR
	asDeviceTypeID(0) = sDeviceTypeID
	asDeviceTypeDefinitionXML(0) = sDeviceTypeXML

    If lErrNumber = NO_ERR Then
        lErrNumber = co_CreateDeviceTypeDefinitions(sSiteID, asDeviceTypeID, asDeviceTypeDefinitionXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateDeviceTypeDefinitions", LogLevelTrace)
        End If
    End If

	cu_CreateDeviceTypeDefinitions = lErrNumber
	Err.Clear
End Function

Function cu_DeleteDeviceType(sDeviceTypeID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
    Const PROCEDURE_NAME = "cu_DeleteDeviceType"
	Dim lErrNumber
	Dim sSiteID

	lErrNumber = NO_ERR
	sSiteID = Application.Value("SITE_ID")

    If lErrNumber = NO_ERR Then
        lErrNumber = co_DeleteDeviceType(sSiteID, sDeviceTypeID)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling co_DeleteDeviceType", LogLevelTrace)
        End If
    End If

	cu_DeleteDeviceType = lErrNumber
	Err.Clear
End Function

Function cu_GetDeviceTypes(sGetDeviceTypesXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetDeviceTypes"
	Dim lErrNumber
	Dim sSiteID

	lErrNumber = NO_ERR
	sSiteID = Application.Value("SITE_ID")

    If lErrNumber = NO_ERR Then
        lErrNumber = co_GetDeviceTypes(sSiteID, sGetDeviceTypesXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetDeviceTypes", LogLevelTrace)
        End If
    End If

	cu_GetDeviceTypes = lErrNumber
	Err.Clear
End Function

Function cu_GetDeviceTypeDefinitions(sGetDeviceTypesXML, sGetDeviceTypeDefinitionsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetDeviceTypeDefinitions"
	Dim lErrNumber
	Dim sSiteID
	Dim asDeviceTypeID()
	Dim oDeviceTypesDOM
	Dim oDeviceTypes
	Dim i
	Dim bHasDeviceTypes

	lErrNumber = NO_ERR
	sSiteID = Application.Value("SITE_ID")
	bHasDeviceTypes = False

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
    oDeviceTypesDOM.async = False
    If oDeviceTypesDOM.loadXML(sGetDeviceTypesXML) = False Then
        lErrNumber = ERR_XML_LOAD_FAILED
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error loading sGetDeviceTypesXML", LogLevelError)
    Else
        Set oDeviceTypes = oDeviceTypesDOM.selectNodes("/mi/in/oi[@tp = '" & TYPE_DEVICE_TYPE & "']")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error retrieving oi nodes", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
        If oDeviceTypes.length > 0 Then
            bHasDeviceTypes = True
            Redim asDeviceTypeID(oDeviceTypes.length - 1)
            For i=0 To (oDeviceTypes.length - 1)
                asDeviceTypeID(i) = oDeviceTypes.item(i).getAttribute("id")
            Next
        End If
    End If

    Set oDeviceTypesDOM = Nothing
    Set oDeviceTypes = Nothing

    If (lErrNumber = NO_ERR) And (bHasDeviceTypes = True) Then
        lErrNumber = co_GetDeviceTypeDefinitions(sSiteID, asDeviceTypeID, sGetDeviceTypeDefinitionsXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetDeviceTypeDefinitions", LogLevelTrace)
        End If
    End If

	cu_GetDeviceTypeDefinitions = lErrNumber
	Err.Clear
End Function

Function cu_UpdateDeviceTypeDefinitions(sDeviceTypeID, sDeviceTypeXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_UpdateDeviceTypeDefinitions"
	Dim lErrNumber
	Dim sSiteID
	Dim asDeviceTypeID(0)
	Dim asDeviceTypeDefinitionXML(0)

	lErrNumber = NO_ERR
	sSiteID = Application.Value("SITE_ID")
	asDeviceTypeID(0) = sDeviceTypeID
	asDeviceTypeDefinitionXML(0) = sDeviceTypeXML

    If lErrNumber = NO_ERR Then
        lErrNumber = co_UpdateDeviceTypeDefinitions(sSiteID, asDeviceTypeID, asDeviceTypeDefinitionXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling co_UpdateDeviceTypeDefinitions", LogLevelTrace)
        End If
    End If

	cu_UpdateDeviceTypeDefinitions = lErrNumber
	Err.Clear
End Function

Function cu_GetFolderContents(sFolderID, sGetFolderContentsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetFolderContents"
	Const N_DEVICES_FOLDER = 14
	Dim lErrNumber
	Dim sSiteID
	Dim iDefaultFolder

	lErrNumber = NO_ERR
	sSiteID = Application.Value("SITE_ID")
	If Len(sFolderID) > 0 Then
	    iDefaultFolder = 0
	Else
	    iDefaultFolder = N_DEVICES_FOLDER
	End If

    If lErrNumber = NO_ERR Then
        lErrNumber = co_GetFolderContents(sSiteID, sFolderID, iDefaultFolder, sGetFolderContentsXML)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetFolderContents", LogLevelTrace)
        End If
    End If

	cu_GetFolderContents = lErrNumber
	Err.Clear
End Function

Function RenderPath_DeviceFolders(sDeviceTypeID, sFoldersXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oFoldersDOM
	Dim oFolder
    Dim iNumFolders
    Dim i

    lErrNumber = NO_ERR

	Set oFoldersDOM = Server.CreateObject("Microsoft.XMLDOM")
	oFoldersDOM.async = False
	If oFoldersDOM.loadXML(sFoldersXML) = False Then
	    lErrNumber = ERR_XML_LOAD_FAILED
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderPath_DeviceFolders", "", "Error loading sFoldersXML", LogLevelError)
	Else
	    iNumFolders = CInt(oFoldersDOM.selectNodes("//a").length)
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", "RenderPath_DeviceFolders", "", "Error retrieving a nodes", LogLevelError)
	    End If
	End If

    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(26) 'Descriptor: You are here:
    If lErrNumber = NO_ERR Then
        If iNumFolders > 0 Then
            Set oFolder = oFoldersDOM.selectSingleNode("/mi/as")
            For i=1 To iNumFolders
                Set oFolder = oFolder.selectSingleNode("a")
                Response.Write " > "
                If i = iNumFolders Then
                    Response.Write "<b>" & oFolder.selectSingleNode("fd").getAttribute("n") & "</b>"
                Else
                    Response.Write "<A HREF=""deviceTypeFolders.asp?dtID=" & sDeviceTypeID & "&folderID=" & oFolder.selectSingleNode("fd").getAttribute("id") & """><font color=""#0000"">" & oFolder.selectSingleNode("fd").getAttribute("n") & "</font></A>"
                End If
            Next
        End If
    End If
    Response.Write "</font>"

    Set oFoldersDOM = Nothing
    Set oFolder = Nothing

    RenderPath_DeviceFolders = lErrNumber
    Err.Clear
End Function

Function GetFilePath()
'********************************************************
'*Purpose: Returns the file path of the asp folder
'*Inputs:  None
'*Outputs: The physical path of the asp folder.
'********************************************************
Dim sFilePath

	On Error Resume Next

	sFilePath = Server.MapPath("./")
	If Right(sFilePath, 5) = "admin" Then
	    sFilePath = Server.MapPath("../") & "\"
	Else
	    sFilePath = sFilePath & "\"
	End If

	GetFilePath = sFilePath
	Err.Clear

End Function


%>