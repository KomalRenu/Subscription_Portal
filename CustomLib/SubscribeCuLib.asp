<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/SubscribeCoLib.asp" -->
<!--#include file="../CoreLib/AddressCoLib.asp" -->
<%
Function ParseRequestForSubscription(oRequest, sServiceID, sServiceName, sAddressID, sAddressName, sPublicationID, sSubSetID, sScheduleName, sFolderID, sESGUID, sEAID, sESSID, sEPUBID, sStatusFlag, sEnabledFlag, sNewAddressValue, sTransPropsID, sSubID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
    Dim lErrNumber
    Dim sAddressType
    Dim sAddressData
    Dim sScheduleData
    Dim iColonPlace

    lErrNumber = NO_ERR

	sServiceID = ""
	sServiceName = ""
	sAddressID = ""
	sAddressName = ""
	sPublicationID = ""
	sSubSetID = ""
	sScheduleName = ""
	sFolderID = ""
	sESGUID = ""
	sEAID = ""
	sESSID = ""
	sStatusFlag = ""
	sEnabledFlag = ""
	sTransPropsID = ""

	sServiceID = Trim(CStr(oRequest("serviceID")))
	sServiceName = Trim(CStr(oRequest("serviceName")))
	sFolderID = Trim(CStr(oRequest("folderID")))
	sESGUID = Trim(CStr(oRequest("eSGUID")))
	sSubID = Trim(CStr(oRequest("esubID")))
	sEAID = Trim(CStr(oRequest("eAID")))
	sESSID = Trim(CStr(oRequest("eSSID")))
	sEPUBID = Trim(CStr(oRequest("ePUBID")))
	sAddressType = Trim(CStr(oRequest("ServAdd")))

	Select Case sAddressType
	    Case "p" 'Portal Address
	        sPublicationID = Trim(CStr(oRequest("PortalPubID")))
	        sTransPropsID = Trim(Cstr(oRequest("PortalTRPS")))
            sAddressID = GetPortalAddress()
	        sAddressName = Trim(CStr(oRequest("PortalAddressName")))
	    Case "a" 'Regular Address
	        sAddressData = Trim(CStr(oRequest("addressID")))
	        If Len(sAddressData) > 0 Then
	            iColonPlace = Instr(1, sAddressData, ":")
	            sPublicationID = Left(sAddressData, iColonPlace - 1)
	            sAddressID = Mid(sAddressData, iColonPlace+1, Instr(iColonPlace + 1, sAddressData, ":") - iColonPlace -1)
	            iColonPlace = Instr(iColonPlace + 1, sAddressData, ":")
	            sTransPropsID = Mid(sAddressData, iColonPlace+1, Instr(iColonPlace + 1, sAddressData, ":") - iColonPlace -1)
	            sAddressName = Right(sAddressData, Len(sAddressData) - InStrRev(sAddressData, ":"))
	        End If
	    Case "n" 'New Address
	        sPublicationID = Trim(CStr(oRequest("NewAddrPub")))
	        sNewAddressValue = Trim(CStr(oRequest("NewAddrVal")))
	    Case Else
	End Select

	sScheduleData = Trim(CStr(oRequest("subsSetID")))
	If Len(sScheduleData) > 0 Then
	    iColonPlace = Instr(1, sScheduleData, ":")
	    sSubsSetID = Left(sScheduleData, iColonPlace - 1)
	    sScheduleName = Right(sScheduleData, Len(sScheduleData) - iColonPlace)
	End If

	If StrComp(CStr(oRequest("enfCheck")), "1", vbBinaryCompare) = 0 Then
	    If StrComp(CStr(oRequest("subsEnabled")), "on", vbBinaryCompare) = 0 Then
	        sEnabledFlag = "1"
	    Else
	        sEnabledFlag = "2"
	    End If
	End If
	If Len(CStr(oRequest("enf"))) > 0 Then sEnabledFlag = CStr(oRequest("enf"))

	If Err.number <> NO_ERR Then
	    lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "ParseRequestForSubscription", "", "Error setting variables equal to Request variables", LogLevelError)
	Else
	    If Len(sServiceID) = 0 Then
	    	lErrNumber = URL_MISSING_PARAMETER
	    End If
	End If

        If Len(sESGUID) > 0 Then
            If StrComp(CStr(oRequest("sf")), "1", vbBinaryCompare) = 0 Then
                sStatusFlag = "1"
            Else
                sStatusFlag = "0"
            End If
        Else
            sStatusFlag = "1"
	    End If

	ParseRequestForSubscription = lErrNumber
	Err.Clear
End Function

Function cu_GetUserAddressesForService(sServiceID, sGetUserAddressesForServiceXML)
'********************************************************
'*Purpose:
'*Inputs: sServiceID
'*Outputs: sGetUserAddressesForServiceXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetUserAddressesForService"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

	lErrNumber = co_GetUserAddressesForService(sSessionID, sServiceID, sGetUserAddressesForServiceXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetUserAddressesForService", LogLevelTrace)
	End If

	cu_GetUserAddressesForService = lErrNumber
	Err.Clear
End Function

Function cu_GetNamedSchedulesForService(sServiceID, bFlagValid, sGetNamedSchedulesForServiceXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetNamedSchedulesForService"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

	lErrNumber = co_GetNamedSchedulesForService(sSessionID, sServiceID, bFlagValid, sGetNamedSchedulesForServiceXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetNamedSchedulesForService", LogLevelTrace)
	End If

	cu_GetNamedSchedulesForService = lErrNumber
	Err.Clear
End Function

Function RenderAddressesForService(sEditAddressID, sAddressXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oAddressDOM
	Dim oAddresses
	Dim oAddress
	Dim oPortalAddress
	Dim oDefaultPublication
	Dim bHasPortalAddress
	Dim bHasNonPortalAddresses
	Dim bSupportsNew

	lErrNumber = NO_ERR

	Set oAddressDOM = Server.CreateObject("Microsoft.XMLDOM")
	oAddressDOM.async = False
	If oAddressDOM.loadXML(sAddressXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "RenderAddressesForService", "", "Error loading sAddressXML", LogLevelError)
		Response.Write "<BR /><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(376) & "</b></font><BR /><BR />" 'Descriptor: Error retrieving addresses
	End If

    If lErrNumber = NO_ERR Then
	    Set oAddresses = oAddressDOM.selectNodes("//oi[@tp = '" & TYPE_ADDRESS & "']")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "RenderAddressesForService", "", "Error retrieving ADDRESS nodes", LogLevelError)
	    	Response.Write "<BR /><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(376) & "</b></font><BR /><BR />" 'Descriptor: Error retrieving addresses
	    Else
	        Set oPortalAddress = oAddressDOM.selectSingleNode("//oi[@tp = '" & TYPE_ADDRESS & "' and @id = '" & GetPortalAddress() & "']")
	        If Not (oPortalAddress Is Nothing) Then
	            bHasPortalAddress = True
	            If oAddresses.length > 1 Then
	                bHasNonPortalAddresses = True
	            Else
	                bHasNonPortalAddresses = False
	            End If
	        Else
	            bHasPortalAddress = False
	            If oAddresses.length > 0 Then
	                bHasNonPortalAddresses = True
	            Else
	                bHasNonPortalAddresses = False
	            End If
	        End If
	    End If
	End If

    If lErrNumber = NO_ERR Then
        Set oDefaultPublication = oAddressDOM.selectSingleNode("mi/in/oi/mi/in/oi[@tp = '" & TYPE_PUBLICATION & "' and mi/in/oi[@id = '" & Application("Default_Device") & "']]")
        If Not (oDefaultPublication Is Nothing) Then
            bSupportsNew = True
        Else
            bSupportsNew = False
        End If
    End If

    If lErrNumber = NO_ERR Then
        If oAddresses.length > 0 Then
            Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""10%"" >"

            If (bHasPortalAddress = True) Then
                Response.Write "<TR BGCOLOR=""#ffffff""><TD WIDTH=""1%""><INPUT TYPE=""RADIO"" NAME=""ServAdd"" VALUE=""p"" "
                If (sEditAddressID = GetPortalAddress()) Or (Len(sEditAddressID) = 0) Then
                    Response.Write "CHECKED"
                End If
                Response.Write " /></TD><TD nowrap WIDTH=""99%""><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(450) & "</font></TD></TR>" 'Descriptor: Send to my Reports page
   	            Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PortalPubID"" VALUE=""" & oAddressDOM.selectSingleNode("/mi/in/oi/mi/in/oi[@tp = '" & TYPE_PUBLICATION & "' and mi/in/oi/mi/in/oi[@id = '" & GetPortalAddress() & "']]").getAttribute("id") & """ />"
	            Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PortalTRPS"" VALUE=""" & oAddressDOM.selectSingleNode("//oi[@tp = '" & TYPE_ADDRESS & "' and @id = '" & GetPortalAddress() & "']").getAttribute("trps") & """ />"
	            Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PortalAddressName"" VALUE=""" & Server.HTMLEncode(oAddressDOM.selectSingleNode("//oi[@tp = '" & TYPE_ADDRESS & "' and @id = '" & GetPortalAddress() & "']").getAttribute("n")) & """ />"
            End If

            If (bHasPortalAddress = True) And (bHasNonPortalAddresses = True) Then
                Response.Write "<TR BGCOLOR=""#ffffff""><TD></TD><TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & asDescriptors(339) & "</b></font></TD></TR>" 'Descriptor: or
            ElseIf (bHasPortalAddress = True) And (bSupportsNew = True) Then
                Response.Write "<TR BGCOLOR=""#ffffff""><TD></TD><TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & asDescriptors(339) & "</b></font></TD></TR>" 'Descriptor: or
            End If

            If (bHasNonPortalAddresses = True) Then
                Set oAddresses = oAddressDOM.selectNodes("//oi[@tp = '" & TYPE_ADDRESS & "' and @id != '" & GetPortalAddress() & "']")
	            Response.Write "<TR BGCOLOR=""#ffffff""><TD WIDTH=""1%""><INPUT TYPE=""RADIO"" NAME=""ServAdd"" VALUE=""a"" "
	            If (sEditAddressID <> "") And (sEditAddressID <> GetPortalAddress()) Then
                    Response.Write "CHECKED"
                ElseIf (bHasPortalAddress = False) Then
                    Response.Write "CHECKED"
                End If
	            Response.Write " /></TD><TD WIDTH=""99%"" NOWRAP><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(373) & "&nbsp;</font>" 'Descriptor: Choose an address

	            If GetJavaScriptSetting() = "1" Then
	              Response.Write "<select name=""addressID"" class=""pullDownClass"" onFocus=""return autoSelect('a');"" >"
	            Else
	              Response.Write "<select name=""addressID"" class=""pullDownClass"" >"
	            End If

	    	    For Each oAddress In oAddresses
	    	    	Response.Write "<option value=""" & oAddress.parentNode.parentNode.parentNode.parentNode.parentNode.parentNode.getAttribute("id") & ":" & oAddress.getAttribute("id") & ":" & oAddress.getAttribute("trps") & ":" & Server.HTMLEncode(oAddress.getAttribute("n")) & """"
	    	    	If oAddress.getAttribute("id") = sEditAddressID Then
	    	    		Response.Write " SELECTED"
	    	    	End If
	    	    	Response.Write ">" & Server.HTMLEncode(oAddress.getAttribute("n")) & "</option>"
	    	    Next
	    	    Response.Write "</select>"
	            Response.Write "</TD></TR>"
            End If

            If (bHasNonPortalAddresses = True) And (bSupportsNew = True) Then
                Response.Write "<TR BGCOLOR=""#ffffff""><TD></TD><TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & asDescriptors(339) & "</b></font></TD></TR>" 'Descriptor: or
            End If

            If (bSupportsNew = True) Then
	            Response.Write "<TR BGCOLOR=""#ffffff""><TD WIDTH=""1%""><INPUT TYPE=""RADIO"" NAME=""ServAdd"" VALUE=""n"" /></TD><TD WIDTH=""99%"" NOWRAP><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(492) & "</font> " 'Descriptor: Enter a new address:
	            If GetJavaScriptSetting() = "1" Then
	                Response.Write "<INPUT TYPE=""TEXT"" NAME=""NewAddrVal"" CLASS=""textboxClass"" VALUE="""" SIZE=""35"" onFocus=""return autoSelect('n');"" />"
	            Else
	                Response.Write "<INPUT TYPE=""TEXT"" NAME=""NewAddrVal"" CLASS=""textboxClass"" VALUE="""" SIZE=""35"" />"
	            End If
	            Response.Write "</TD></TR>"
	            Response.Write "<TR><TD COLSPAN=""2"" ALIGN=""RIGHT""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"
	            Response.Write "(" & asDescriptors(510) & ": " & Server.HTMLEncode(Application("Default_Device_Name")) & ")" 'Descriptor: Style
	            Response.Write "</FONT>"
	            Response.Write "</TD></TR>"
	            Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""NewAddrPub"" VALUE=""" & oDefaultPublication.getAttribute("id") & """ />"
            End If

            Response.Write "</TABLE>"
        ElseIf (bSupportsNew = True) Then
            Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""10%"">"
	        Response.Write "<TR BGCOLOR=""#ffffff""><TD WIDTH=""1%""><INPUT TYPE=""RADIO"" NAME=""ServAdd"" VALUE=""n"" CHECKED /></TD><TD WIDTH=""99%"" NOWRAP><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(492) & "</font> " 'Descriptor: Enter a new address:
	        Response.Write "<INPUT TYPE=""TEXT"" NAME=""NewAddrVal"" CLASS=""textboxClass"" VALUE="""" SIZE=""35"" /></TD></TR>"
	        Response.Write "<TR><TD COLSPAN=""2"" ALIGN=""RIGHT""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"
	        Response.Write "(" & asDescriptors(510) & ": " & Application("Default_Device_Name") & ")" 'Descriptor: Style
	        Response.Write "</FONT>"
	        Response.Write "</TD></TR>"
	        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""NewAddrPub"" VALUE=""" & oDefaultPublication.getAttribute("id") & """ />"
            Response.Write "</TABLE>"
        Else
            Response.Write "<BR /><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(377) & "</b></font><BR /><BR />" 'Descriptor: You do not have any compatible addresses
        End If
    End If

	Set oAddressDOM = Nothing
	Set oAddresses = Nothing
	Set oAddress = Nothing
	Set oPortalAddress = Nothing
	Set oDefaultPublication = Nothing

	RenderAddressesForService = lErrNumber
	Err.Clear
End Function

Function RenderSchedulesForService(sEditSubsSetID, sScheduleXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oScheduleDOM
	Dim oSubsSets
	Dim oCurrentSubsSet
	Dim oSchedules
	Dim oCurrentSchedule
	Dim sSchedules

	lErrNumber = NO_ERR

	Set oScheduleDOM = Server.CreateObject("Microsoft.XMLDOM")
	oScheduleDOM.async = False
	If oScheduleDOM.loadXML(sScheduleXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "RenderSchedulesForService", "", "Error loading sScheduleXML", LogLevelError)
		Response.Write "<BR /><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(378) & "</b></font><BR /><BR />" 'Descriptor: Error retrieving schedules
	End If

    If lErrNumber = NO_ERR Then
	    Set oSubsSets = oScheduleDOM.selectNodes("//oi[@tp = '" & TYPE_SUBSET & "']")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "RenderSchedulesForService", "", "Error retrieving subscriptionSet nodes", LogLevelError)
	    	Response.Write "<BR /><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(378) & "</b></font><BR /><BR />" 'Descriptor: Error retrieving schedules
	    End If
    End If

	If lErrNumber = NO_ERR Then
	    If oSubsSets.length > 0 Then
	        Response.Write "<TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0>"
	    	Response.Write "<TR><TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(534) & "</font></TD></TR>" 'Descriptor: Select a schedule from the list below.
	    	Response.Write "<TR><TD><select name=""subsSetID"" class=""pullDownClass"">"
	    	For Each oCurrentSubsSet In oSubsSets 'for each subscriptionSet
	    		sSchedules = ""
	    		Set oSchedules = oCurrentSubsSet.selectNodes("mi/oi[@tp = '" & TYPE_SCHEDULE & "']")
	    		'QUESTION: Error handling here?
	    		If oSchedules.length > 0 Then
	    			For Each oCurrentSchedule In oSchedules
	    				sSchedules = sSchedules & Server.HTMLEncode(oCurrentSchedule.getAttribute("n")) & ", "
	    			Next
	    			sSchedules = Left(sSchedules, Len(sSchedules)-2)
	    		Else
	    			'QUESTION: What should be done here?
	    		End If

	    		Set oSchedules = Nothing
	    		Set oCurrentSchedule = Nothing

	    		Response.Write "<option value=""" & oCurrentSubsSet.getAttribute("id") & ":" & sSchedules & """"
	    		If oCurrentSubsSet.getAttribute("id") = sEditSubsSetID Then
	    			Response.Write " SELECTED"
	    		End If
	    		Response.Write ">" & sSchedules & "</option>"
	    	Next
	    	Response.Write "</select></TD></TR>"
	    	Response.Write "</TABLE>"
	    Else
	    	Response.Write "<BR /><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(379) & "</b></font><BR /><BR />" 'Descriptor: There are no schedules available
	    End If
	End If

	Set oScheduleDOM = Nothing
	Set oSubsSets = Nothing
	Set oCurrentSubsSet = Nothing

	RenderSchedulesForService = lErrNumber
	Err.Clear
End Function

	Function RenderServiceInformation(sServiceInfoXML)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'*TO DO: Add error handling!
	'********************************************************
		On Error Resume Next
		Dim oSInfoDOM
		Dim oSInfo
		Dim lErrNumber
		lErrNumber = NO_ERR

		Set oSInfoDOM = Server.CreateObject("Microsoft.XMLDOM")
		oSInfoDOM.async = False
		If oSInfoDOM.loadXML(sServiceInfoXML) = False Then
			lErrNumber = ERR_XML_LOAD_FAILED
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "RenderServiceInformation", "", "Error loading sServiceInfoXML", LogLevelError)
			'Add error message
		End If

		If lErrNumber = NO_ERR Then
		    Set oSInfo = oSInfoDOM.selectSingleNode("/RESPONSE/OBJECTS/OBJECT")
		    If Err.number <> NO_ERR Then
		        lErrNumber = Err.number
		        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "RenderServiceInformation", "", "Error retrieving ROW node", LogLevelError)
		        'Add error message
		    Else
		        Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
		        Response.Write "<TR>"
		        Response.Write "<TD NOWRAP>"
		        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """ color=""#cc0000""><b>" & oSInfo.selectSingleNode("OBJECT_NAME").text & "</b></font>"
		        Response.Write "</TD>"
		        Response.Write "</TR>"
		        Response.Write "<TR>"
		        Response.Write "<TD>"
		        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oSInfo.selectSingleNode("OBJECT_DESCRIPTION").text & "</font>"
		        Response.Write "</TD>"
		        Response.Write "</TR>"
		        Response.Write "</TABLE>"
		    End If
		End If

		Set oSInfoDOM = Nothing
		Set oSInfo = Nothing

		RenderServiceInformation = lErrNumber
		Err.Clear
	End Function

Function CheckNumberOfSchedules(sGetNamedSchedulesForServiceXML, bHasSchedules)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oScheduleDOM
    Dim oSchedules

    lErrNumber = NO_ERR
    bHasSchedules = False

	Set oScheduleDOM = Server.CreateObject("Microsoft.XMLDOM")
	oScheduleDOM.async = False
	If oScheduleDOM.loadXML(sGetNamedSchedulesForServiceXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "CheckNumberOfSchedules", "", "Error loading sGetNamedSchedulesForServiceXML", LogLevelError)
    Else
        Set oSchedules = oScheduleDOM.selectNodes("//oi[@tp = '" & TYPE_SUBSET & "']")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "CheckNumberOfSchedules", "", "Error retrieving subscription set oi nodes", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
        If oSchedules.length > 0 Then
            bHasSchedules = True
        End If
    End If

    Set oScheduleDOM = Nothing
    Set oSchedules = Nothing

    CheckNumberOfSchedules = lErrNumber
    Err.Clear
End Function

Function CheckNumberOfAddresses(sGetUserAddressesForServiceXML, bHasAddresses)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oAddressDOM
    Dim oAddresses
    Dim oDefaultPublication

    lErrNumber = NO_ERR
    bHasAddresses = False

	Set oAddressDOM = Server.CreateObject("Microsoft.XMLDOM")
	oAddressDOM.async = False
	If oAddressDOM.loadXML(sGetUserAddressesForServiceXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "CheckNumberOfAddresses", "", "Error loading sGetUserAddressesForServiceXML", LogLevelError)
    Else
        Set oAddresses = oAddressDOM.selectNodes("//oi[@tp = '" & TYPE_ADDRESS & "']")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "CheckNumberOfAddresses", "", "Error retrieving address oi nodes", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
        If oAddresses.length > 0 Then
            bHasAddresses = True
        Else
            Set oDefaultPublication = oAddressDOM.selectSingleNode("mi/in/oi/mi/in/oi[@tp = '" & TYPE_PUBLICATION & "' and mi/in/oi[@id = '" & Application("Default_Device") & "']]")
            If Not (oDefaultPublication Is Nothing) Then
                bHasAddresses = True
            End If
        End If
    End If

    Set oAddressDOM = Nothing
    Set oAddresses = Nothing
    Set oDefaultPublication = Nothing

    CheckNumberOfAddresses = lErrNumber
    Err.Clear
End Function

Function GetOriginalPublication(sEAID, sGetUserAddressesForServiceXML, sEPUBID)
'********************************************************
'*Purpose:
'*Inputs: sEAID, sGetUserAddressesForServiceXML
'*Outputs: sEPUBID
'*TO DO: Add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oAddressDOM

    lErrNumber = NO_ERR
    sEPUBID = ""

	Set oAddressDOM = Server.CreateObject("Microsoft.XMLDOM")
	oAddressDOM.async = False
	If oAddressDOM.loadXML(sGetUserAddressesForServiceXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "GetOriginalPublication", "", "Error loading sGetUserAddressesForServiceXML", LogLevelError)
    Else
        sEPUBID = CStr(oAddressDOM.selectSingleNode("/mi/in/oi/mi/in/oi[mi/in/oi/mi/in/oi/@id = '" & sEAID & "']").getAttribute("id"))
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "GetOriginalPublication", "", "Error retrieving publication ID", LogLevelError)
        End If
	End If

    Set oAddressDOM = Nothing

    GetOriginalPublication = lErrNumber
    Err.Clear
End Function

Function RenderPath_Subscribe(sServiceID, sServiceName, sFolderID, sGetFolderContentsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: add error handling, messages
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oContentsDOM
    Dim oFolder
    Dim iNumFolders
    Dim i
    Dim sLastFolder

    iNumFolders = 0
    lErrNumber = NO_ERR
    sLastFolder = ""

    If sFolderID <> "" Then
        Set oContentsDOM = Server.CreateObject("Microsoft.XMLDOM")
	    oContentsDOM.async = False
	    If oContentsDOM.loadXML(sGetFolderContentsXML) = False Then
	    	lErrNumber = ERR_XML_LOAD_FAILED
	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "RenderPath_Subscribe", "", "Error loading sGetFolderContentsXML", LogLevelError)
	    	'add error message
        Else
            iNumFolders = CInt(oContentsDOM.selectNodes("//a").length)
            If Err.number <> NO_ERR Then
                lErrNumber = Err.number
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscribeCuLib.asp", "RenderPath_Subscribe", "", "Error retrieving oi nodes", LogLevelError)
                'add error message
            End If
	    End If
	End If

    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"
    Response.Write asDescriptors(26) & " " 'Descriptor: You are here:
    If lErrNumber = NO_ERR Then
        If iNumFolders > 0 Then
            Response.Write "<A HREF=""services.asp""><font color=""#000000"">" & asDescriptors(362) & "</font></A>" 'Descriptor: Services
            Set oFolder = oContentsDOM.selectSingleNode("/mi/as")
            For i=1 To iNumFolders
                Set oFolder = oFolder.selectSingleNode("a")
                Response.Write " > "
                Response.Write "<A HREF=""services.asp?folderID=" & oFolder.selectSingleNode("fd").getAttribute("id") & """><font color=""#0000"">" & oFolder.selectSingleNode("fd").getAttribute("n") & "</font></A>"
                If i=iNumFolders Then sLastFolder = oFolder.selectSingleNode("fd").getAttribute("id")
            Next
            Response.Write " > <b>" & asDescriptors(457) & " " & oContentsDOM.selectSingleNode("/mi/fct/oi[@id = '" & sServiceID & "']").getAttribute("n") & "</b>" 'Descriptor: Subscribe to:
        Else
            Response.Write "<b>" & asDescriptors(457) & " " & sServiceName & "</b>" 'Descriptor: Subscribe to:
        End If
    Else
        'add handling
    End If
    Response.Write "</font>"

    Set oContentsDOM = Nothing
    Set oFolder = Nothing

    RenderPath_Subscribe = lErrNumber
    Err.Clear
End Function

Function GetVariablesFromCache_Subscribe(sCacheXML, sSelectedAddressID, sSelectedSubSetID, sEnabledFlag)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oSub

    lErrNumber = NO_ERR

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
	oCacheDOM.async = False
    oCacheDOM.loadXML(sCacheXML)

    Set oSub = oCacheDOM.selectSingleNode("/mi/sub")

    sSelectedAddressID = oSub.getAttribute("adid")
    sSelectedSubSetID = oSub.getAttribute("sbstid")
    If Len(sEnabledFlag) = 0 Then
        sEnabledFlag = oSub.getAttribute("enf")
    End If

    Set oSub = Nothing
    Set oCacheDOM = Nothing

    GetVariablesFromCache_Subscribe = lErrNumber
    Err.Clear
End Function

Function GenerateCacheXML(sSubGUID, sServiceID, sServiceName, sFolderID, sStatusFlag, sSubsSetID, sAddressID, sTransPropsID, sPublicationID, sAddressName, sScheduleName, sEditSubsSetID, sEditAddressID, sEditPublicationID, sSubsEnabled, sSubID, sCacheXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber

    lErrNumber = NO_ERR

    sCacheXML = ""
    sCacheXML = "<sub subGUID=""" & sSubGUID & """ subid=""" & sSubID & """ svcid=""" & sServiceID & """ svn=""" & Server.HTMLEncode(sServiceName) & """ "
    sCacheXML = sCacheXML & "fid=""" & sFolderID & """ sf=""" & sStatusFlag & """ sbstid=""" & sSubsSetID & """ "
    sCacheXML = sCacheXML & "adid=""" & sAddressID & """ trps=""" & sTransPropsID & """ pubid=""" & sPublicationID & """ adn=""" & Server.HTMLEncode(sAddressName) & """ "
    sCacheXML = sCacheXML & "scn=""" & Server.HTMLEncode(sScheduleName) & """ esbstid=""" & sEditSubsSetID & """ eaid=""" & sEditAddressID & """ "
    sCacheXML = sCacheXML & "epubid=""" & sEditPublicationID & """ enf=""" & sSubsEnabled & """>"
    sCacheXML = sCacheXML & "</sub>"
    sCacheXML = "<mi>" & sCacheXML & "</mi>"

    GenerateCacheXML = lErrNumber
    Err.Clear
End Function

Function AddNewSubscriptionAddress(sNewAddressValue, sNewSubsAddrID, sNewAddressTRPS)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
    Const PROCEDURE_NAME = "AddNewSubscriptionAddress"
	Dim lErrNumber
	Dim asAddressProperties()
	Redim asAddressProperties(MAX_ADDR_PROP)
	Dim sSessionID
	Dim sNewSubsAddressXML
	Dim bGenerateTransProps

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
	bGenerateTransProps = False

    sNewSubsAddrID = GetGUID()
    sNewAddressTRPS = GetGUID()

	asAddressProperties(ADDR_PROP_ADDRESS_ID) = sNewSubsAddrID      'addressID
	asAddressProperties(ADDR_PROP_ADDRESS_NAME) = sNewAddressValue       'addressName
	asAddressProperties(ADDR_PROP_PHYSICAL_ADDRESS) = sNewAddressValue      'physicalAddress
	asAddressProperties(ADDR_PROP_ADDRESS_DISPLAY) = sNewAddressValue		'addressDisplay
	asAddressProperties(ADDR_PROP_DEVICE_ID) = Application("Default_Device")     'deviceID
	asAddressProperties(ADDR_PROP_DELIVERY_WINDOW) = ""						'deliveryWindow
	asAddressProperties(ADDR_PROP_TIMEZONE_ID) = GetDefaultTimeZone()								'DefaultTimezoneStdName
	asAddressProperties(ADDR_PROP_STATUS) = "1"								'status
	asAddressProperties(ADDR_PROP_CREATED_BY) = ""								'createdBy
	asAddressProperties(ADDR_PROP_LAST_MODIFIED_BY) = ""								'lastModBy
	asAddressProperties(ADDR_PROP_TRANSMISSION_PROPERTIES_ID) = sNewAddressTRPS               'transPropsID
	asAddressProperties(ADDR_PROP_PIN) = ""								'PIN
	asAddressProperties(ADDR_PROP_EXPIRATION_DATE) = ""		'expirationDate
	asAddressProperties(ADDR_PROP_CREATED_DATE) = ""		'createdDate
	asAddressProperties(ADDR_PROP_LAST_MODIFIED_DATE) = ""		'lastModDate

	lErrNumber = co_AddAddress(sSessionID, asAddressProperties, bGenerateTransProps, sNewSubsAddressXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "Personalize.asp", PROCEDURE_NAME, "", "Error while calling co_AddAddress", LogLevelTrace)
	End If

	AddNewSubscriptionAddress = lErrNumber
	Err.Clear
End Function

Function UpdateCache_Subscribe(sSubsSetID, sAddressID, sTransPropsID, sPublicationID, sAddressName, sScheduleName, sEnabledFlag, sStatusFlag, sSubID, sCacheXML)
'********************************************************
'*Purpose:	Update Cache when user go back to change the selections
'*Inputs:	sSubsSetID, sAddressID, sTransPropsID, sPublicationID, sAddressName, sScheduleName, sEnabledFlag, sCacheXML
'*Outputs:	sCacheXML
'********************************************************
	On Error Resume Next
    Const PROCEDURE_NAME = "UpdateCache_Subscribe"
	Dim lErrNumber
	Dim oCacheDOM
	Dim oSub

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
    If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.Source, "PromptCuLib.asp", PROCEDURE_NAME, "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
    Else
		Set oSub = oCacheDOM.selectSingleNode("/mi/sub")

		Call oSub.setAttribute("sbstid", sSubsSetID)
		Call oSub.setAttribute("adid", sAddressID)
		Call oSub.setAttribute("trps", sTransPropsID)
		Call oSub.setAttribute("pubid", sPublicationID)
		Call oSub.setAttribute("adn", sAddressName)
		Call oSub.setAttribute("scn", sScheduleName)
		Call oSub.setAttribute("enf", sEnabledFlag)
		Call oSub.setAttribute("sf", sStatusFlag)
		Call oSub.setAttribute("subid", sSubID)
		sCacheXML = oCacheDOM.xml
	End If

	UpdateCache_Subscribe = lErrNumber
	Err.Clear
End Function

%>