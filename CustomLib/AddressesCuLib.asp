<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/AddressCoLib.asp" -->
<%
Function ParseRequestForAddresses(oRequest, sDeviceTypeID, sEditAddID, sAction, sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld, sDevice, sPIN, sDeliveryWindow, sDelAddrID, sTransPropsID, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
'********************************************************
'*Purpose: Sets variables used on addresses.asp equal to items in Request object.
'*Inputs: oRequest
'*Outputs: sDeviceTypeID, sEditAddID, sAction, sReqUserAddresses
'********************************************************
	On Error Resume Next
    Dim lErrNumber

    lErrNumber = NO_ERR

	sDeviceTypeID = ""
	sEditAddID = ""
	sAction = ""
	sAddressName = ""
	sPhysicalAddress = ""
	sAddressNameVld = ""
	sPhysicalAddressVld = ""
	sDevice = ""
	sPIN = ""
	sDeliveryWindow = ""
	sDelAddrID = ""
	sTransPropsID = ""
	sWizardDeviceID = ""
	sWizardAddressName = ""
	sWizardPhysicalAddress = ""

	sDeviceTypeID = Trim(CStr(oRequest("deviceTypeID")))
	sEditAddID = Trim(CStr(oRequest("editAddID")))
	sAction = Trim(CStr(oRequest("action")))
	sAddressName = Trim(CStr(oRequest("AddressName")))
	sPhysicalAddress = Trim(CStr(oRequest("PhysicalAddress")))
	sAddressNameVld = Trim(CStr(oRequest("AddressNameVld")))
	sPhysicalAddressVld = Trim(CStr(oRequest("PhysicalAddressVld")))
	sDevice = Trim(CStr(oRequest("Device")))
	sPIN = Trim(CStr(oRequest("PIN")))
	sDeliveryWindow = Trim(CStr(oRequest("DeliveryWindowStart"))) & Trim(CStr(oRequest("DeliveryWindowEnd")))
	sDelAddrID = Trim(CStr(oRequest("delAddrID")))
	sTransPropsID = Trim(CStr(oRequest("transPropsID")))
	sWizardDeviceID = Trim(CStr(oRequest("wdvid")))
	sWizardAddressName = Trim(CStr(oRequest("wadn")))
	sWizardPhysicalAddress = Trim(CStr(oRequest("wpa")))

	If Err.number <> NO_ERR Then
	    lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "ParseRequestForAddresses", "", "Error setting variables to Request variables", LogLevelError)
	End If

	ParseRequestForAddresses = lErrNumber
	Err.Clear
End Function

Function CleanRequestForAddresses(sDeviceTypeID, sEditAddID, sAction, sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld, sDevice, sPIN, sDeliveryWindow, sDelAddrID, sTransPropsID, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
'********************************************************
'*Purpose: Clean variables of the Address Request
'*Inputs:
'*Outputs:
'********************************************************
Dim lErrNumber

    On Error Resume Next
    lErrNumber = NO_ERR

	sDeviceTypeID = ""
	sEditAddID = ""
	sAction = ""
	sAddressName = ""
	sPhysicalAddress = ""
	sAddressNameVld = ""
	sPhysicalAddressVld = ""
	sDevice = ""
	sPIN = ""
	sDeliveryWindow = ""
	sDelAddrID = ""
	sTransPropsID = ""
	sWizardDeviceID = ""
	sWizardAddressName = ""
	sWizardPhysicalAddress = ""

	CleanRequestForAddresses = lErrNumber
	Err.Clear

End Function

Function ParseRequestForAddressWiz(oRequest, sDeviceTypeID, sAddressName, sPhysicalAddress, sAction, sEditAddID, iAddressWizardStep, sCategoryID, sCategoryName)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'********************************************************
		On Error Resume Next
        Dim lErrNumber

        lErrNumber = NO_ERR

        sDeviceTypeID = ""
        sAddressName = ""
        sPhysicalAddress = ""
        sAction = ""
        sEditAddID = ""
        sCategoryID = ""
        sCategoryName = ""

        sDeviceTypeID = Trim(CStr(oRequest("deviceTypeID")))
        sAddressName = Trim(CStr(oRequest("AddressName")))
        sPhysicalAddress = Trim(CStr(oRequest("PhysicalAddress")))
        sAction = Trim(CStr(oRequest("action")))
        sEditAddID = Trim(CStr(oRequest("editAddID")))
        If Trim(CStr(oRequest("awstep"))) <> "" Then
            iAddressWizardStep = CInt(Trim(CStr(oRequest("awstep"))))
        End If
        sCategoryID = Trim(CStr(oRequest("dtfid")))
        sCategoryName = Trim(CStr(oRequest("dtfn")))

		If Err.number <> NO_ERR Then
		    lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "ParseRequestForAddressWiz", "", "Error setting variables to Request variables", LogLevelError)
        Else
		    If Len(sDeviceTypeID) = 0 Then
		    	lErrNumber = URL_MISSING_PARAMETER
		    End If
		End If

        ParseRequestForAddressWiz = lErrNumber
        Err.Clear
End Function

Function cu_GetUserAddresses(sGetUserAddressesXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs: sGetUserAddressesXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetUserAddresses"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

	lErrNumber = co_GetUserAddresses(sSessionID, sGetUserAddressesXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_GetUserAddresses", LogLevelTrace)
	End If

	cu_GetUserAddresses = lErrNumber
	Err.Clear
End Function

Function RenderAddresses(sDeviceTypeID, sEditAddID, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress, sDeviceTypesXML, sUserAddressesXML, sGetDevicesFromFoldersXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Add error message
'********************************************************
	On Error Resume Next
	Dim oDeviceTypeDOM
	Dim oDeviceTypes
	Dim oCurrentType
	Dim oDisplayFields
	Dim iColWidth
	Dim oCurrentField
	Dim lErrNumber
	Dim oUserAddressesDOM
	Dim oDevices
	Dim oDevice
	Dim bHasAddresses
	Dim bHasAddressesOfType
	Dim oUserAddresses
	Dim oObjectsInFolderDOM
	Dim oPortalAddress

	lErrNumber = NO_ERR
	bHasAddresses = False

	Set oUserAddressesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oUserAddressesDOM.async = False
	If oUserAddressesDOM.loadXML(sUserAddressesXML) = False Then
        lErrNumber = ERR_XML_LOAD_FAILED
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderAddresses", "", "Error while loading userAddresses XML", LogLevelError)
		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(376) & "</b></font>" 'Descriptor: Error retrieving addresses
	End If

	If lErrNumber = NO_ERR Then
	    Set oUserAddresses = oUserAddressesDOM.selectNodes("//oi")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderAddresses", "", "Error while retrieving oi nodes", LogLevelError)
	        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(376) & "</b></font>" 'Descriptor: Error retrieving addresses
	    Else
	        Set oPortalAddress = oUserAddressesDOM.selectSingleNode("//oi[@tp = '" & TYPE_ADDRESS & "' and @id = '" & GetPortalAddress() & "']")
	        If (oUserAddresses.length > 1) And (Not (oPortalAddress Is Nothing)) Then
	        	bHasAddresses = True
	        ElseIf (oUserAddresses.length > 0) And (oPortalAddress Is Nothing) Then
	            bHasAddresses = True
	        Else
	        	If Len(sDeviceTypeID) = 0 Then
	        		Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(396) & "</b></font>" 'Descriptor: No addresses created
	        	End If
	        End If
	    End If
	End If

	Set oPortalAddress = Nothing

    If lErrNumber = NO_ERR Then
        If (bHasAddresses) Or (sDeviceTypeID <> "") Then
	        Set oDeviceTypeDOM = Server.CreateObject("Microsoft.XMLDOM")
	        oDeviceTypeDOM.async = False
	        If oDeviceTypeDOM.loadXML(sDeviceTypesXML) = False Then
                lErrNumber = ERR_XML_LOAD_FAILED
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderAddresses", "", "Error while loading deviceTypes XML", LogLevelError)
	        	Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(376) & "</b></font>" 'Descriptor: Error retrieving addresses
	        Else
	            Set oDeviceTypes = oDeviceTypeDOM.selectNodes("/devicetypes/devicetype")
	            If Err.number <> NO_ERR Then
	                lErrNumber = Err.number
	                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderAddresses", "", "Error while retrieving device nodes", LogLevelError)
                    'add error message
	            End If
	        End If

	        If lErrNumber = NO_ERR Then
	            Set oObjectsInFolderDOM = Server.CreateObject("Microsoft.XMLDOM")
	            oObjectsInFolderDOM.async = False
	            If oObjectsInFolderDOM.loadXML(sGetDevicesFromFoldersXML) = False Then
	                lErrNumber = ERR_XML_LOAD_FAILED
	                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderAddresses", "", "Error while loading sGetDevicesFromFoldersXML", LogLevelError)
                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#cc0000""><b>" & asDescriptors(376) & "</b></font>" 'Descriptor: Error retrieving addresses
	            'Else
	            End If
	        End If
	    End If
	End If

	If lErrNumber = NO_ERR Then
	    If (bHasAddresses) Or (sDeviceTypeID <> "") Then
	        If oDeviceTypes.length > 0 Then
	        	For Each oCurrentType in oDeviceTypes 'for each device type
	        		bHasAddressesOfType = False

	        		Dim oFolders
	        		Dim oFolder

	        		Set oFolders = oCurrentType.selectNodes("dfs/f")
	        		If oFolders.length > 0 Then
	        		    For Each oFolder in oFolders
	        		        Set oDevices = oObjectsInFolderDOM.selectNodes("/mi/in/oi[@id = '" & oFolder.getAttribute("id") & "']/in/oi[@tp = '" & TYPE_DEVICE & "']")
	        		        If oDevices.length > 0 Then
	        		            For Each oDevice In oDevices
	        		                Set oUserAddresses = oUserAddressesDOM.selectNodes("//oi[@dvid = '" & oDevice.getAttribute("id") & "']")
	        		                If oUserAddresses.length > 0 Then
	        		                    bHasAddressesOfType = True
	        					        Exit For
	        		                End If
	        		            Next
	        		        End If
	        		        If bHasAddressesOfType = True Then
	        		            Exit For
	        		        End If
	        		    Next
	        		End If

	        		Set oDevices = Nothing
	        		Set oDevice = Nothing
	        		Set oUserAddresses = Nothing

	        		If bHasAddressesOfType Or (oCurrentType.selectSingleNode("devicetypeID").text = sDeviceTypeID) Then
	        			Set oDisplayFields = oCurrentType.selectNodes("displayfields/field")

	        			Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
	        			Response.Write "<TR>"
	        			Response.Write "<TD WIDTH=""1%"" VALIGN=TOP ALIGN=CENTER><font size=""" & aFontInfo(N_MEDIUM_FONT) & """ face=""Arial""><b>" & oCurrentType.selectSingleNode("name").text & "</b></font><BR /><img src=""" & oCurrentType.selectSingleNode("largeicon").text & """ WIDTH=""" & oCurrentType.selectSingleNode("largeicon").getAttribute("width") & """ HEIGHT=""" & oCurrentType.selectSingleNode("largeicon").getAttribute("height") & """ BORDER=""0"" ALT=""" & oCurrentType.selectSingleNode("name").text & """></TD>"
	        			Response.Write "<TD WIDTH=""99%"" valign=top rowspan=2>"
	        			Response.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH=""100%"">"
	        				        			'Render addresses here (as rows)

	        			Response.Write "<TR bgcolor=""#cccccc"">"

	        			If oDisplayFields.length > 0 Then
	        			iColWidth = 99 \ (oDisplayFields.length)
	        				Response.Write "<TD width=""1%""><img src=""images/1ptrans.gif"" WIDTH=""17"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
	        				For Each oCurrentField In oDisplayFields
	        					Response.Write "<TD width=" & iColWidth & "% height=20><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(oCurrentField.getAttribute("di")) & "</b></font></TD>"
	        				Next
	        				Response.Write "<TD width=""1%"" height=20><img src=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD>"
	        			End If

	        			Response.Write "</TR>"

	        			If bHasAddressesOfType Then
	        			    Call RenderUserAddresses(oCurrentType.selectSingleNode("devicetypeID").text, sEditAddID, sDeviceTypesXML, sUserAddressesXML, oObjectsInFolderDOM, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
	        			End If

	        			If (oCurrentType.selectSingleNode("devicetypeID").text = sDeviceTypeID) And Len(sEditAddID) = 0 Then
	        				Call RenderAddressEditor(sDeviceTypesXML, oCurrentType.selectSingleNode("devicetypeID").text, oDisplayFields.length, "", "", oObjectsInFolderDOM, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
	        			End If

	        			Response.Write "</TABLE>"
	        			Response.Write "</TD>"
	        			Response.Write "</TR>"
	        			Response.Write "</TABLE>"
	        			Response.Write "<BR>"
	        		End If
	        	Next
	        End If
	    End If
	End If

	Set oDeviceTypeDOM = Nothing
	Set oDeviceTypes = Nothing
	Set oCurrentType = Nothing
	Set oDisplayFields = Nothing
	Set oCurrentField = Nothing
	Set oUserAddressesDOM = Nothing
	Set oObjectsInFolderDOM = Nothing
	Set oFolders = Nothing
	Set oFolder = Nothing

	RenderAddresses = lErrNumber
	Err.Clear
End Function

Function RenderAddressEditor(szDeviceTypeXML, szDeviceTypeID, iNumCols, szEditAddressID, szUserAddressesXML, oObjectsInFolderDOM, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim i
	Dim oEditFieldsDOM
	Dim oEditFields
	Dim oFolders
	Dim oCurrentFolder
	Dim oDevices
	Dim oCurrentDevice
	Dim iOverflowCols
	Dim iOverflowRows
	Dim iNextCols
	Dim j
	Dim oUserAddressesDOM
	Dim oEditAddress
	Dim szEditDevice
	Dim lErrNumber
	Dim sHelpLink
	Dim bHasDevices
	Dim sOptions

	lErrNumber = NO_ERR
	bHasDevices = False

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, szDeviceTypeXML, oEditFieldsDOM)
	If lErrNumber = NO_ERR Then
	    Set oEditFields = oEditFieldsDOM.selectNodes("/devicetypes/devicetype[devicetypeID = '" & szDeviceTypeID & "']/editfields/field")
	End If

	If lErrNumber = NO_ERR Then
	    If szEditAddressID <> "" Then
	    	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, szUserAddressesXML, oUserAddressesDOM)
			If lErrNumber = NO_ERR Then
	    	    Set oEditAddress = oUserAddressesDOM.selectSingleNode("/mi/in/oi[@id = '" & szEditAddressID & "']")
	    	    szEditDevice = oEditAddress.getAttribute("dvid")
	    	End If
	    End If
	End If

    If lErrNumber = NO_ERR Then

        'Get the devices (styles) for this address:
	    sOptions = ""
	    Set oFolders = oEditFieldsDOM.selectNodes("/devicetypes/devicetype[devicetypeID = '" & szDeviceTypeID & "']/dfs/f")
	    If oFolders.length > 0 Then
	        For Each oCurrentFolder In oFolders
	            Set oDevices = oObjectsInFolderDOM.selectNodes("/mi/in/oi[@id = '" & oCurrentFolder.getAttribute("id") & "']/in/oi[@id != '" & Application("Portal_Device") & "']")
	            If oDevices.length > 0 Then
	               For Each oCurrentDevice in oDevices
	        	        sOptions = sOptions & "<OPTION "
						If szEditAddressID <> "" Then
	        	        	If oCurrentDevice.getAttribute("id") = szEditDevice Then
	        	        		sOptions = sOptions & "SELECTED "
	        	        	End If
						Else
			                If sWizardDeviceID <> "" Then
	        	        		If oCurrentDevice.getAttribute("id") = sWizardDeviceID Then
	        	        			sOptions = sOptions & "SELECTED "
	        	        		End If
			                Else
	        	        		If oCurrentDevice.getAttribute("id") = Application("Default_Device") Then
	        	        			sOptions = sOptions & "SELECTED "
	        	        		End If
							End If
	        	        End If
	                    sOptions = sOptions & "VALUE=""" & oCurrentDevice.getAttribute("id") & """>" & Server.HTMLEncode(oCurrentDevice.getAttribute("n")) & "</OPTION>"
	                Next
	            End If
	        Next
		End If


        'If an address has no devices, it cannot be edited, show just a message to the user:
        If Len(sOptions) = 0 Then

	    	Response.Write "<FORM NAME=""addressForm"" METHOD=""GET"" ACTION=""addresses.asp"" >"
	    	Response.Write "<TR BGCOLOR=""#ffffcc"">"
	    	Response.Write "<TD><IMG SRC=""images/redtri.gif"" WIDTH=""17"" HEIGHT=""11"" BORDER=""0"" ALT=""""></TD>"
            Response.Write "<TD COLSPAN=""" & iNumCols & """><B><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ COLOR=""#cc0000"" SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(610) & "</FONT></B></TD>" 'Descriptor: No styles available.
	    	Response.Write "<TD VALIGN=""MIDDLE"" ALIGN=""RIGHT"" NOWRAP><INPUT TYPE=""SUBMIT"" NAME=""addrCancel"" value=""" & asDescriptors(120) & """ class=""buttonClass"" /></TD>" 'Descriptor: Cancel
            Response.Write "</TR></FORM>"

        Else

	    	If GetJavaScriptSetting() = "1" Then
	    	    Response.Write "<FORM NAME=""addressForm"" METHOD=""POST"" ACTION=""addresses.asp"" onSubmit=""return validateAddressForm(document.addressForm);"" >"
	    	Else
	    		Response.Write "<FORM NAME=""addressForm"" METHOD=""POST"" ACTION=""addresses.asp"" >"
	    	End If

	    	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""devicetypeID"" value=""" & szDeviceTypeID & """>"

	    	If szEditAddressID <> "" Then
	    		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""action"" value=""edit"">"
	    		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""editAddID"" value=""" & szEditAddressID & """>"
	    		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""transPropsID"" value=""" & oEditAddress.getAttribute("trps") & """>"
	    	Else
	    		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""action"" value=""add"">"
	    	End If

	    	Response.Write "<TR BGCOLOR=""#ffffcc"">"
	    	Response.Write "<TD VALIGN=""TOP""><IMG SRC=""images/redtri.gif"" WIDTH=""17"" HEIGHT=""11"" BORDER=""0"" ALT=""""></TD>"
	    	Response.Write "<TD COLSPAN=""" & iNumCols & """>"

		    Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""2"" WIDTH=""100%"" VALIGN=""TOP"">"
		    Response.Write "<TR>"
	    	For i = 0 to (oEditFields.length - 1)

	    	    'Line break:
	    	    If (i > 0) And ((i mod iNumCols) = 0) Then Response.Write "</TR><TR>"

	    		'Special editing fields:
	    		If oEditFields.item(i).getAttribute("col") = "dvid" Then


	    		    Response.Write "<TD NOWRAP><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(oEditFields.item(i).getAttribute("di")) & ":</B><BR /></FONT>"
	    		    Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
	    		    Response.Write "<TD VALIGN=""TOP""><SELECT CLASS=""pullDownClass"" NAME=""Device"">"
                    Response.Write sOptions
	    		    Response.Write "</SELECT></TD>"

                    Response.Write "<TD VALIGN=""TOP""> "
                    If GetJavaScriptSetting() = "1" Then
                        Response.Write "<INPUT TYPE=IMAGE NAME=""addrHelp"" SRC=""images/questionMark.gif"" BORDER=""0"" ALT="""" onClick=""bValidate=false;"" />"
                    Else
                        Response.Write "<INPUT TYPE=IMAGE NAME=""addrHelp"" SRC=""images/questionMark.gif"" BORDER=""0"" ALT="""" />"
                    End If
                    Response.Write "</TD></TR></TABLE>"
	    		    Response.Write "</TD>"

	    		ElseIf oEditFields.item(i).getAttribute("col") = "cb" Then
                    Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(oEditFields.item(i).getAttribute("di")) & ":</B><BR /></FONT>"
                    If szEditAddressID <> "" Then
                        Call RenderDeliveryWindow(oEditAddress.getAttribute("cb"))
                    Else
                        Call RenderDeliveryWindow("")
                    End If
                    Response.Write "</TD>"

	    		Else
	    			Response.Write "<INPUT TYPE=""HIDDEN"" name=""" & oEditFields.item(i).getAttribute("n") & "Vld"" value=""" & oEditFields.item(i).getAttribute("vld") & """ />"
	    			Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(oEditFields.item(i).getAttribute("di")) & ":</B><BR /> </FONT>"

	    			Response.Write "<INPUT CLASS=""textBoxClass"" MAXLENGTH=""250"" type=""" & oEditFields.item(i).getAttribute("type") & """ name=""" & oEditFields.item(i).getAttribute("n") & """ size=""" & oEditFields.item(i).getAttribute("size") & """"
	    			If szEditAddressID <> "" Then
	    			    If oEditFields.item(i).getAttribute("col") = "cpwd" Then
	    			        Response.Write " VALUE=""" & Server.HTMLEncode(oEditAddress.getAttribute("pwd")) & """"
	    			    Else
	    			        Response.Write " VALUE=""" & Server.HTMLEncode(oEditAddress.getAttribute(oEditFields.item(i).getAttribute("col"))) & """"
	    			    End If

					Else
	    			    If (oEditFields.item(i).getAttribute("col") = "n") And (Len(sWizardAddressName) > 0) Then
	    			        Response.Write " VALUE=""" & Server.HTMLEncode(sWizardAddressName) & """"
	    			    End If
	    			    If (oEditFields.item(i).getAttribute("col") = "v") And (Len(sWizardPhysicalAddress) > 0) Then
	    			        Response.Write " VALUE=""" & Server.HTMLEncode(sWizardPhysicalAddress) & """"
	    			    End If
	    			End If
	    			Response.Write ">"
	    			Response.Write "</TD>"
	    		End If
	    	Next

	    	Response.Write "</TR></TABLE></TD>"
	    	Response.Write "<TD VALIGN=""MIDDLE"" ALIGN=""RIGHT"" NOWRAP>"

	    	If GetJavaScriptSetting() = "1" Then
                Response.Write "<INPUT TYPE=""submit"" name=""addrSave"" value=""" & asDescriptors(59) & """ class=""buttonClass"" onClick=""bValidate=true;"" >&nbsp;" 'Descriptor: Save
                Response.Write "<INPUT TYPE=""submit"" name=""addrCancel"" value=""" & asDescriptors(120) & """ class=""buttonClass"" onClick=""bValidate=false;"" >" 'Descriptor: Cancel
	    	Else
	    	    Response.Write "<INPUT TYPE=""submit"" name=""addrSave"" value=""" & asDescriptors(59) & """ class=""buttonClass"" >&nbsp;" 'Descriptor: Save
                Response.Write "<INPUT TYPE=""submit"" name=""addrCancel"" value=""" & asDescriptors(120) & """ class=""buttonClass"" >" 'Descriptor: Cancel
	    	End If

            Response.Write "</TD>"
            Response.Write "</TR>"

            Response.Write "<TR>"
	    	Response.Write "<TD ALIGN=CENTER COLSPAN=""" & CStr(iNumCols + 2) & """>"
            Response.Write "<DIV STYLE=""display:none;"" class=""validation"" id=""validation""></DIV>"
	    	Response.Write "</TD>"
	    	Response.Write "</TR>"

	    	Response.Write "</FORM>"
	    End If
	End If

	Set oEditFieldsDOM = Nothing
	Set oEditFields = Nothing
	Set oDevices = Nothing
	Set oFolders = Nothing
	Set oCurrentFolder = Nothing
	Set oCurrentDevice = Nothing
	Set oUserAddressesDOM = Nothing
	Set oEditAddress = Nothing

	RenderAddressEditor = lErrNumber
	Err.Clear
End Function

Function DisplayDeliveryWindow(sDeliveryWindowString)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim sOutput
    Dim sTimeFormat
    Dim b24HourClock
    Dim sFromHour
    Dim sUntilHour

    lErrNumber = NO_ERR
    sOutput = ""
    b24HourClock = True
    sTimeFormat = asDescriptors(328) 'Descriptor: HH:MM:SS PM

    If Len(sTimeFormat) = 0 Then sTimeFormat = "HH:MM:SS PM"

    If InStr(1, sTimeFormat, "PM", 0) > 0 Then b24HourClock = False

    sFromHour = Left(sDeliveryWindowString, 2)
    sUntilHour = Right(sDeliveryWindowString, 2)

    If (b24HourClock = False) Then
        If sFromHour = "00" Then
            sFromHour = "12:00 AM"
        ElseIf CInt(sFromHour) < 10 Then
            sFromHour = Right(sFromHour, 1) & ":00 AM"
        ElseIf CInt(sFromHour) < 12 Then
            sFromHour = sFromHour & ":00 AM"
        ElseIf sFromHour = "12" Then
            sFromHour = sFromHour & ":00 PM"
        Else
            sFromHour = CStr(CInt(sFromHour) - 12) & ":00 PM"
        End If

        If sUntilHour = "00" Then
            sUntilHour = "12:00 AM"
        ElseIf CInt(sUntilHour) < 10 Then
            sUntilHour = Right(sUntilHour, 1) & ":00 AM"
        ElseIf CInt(sUntilHour) < 12 Then
            sUntilHour = sUntilHour & ":00 AM"
        ElseIf sUntilHour = "12" Then
            sUntilHour = sUntilHour & ":00 PM"
        Else
            sUntilHour = CStr(CInt(sUntilHour) - 12) & ":00 PM"
        End If
    Else
        sFromHour = sFromHour & ":00"
        sUntilHour = sUntilHour & ":00"
    End If

    sOutput = asDescriptors(86) & " " & sFromHour & " " & asDescriptors(448) & " " & sUntilHour 'Descriptor: From:, Until:

    DisplayDeliveryWindow = sOutput
    Err.Clear
End Function

Function RenderDeliveryWindow(sDeliveryWindowString)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Replace text strings with descriptors
'********************************************************
    On Error Resume Next
    Dim sTimeFormat
    Dim b24HourClock
    Dim i
    Dim sAMPM
    Dim sSelected

    b24HourClock = True
    sTimeFormat = asDescriptors(328) 'Descriptor: HH:MM:SS PM
	If Len(sTimeFormat) = 0 Then sTimeFormat = "HH:MM:SS PM"

	If InStr(1, sTimeFormat, "PM", 0) > 0 Then b24HourClock = False

    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(86) & "</font>" 'Descriptor: From:
    Response.Write "<SELECT CLASS=""pullDownClass"" SIZE=""1"" NAME=""DeliveryWindowStart"">"
        For i=0 To 23
            sSelected = ""
            If sDeliveryWindowString <> "" Then
                If CInt(Left(sDeliveryWindowString, 2)) = i Then
                    sSelected = "SELECTED"
                End If
            End If

            If b24HourClock = True Then
                If i < 10 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""0" & i & """>0" & i & ":00</OPTION>"
                Else
                    Response.Write "<OPTION " & sSelected & " VALUE=""" & i & """>" & i & ":00</OPTION>"
                End If
            Else
                If i = 0 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""00"">12:00 AM</OPTION>"
                ElseIf i < 10 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""0" & i & """>" & i & ":00 AM</OPTION>"
                ElseIf i < 12 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""" & i & """>" & i & ":00 AM</OPTION>"
                ElseIf i = 12 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""" & i & """>" & i & ":00 PM</OPTION>"
                Else
                    Response.Write "<OPTION " & sSelected & " VALUE=""" & i & """>" & CInt(i-12) & ":00 PM</OPTION>"
                End If
            End If
        Next
    Response.Write "</SELECT>"

    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(448) & "</font>" 'Descriptor: Until:
    Response.Write "<SELECT CLASS=""pullDownClass"" SIZE=""1"" NAME=""DeliveryWindowEnd"">"
        For i=0 To 23
            sSelected = ""
            If sDeliveryWindowString <> "" Then
                If CInt(Right(sDeliveryWindowString, 2)) = i Then
                    sSelected = "SELECTED"
                End If
            End If

            If b24HourClock = True Then
                If i < 10 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""0" & i & """>0" & i & ":00</OPTION>"
                Else
                    Response.Write "<OPTION " & sSelected & " VALUE=""" & i & """>" & i & ":00</OPTION>"
                End If
            Else
                If i = 0 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""00"">12:00 AM</OPTION>"
                ElseIf i < 10 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""0" & i & """>" & i & ":00 AM</OPTION>"
                ElseIf i < 12 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""" & i & """>" & i & ":00 AM</OPTION>"
                ElseIf i = 12 Then
                    Response.Write "<OPTION " & sSelected & " VALUE=""" & i & """>" & i & ":00 PM</OPTION>"
                Else
                    Response.Write "<OPTION " & sSelected & " VALUE=""" & i & """>" & CInt(i-12) & ":00 PM</OPTION>"
                End If
            End If
        Next
    Response.Write "</SELECT>"
End Function

Function RenderUserAddresses(szDeviceTypeID, szEditAddressID, szDeviceTypeXML, szUserAddressesXML, oObjectsInFolderDOM, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Add error messages
'********************************************************
	On Error Resume Next
	Dim oDeviceTypeDOM
	Dim oDevices
	Dim oUserAddressesDOM
	Dim oUserAddresses
	Dim oAddress
	Dim oDevice
	Dim oDisplayFields
	Dim oDisplayField
	Dim lErrNumber
	Dim oFolders
	Dim oCurrentFolder

	lErrNumber = NO_ERR

	Set oDeviceTypeDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypeDOM.async = False
	If oDeviceTypeDOM.loadXML(szDeviceTypeXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderUserAddresses", "", "Error while loading deviceTypes XML", LogLevelError)
		'add error message
	Else
	    Set oFolders = oDeviceTypeDOM.selectNodes("//devicetype[devicetypeID = '" & szDeviceTypeID & "']/dfs/f")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderUserAddresses", "", "Error retrieving f nodes", LogLevelError)
	        'add error message
	    End If
	End If

	If lErrNumber = NO_ERR Then
	    If oFolders.length > 0 Then
	        For Each oCurrentFolder in oFolders
	            Set oDevices = oObjectsInFolderDOM.selectNodes("/mi/in/oi[@id = '" & oCurrentFolder.getAttribute("id") & "']/in/oi[@id != '" & Application("Portal_Device") & "']")
	            If oDevices.length > 0 Then
	            	Set oUserAddressesDOM = Server.CreateObject("Microsoft.XMLDOM")
	            	oUserAddressesDOM.async = False
	            	If oUserAddressesDOM.loadXML(szUserAddressesXML) = False Then
	            		lErrNumber = ERR_XML_LOAD_FAILED
	            		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderUserAddresses", "", "Error while loading userAddresses XML", LogLevelError)
	            		'add error message
	            	Else
                        For Each oDevice in oDevices
	            	    	Set oUserAddresses = oUserAddressesDOM.selectNodes("//oi[@dvid = '" & oDevice.getAttribute("id") & "']")
	            	    	If oUserAddresses.length > 0 Then
	            	    		For Each oAddress In oUserAddresses
	            	    			Set oDisplayFields = oDeviceTypeDOM.selectNodes("//devicetype[devicetypeID = '" & szDeviceTypeID & "']/displayfields/field")

	            	    			If szEditAddressID = oAddress.getAttribute("id") Then
	            	    				lErrNumber = RenderAddressEditor(szDeviceTypeXML, szDeviceTypeID, oDisplayFields.length, szEditAddressID, szUserAddressesXML, oObjectsInFolderDOM, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
	            	    				If lErrNumber <> NO_ERR Then
	            	    				    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderUserAddresses", "", "Error while calling RenderAddressEditor", LogLevelTrace)
	            	    				    'add error message?
	            	    				End If
	            	    			Else
	            	    				Response.Write "<TR>"
	            	    				Response.Write "<TD></TD>"
	            	    				If oDisplayFields.length > 0 Then
	            	    					For Each oDisplayField In oDisplayFields
	            	    						If oDisplayField.getAttribute("col") = "dvid" Then
	            	    							Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & oDevice.getAttribute("n") & "</font></TD>"
	            	    						ElseIf oDisplayField.getAttribute("col") = "cb" Then
	            	    						    Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & DisplayDeliveryWindow(oAddress.getAttribute("cb")) & "</font></TD>"
	            	    						Else
	            	    							Response.Write "<TD>"
	            	    							Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & Server.HTMLEncode(oAddress.getAttribute(oDisplayField.getAttribute("col"))) & "</font>"
	            	    							Response.Write "</TD>"
	            	    						End If
	            	    					Next
	            	    				End If
	            	    				'Go to confirm_delete.asp
	            	    				Response.Write "<TD NOWRAP align=right><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><A HREF=""addresses.asp?action=edit&devicetypeID=" & szDeviceTypeID & "&editAddID=" & oAddress.getAttribute("id") & """>" & asDescriptors(353) & "</A> " 'Descriptor: Edit
	            	    				Response.Write "<A HREF=""confirm_delete.asp?formPage=addresses.asp&deleteType=1&cancelButton=addrCancel&action=delete&delAddrID=" & oAddress.getAttribute("id") & """>" & asDescriptors(249) & "</A></font></TD>" 'Descriptor: Delete
	            	    				Response.Write "</TR>"
	            	    			End If
	            	    		Next
	            	    	End If
	            	    Next
	            	End If
	            End If
	        Next
	    End If
	End If

	Set oDeviceTypeDOM = Nothing
	Set oDevices = Nothing
	Set oDevice = Nothing
	Set oUserAddressesDOM = Nothing
	Set oUserAddresses = Nothing
	Set oAddress = Nothing
	Set oDisplayFields = Nothing
	Set oDisplayField = Nothing
	Set oFolders = Nothing
	Set oCurrentFolder = Nothing

	RenderUserAddresses = lErrNumber
	Err.Clear
End Function

Function RenderNewAddressBox(sDeviceTypesXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: ADD Error messages
'*QUESTION: error handling in loop?
'********************************************************
	On Error Resume Next
	Dim oDeviceTypesDOM
	Dim oDeviceTypes
	Dim oCurrentDeviceType
	Dim lErrNumber
	lErrNumber = NO_ERR

	Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
	oDeviceTypesDOM.async = False
	If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderNewAddressBox", "", "Error while loading deviceTypes XML", LogLevelError)
	    'add error message
	End If

	If lErrNumber = NO_ERR Then
	    Set oDeviceTypes = oDeviceTypesDOM.selectNodes("//devicetype")
	    If Err.number <> NO_ERR Then
	        lErrNumber = Err.number
	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderNewAddressBox", "", "Error reading deviceType nodes", LogLevelError)
	        'add error message
	    Else
	        If oDeviceTypes.length > 0 Then
	        	Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
	        	Response.Write "<TR><TD bgcolor=""666666"">"
	        	Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"
	        	Response.Write "<TR><TD bgcolor=""cccccc"" width=400 height=20 COLSPAN=4><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""000000"" size=" & aFontInfo(N_SMALL_FONT) & "><b>&nbsp;&nbsp;" & asDescriptors(374) & "</b></font></TD></TR>" 'Descriptor: Create new address
	        	Response.Write "<TR><TD COLSPAN=4 HEIGHT=5 bgcolor=ffffff><img src=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD></TR>"
	        	Response.Write "<TR><TD bgcolor=ffffff align=CENTER COLSPAN=4>"
	        	Response.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>"
	        	Response.Write "<TR><TD></TD>"
	        	For Each oCurrentDeviceType In oDeviceTypes
	        		Response.Write "<TD><A HREF=""addresses.asp?action=new&devicetypeID=" & oCurrentDeviceType.selectSingleNode("devicetypeID").text & """><img src=""" & oCurrentDeviceType.selectSingleNode("icon").text & """ WIDTH=""" & oCurrentDeviceType.selectSingleNode("icon").getAttribute("width") & """ HEIGHT=""" & oCurrentDeviceType.selectSingleNode("icon").getAttribute("height") & """ BORDER=""0"" ALT=""" & oCurrentDeviceType.selectSingleNode("name").text & """></A></TD>"
	        		Response.Write "<TD NOWRAP><A HREF=""addresses.asp?action=new&devicetypeID=" & oCurrentDeviceType.selectSingleNode("devicetypeID").text & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oCurrentDeviceType.selectSingleNode("name").text & "</font></A></TD>"
	        		Response.Write "<TD>&nbsp;|&nbsp;</TD>"
	        	Next
	        	Response.Write "</TR></TABLE></TD></TR>"
	        	Response.Write "<TR><TD COLSPAN=4 HEIGHT=5 bgcolor=ffffff><img src=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" ALT=""""></TD></TR>"
	        	Response.Write "</TABLE></TD></TR></TABLE>"
	        End If
	    End If
	End If


	Set oDeviceTypesDOM = Nothing
	Set oDeviceTypes = Nothing
	Set oCurrentDeviceType = Nothing

	RenderNewAddressBox = lErrNumber
	Err.Clear
End Function

Function GetDeviceDescForFolder_Wizard(sCategoryID, sDeviceDescXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "GetDeviceDescForFolder_Wizard"
	Dim lErrNumber
	Dim sSessionID
    Dim asFolderID(0)
    Dim bBrowseSubFolders

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
	bBrowseSubFolders = True

    asFolderID(0) = sCategoryID

	If lErrNumber = NO_ERR Then
	    lErrNumber = co_GetDevicesInFolders(sSessionID, asFolderID, N_DEVICE_DESC_HTML, bBrowseSubFolders, sDeviceDescXML)
	    If lErrNumber <> NO_ERR Then
	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_GetDevicesInFolders", LogLevelTrace)
	    End If
	End If

	GetDeviceDescForFolder_Wizard = lErrNumber
	Err.Clear
End Function


Function cu_GetDevicesInFolders(sDeviceTypesXML, sGetDevicesInFoldersXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetDevicesInFolders"
	Dim lErrNumber
	Dim sSessionID
    Dim asFolderID()
    Dim bBrowseSubFolders
    Dim oDeviceTypesDOM
    Dim oFolders
    Dim i

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
	bBrowseSubFolders = True

    Set oDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")
    oDeviceTypesDOM.async = False
    If oDeviceTypesDOM.loadXML(sDeviceTypesXML) = False Then
        lErrNumber = ERR_XML_LOAD_FAILED
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", PROCEDURE_NAME, "", "Error loading sDeviceTypesXML", LogLevelError)
    Else
        Set oFolders = oDeviceTypesDOM.selectNodes("//f")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", PROCEDURE_NAME, "", "Error retrieving f nodes", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
        If oFolders.length > 0 Then
            Redim asFolderID(oFolders.length - 1)
            For i=0 To (oFolders.length - 1)
                asFolderID(i) = oFolders.item(i).getAttribute("id")
            Next
        End If
    End If

    Set oFolders = Nothing
    Set oDeviceTypesDOM = Nothing

	If lErrNumber = NO_ERR Then
	    lErrNumber = co_GetDevicesInFolders(sSessionID, asFolderID, N_DEVICE_DESC_NONE, bBrowseSubFolders, sGetDevicesInFoldersXML)
	    If lErrNumber <> NO_ERR Then
	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", PROCEDURE_NAME, "", "Error while calling co_GetDevicesInFolders", LogLevelTrace)
	    End If
	End If

	cu_GetDevicesInFolders = lErrNumber
	Err.Clear
End Function

Function GetDeviceTypesProperties_AddressWizard(sDeviceTypeID, sDeviceTypesXML, sDeviceTypeName, sDeviceTypeImage, asDTFolders)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDeviceTypesDOM
	Dim oDeviceType
	Dim oDTFolders
	Dim i

	lErrNumber = NO_ERR

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sDeviceTypesXML, oDeviceTypesDOM)
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "GetDeviceTypesProperties_AddressWizard", "", "Error loading sDeviceTypesXML", LogLevelError)
	Else
	    Set oDeviceType = oDeviceTypesDOM.selectSingleNode("//devicetype[devicetypeID = '" & sDeviceTypeID & "']")
	    If Not (oDeviceType Is Nothing) Then
	        sDeviceTypeName = oDeviceType.selectSingleNode("name").text
	        sDeviceTypeImage = oDeviceType.selectSingleNode("largeicon").text

	        Set oDTFolders = oDeviceType.selectNodes("dfs/f")
	        If oDTFolders.length > 0 Then
	            Redim asDTFolders(oDTFolders.length - 1, 1)
	            For i=0 to (oDTFolders.length - 1)
	                asDTFolders(i, 0) = oDTFolders.item(i).getAttribute("id")
	                asDTFolders(i, 1) = oDTFolders.item(i).getAttribute("n")
	            Next
	        End If
	    End If
	End If

	Set oDeviceTypesDOM = Nothing
	Set oDeviceType = Nothing
	Set oDTFolders = Nothing

	GetDeviceTypesProperties_AddressWizard = lErrNumber
	Err.Clear
End Function

Function RenderDeviceDescriptions(sDeviceDescXML, sDeviceTypeID, sAction, sEditAddID, sAddressName, sPhysicalAddress)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oDevicesDOM
	Dim oDevices
	Dim i
	Dim iNumDevices
	Dim sHTMLDescription

	lErrNumber = NO_ERR

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sDeviceDescXML, oDevicesDOM)
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "RenderDeviceDescriptions", "", "Error loading sDeviceDescXML", LogLevelError)
	Else
	    Set oDevices = oDevicesDOM.selectNodes("//oi[@tp = '" & TYPE_DEVICE & "']")
	    If oDevices.length > 0 Then
	        iNumDevices = oDevices.length
	        Response.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH=""100%"">"
	        For i=0 to (iNumDevices - 1)
	            If (i+1) Mod 3 = 1 Then
	                Response.Write "<TR><TD ALIGN=CENTER>"
	            Else
	                Response.Write "<TD ALIGN=CENTER>"
	            End If

	            Response.Write "<A HREF=""addresses.asp?wadn=" & Server.URLEncode(sAddressName) & "&wpa=" & Server.URLEncode(sPhysicalAddress) & "&wdvid=" & oDevices.item(i).getAttribute("id") & "&deviceTypeID=" & sDeviceTypeID & "&action=" & sAction & "&editAddID=" & sEditAddID & """>"
	            sHTMLDescription = oDevices.item(i).getAttribute("hdes")
	            If sHTMLDescription <> "" Then
	                Response.Write Replace(Replace(Replace(sHTMLDescription, "&lt;", "<"), "&gt;", ">"), "&amp;", "&")
	            End If
	            Response.Write "</A>"

	            Response.Write "<BR />"
	            Response.Write "<A HREF=""addresses.asp?wadn=" & Server.URLEncode(sAddressName) & "&wpa=" & Server.URLEncode(sPhysicalAddress) & "&wdvid=" & oDevices.item(i).getAttribute("id") & "&deviceTypeID=" & sDeviceTypeID & "&action=" & sAction & "&editAddID=" & sEditAddID & """>"
	            Response.Write "<b><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oDevices.item(i).getAttribute("n") & "</font></b>"
	            Response.Write "</A>"
	            If (i+1) Mod 3 = 1 Then
	                Response.Write "</TD>"
	            ElseIf (i+1) Mod 3 = 2 Then
	                Response.Write "</TD>"
	            Else
	                Response.Write "</TD></TR>"
	            End If
	        Next
	        Response.Write "</TABLE>"
	    Else
	        Response.Write asDescriptors(608) 'Descriptor: There are no device styles in this category.
	    End If
	End If

	Set oDevicesDOM = Nothing
	Set oDevices = Nothing

	'RenderDeviceDescriptions = lErrNumber
	Err.Clear
End Function

Function CheckForSubFolders_AddressWizard(sGetFolderContentsXML, asDTFolders)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oFolderContentsDOM
	Dim oFolders
	Dim i

	lErrNumber = NO_ERR

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetFolderContentsXML, oFolderContentsDOM)
	If lErrNumber <> NO_ERR Then
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "AddressesCuLib.asp", "CheckForSubFolders_AddressWizard", "", "Error loading sGetFolderContentsXML", LogLevelError)
	Else
	    Set oFolders = oFolderContentsDOM.selectNodes("//oi[@tp = '" & TYPE_FOLDER & "']")
	    If oFolders.length > 0 Then
	        Redim asDTFolders(oFolders.length - 1, 1)
	        For i=0 to (oFolders.length - 1)
	            asDTFolders(i, 0) = oFolders.item(i).getAttribute("id")
	            asDTFolders(i, 1) = oFolders.item(i).getAttribute("n")
	        Next
	    End If
	End If

	Set oFolderContentsDOM = Nothing
	Set oFolders = Nothing

	CheckForSubFolders_AddressWizard = lErrNumber
	Err.Clear
End Function
%>