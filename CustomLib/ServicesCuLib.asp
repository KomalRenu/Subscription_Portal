<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Function ParseRequestForServices(oRequest, sFolderID)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'*TO DO: Set folderID according to an "ROOT FOLDER" setting
	'*QUESTION: Should SetServiceViewMode be called here?
	'********************************************************
		On Error Resume Next
		Dim lErrNumber
        Dim sServiceViewMode

        lErrNumber = NO_ERR

		sFolderID = ""
		sServiceViewMode = ""

		sFolderID = Trim(CStr(oRequest("folderID")))
		sServiceViewMode = Trim(CStr(oRequest("svm")))

		If Err.number <> NO_ERR Then
		    lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ServicesCuLib.asp", "ParseRequestForServices", "", "Error setting variables equal to Request variables", LogLevelError)
		End If

        If Len(sServiceViewMode) > 0 Then
            Call SetServiceViewMode(sServiceViewMode)
        End If

		ParseRequestForServices = lErrNumber
		Err.Clear
	End Function

	Function RenderPath_Services(sGetFolderContentsXML)
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
        Dim oRootFolder
        Dim iNumFolders
        Dim i
        lErrNumber = NO_ERR

        Set oContentsDOM = Server.CreateObject("Microsoft.XMLDOM")
		oContentsDOM.async = False
		If oContentsDOM.loadXML(sGetFolderContentsXML) = False Then
			lErrNumber = ERR_XML_LOAD_FAILED
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ServicesCuLib.asp", "RenderPath_Services", "", "Error loading folderContents.xml file", LogLevelError)
			'add error message
        Else
			Set oRootFolder = oContentsDOM.selectSingleNode("//a/fd[@id='" & APP_ROOT_FOLDER & "']").parentNode
            iNumFolders = CInt(oRootFolder.selectNodes(".//a").length)
            If Err.number <> NO_ERR Then
                lErrNumber = Err.number
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ServicesCuLib.asp", "RenderPath_Services", "", "Error retrieving oi nodes", LogLevelError)
                'add error message
            End If
		End If

        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"
        Response.Write asDescriptors(26) & " " 'Descriptor: You are here:
        If lErrNumber = NO_ERR Then
            If iNumFolders > 0 Then
                Response.Write "<A HREF=""services.asp""><font color=""#000000"">" & asDescriptors(362) & "</font></A>" 'Descriptor: Services
                Set oFolder = oRootFolder
                For i=1 To iNumFolders
                    Set oFolder = oFolder.selectSingleNode("a")
                    Response.Write " > "
                    If i = iNumFolders Then
                        Response.Write "<b>" & oFolder.selectSingleNode("fd").getAttribute("n") & "</b>"
                    Else
                        Response.Write "<A HREF=""services.asp?folderID=" & oFolder.selectSingleNode("fd").getAttribute("id") & """><font color=""#0000"">" & oFolder.selectSingleNode("fd").getAttribute("n") & "</font></A>"
                    End If
                Next
            Else
                Response.Write "<b>" & asDescriptors(362) & "</b>" 'Descriptor: Services
            End If
        Else
            'add handling
        End If
        Response.Write "</font>"

        Set oContentsDOM = Nothing
        Set oFolder = Nothing
        Set oRootFolder = Nothing

        RenderPath_Services = lErrNumber
        Err.Clear
	End Function

	Function RenderList_Services(sFolderID, sGetFolderContentsXML)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'********************************************************
        On Error Resume Next
        Dim lErrNumber
        Dim oContentsDOM
        Dim oContents
        Dim oItem
        lErrNumber = NO_ERR

        Set oContentsDOM = Server.CreateObject("Microsoft.XMLDOM")
		oContentsDOM.async = False
		If oContentsDOM.loadXML(sGetFolderContentsXML) = False Then
			lErrNumber = ERR_XML_LOAD_FAILED
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ServicesCuLib.asp", "RenderList_Services", "", "Error loading folderContents.xml file", LogLevelError)
			'add error message
        Else
            Set oContents = oContentsDOM.selectNodes("/mi/fct/oi[@tp = '" & TYPE_FOLDER & "' $or$ @tp = '" & TYPE_SERVICE & "']")
            If Err.number <> NO_ERR Then
                lErrNumber = Err.number
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ServicesCuLib.asp", "RenderList_Services", "", "Error retrieving oi nodes", LogLevelError)
                'add error message
            End If
		End If

        If lErrNumber = NO_ERR Then
            Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"
            Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#99ccff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"
            Response.Write "<TR BGCOLOR=""#6699cc"">"
            Response.Write "<TD>&nbsp;</TD>"
            Response.Write "<TD>&nbsp;&nbsp;</TD>"
            Response.Write "<TD NOWRAP=""1""><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#ffffff"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(306) & "</b></font></TD>" 'Descriptor: Name
            Response.Write "<TD>&nbsp;&nbsp;</TD>"
            Response.Write "<TD>&nbsp;&nbsp;</TD>"
            Response.Write "<TD NOWRAP=""1""><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#ffffff"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(34) & "</b></font></TD>" 'Descriptor: Modified
            Response.Write "<TD>&nbsp;&nbsp;</TD>"
            Response.Write "<TD>&nbsp;&nbsp;</TD>"
            Response.Write "<TD NOWRAP=""1""><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#ffffff"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(22) & "</b></font></TD>" 'Descriptor: Description
            Response.Write "<TD>&nbsp;&nbsp;</TD>"
            Response.Write "<TD>&nbsp;&nbsp;</TD>"
            Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#ffffff"" size=""" & aFontInfo(N_SMALL_FONT) & """><b></b></font></TD>"
            Response.Write "<TD>&nbsp;&nbsp;</TD>"
            Response.Write "</TR>"
            Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#003366""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"
            If oContents.length > 0 Then
                For Each oItem in oContents
                    If (oItem.getAttribute("tp") = TYPE_FOLDER) Or (oItem.getAttribute("act") = "1") Then
                        Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"
                        Response.Write "<TR>"
                        Response.Write "<TD>"
                        If oItem.getAttribute("tp") = TYPE_FOLDER Then
                            Response.Write "<A HREF=""services.asp?folderID=" & oItem.getAttribute("id") & """><IMG SRC=""images/folder2.gif"" HEIGHT=""16"" WIDTH=""16"" BORDER=""0"" ALT="""" /></A>"
                        ElseIf oItem.getAttribute("tp") = TYPE_SERVICE Then
                            Response.Write "<A HREF=""subscribe.asp?serviceID=" & oItem.getAttribute("id") & "&folderID=" & sFolderID & "&serviceName=" & Server.URLEncode(oItem.getAttribute("n")) & """><IMG SRC=""images/report2.gif"" HEIGHT=""16"" WIDTH=""16"" BORDER=""0"" ALT="""" /></A>"
                        End If
                        Response.Write "</TD>"
                        Response.Write "<TD></TD>"
                        Response.Write "<TD>"
                        If oItem.getAttribute("tp") = TYPE_FOLDER Then
                            Response.Write "<A HREF=""services.asp?folderID=" & oItem.getAttribute("id") & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & oItem.getAttribute("n") & "</b></font></A>"
                        ElseIf oItem.getAttribute("tp") = TYPE_SERVICE Then
                            Response.Write "<A HREF=""subscribe.asp?serviceID=" & oItem.getAttribute("id") & "&folderID=" & sFolderID & "&serviceName=" & Server.URLEncode(oItem.getAttribute("n")) & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & oItem.getAttribute("n") & "</b></font></A>"
                        End If
                        Response.Write "</TD>"
                        Response.Write "<TD></TD>"
                        Response.Write "<TD></TD>"
                        Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & DisplayDateAndTime(CDate(oItem.getAttribute("mdt")), "") & "</font>"
                        Response.Write "</TD>"
                        Response.Write "<TD></TD>"
                        Response.Write "<TD></TD>"
                        Response.Write "<TD>"
                        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oItem.getAttribute("des") & "</font>"
                        Response.Write "</TD>"
                        Response.Write "<TD></TD>"
                        Response.Write "<TD></TD>"
                        Response.Write "<TD NOWRAP>"
                        If oItem.getAttribute("tp") = TYPE_SERVICE Then
                            Response.Write "<A TITLE=""" & asDescriptors(627) & """ HREF=""subscriptions.asp?serviceID=" & oItem.getAttribute("id") & "&folderID=" & sFolderID & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(354) & "</font></A>" 'Descriptor: Subscriptions
                        End If
                        Response.Write "</TD>"
                        Response.Write "<TD></TD>"
                        Response.Write "</TR>"
                        Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""2"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"
                        Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#99ccff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"
                    End If
                Next
            Else
                'add handling
            End If
            Response.Write "</TABLE>"
        End If

        Set oContentsDOM = Nothing
        Set oContents = Nothing
        Set oItem = Nothing

        RenderList_Services = lErrNumber
        Err.Clear
	End Function

    Function RenderLargeIcon_Services(sFolderID, sGetFolderContentsXML)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'********************************************************
        On Error Resume Next
        Dim lErrNumber
        Dim oContentsDOM
        Dim oContents
        Dim iNumContents
        Dim i
        lErrNumber = NO_ERR

        Set oContentsDOM = Server.CreateObject("Microsoft.XMLDOM")
		oContentsDOM.async = False
		If oContentsDOM.loadXML(sGetFolderContentsXML) = False Then
			lErrNumber = ERR_XML_LOAD_FAILED
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ServicesCuLib.asp", "RenderLargeIcon_Services", "", "Error loading folderContents.xml file", LogLevelError)
			'add error message
        Else
            Set oContents = oContentsDOM.selectNodes("/mi/fct/oi[@tp = '" & TYPE_FOLDER & "' $or$ (@tp = '" & TYPE_SERVICE & "' $and$ @act = '1')]")
            If Err.number <> NO_ERR Then
                lErrNumber = Err.number
                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "ServicesCuLib.asp", "RenderLargeIcon_Services", "", "Error retrieving oi nodes", LogLevelError)
                'add error message
            End If
		End If

        If lErrNumber = NO_ERR Then
            If oContents.length > 0 Then
                iNumContents = oContents.length
                For i=0 To (iNumContents - 1)
	    	    	If (i+1) Mod 2 = 1 Then
	    	    		Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
	    	    		Response.Write "<TR><TD VALIGN=TOP WIDTH=""50%"">"
	    	    	Else
	    	    		Response.Write "<TD VALIGN=TOP WIDTH=""50%"">"
	    	    	End If
	    	    	Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>"
	    	    	Response.Write "<TR><TD COLSPAN=2 VALIGN=TOP>"
	    	    	If oContents.item(i).getAttribute("tp") = TYPE_FOLDER Then
	    	    	    Response.Write "<A HREF=""services.asp?folderID=" & oContents.item(i).getAttribute("id") & """ STYLE=""text-decoration:none""><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """ color=""#cc0000""><b>" & oContents.item(i).getAttribute("n") & "</b></font></A>"
	    	    	ElseIf oContents.item(i).getAttribute("tp") = TYPE_SERVICE Then
	    	    	    Response.Write "<A HREF=""subscribe.asp?serviceID=" & oContents.item(i).getAttribute("id") & "&folderID=" & sFolderID & "&serviceName=" & Server.URLEncode(oContents.item(i).getAttribute("n")) & """ STYLE=""text-decoration:none""><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_MEDIUM_FONT) & """ color=""#cc0000""><b>" & oContents.item(i).getAttribute("n") & "</b></font></A>"
	    	    	End If
	    	    	Response.Write "</TD></TR>"
	    	    	Response.Write "<TR>"
	    	    	Response.Write "<TD WIDTH=""1%"">"
	    	    	If oContents.item(i).getAttribute("tp") = TYPE_FOLDER Then
	    	    	    Response.Write "<A HREF=""services.asp?folderID=" & oContents.item(i).getAttribute("id") & """><img src=""images/folder_big.gif"" HEIGHT=""76"" WIDTH=""60"" BORDER=""0"" ALT=""""></A>"
	    	    	ElseIf oContents.item(i).getAttribute("tp") = TYPE_SERVICE Then
	    	    	    Response.Write "<A HREF=""subscribe.asp?serviceID=" & oContents.item(i).getAttribute("id") & "&folderID=" & sFolderID & "&serviceName=" & Server.URLEncode(oContents.item(i).getAttribute("n")) & """><img src=""images/report_big.gif"" HEIGHT=""57"" WIDTH=""60"" BORDER=""0"" ALT=""""></A>"
	    	    	End If
	    	    	Response.Write "</TD>"

	    	    	Response.Write "<TD VALIGN=""TOP"" WIDTH=""99%"">"
	    	    	Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"
	    	    	Response.Write oContents.item(i).getAttribute("des") & "<BR />"
	    	    	Response.Write "<b>" & asDescriptors(34) & ": </b>" 'Descriptor: Modified
	    	    	If IsDate(oContents.item(i).getAttribute("mdt")) Then
	    	    	    Response.Write DisplayDateAndTime(CDate(oContents.item(i).getAttribute("mdt")), "")
	    	    	Else
	    	    	    Response.Write asDescriptors(359) 'Descriptor: Information not available
	    	    	End If
	    	    	If oContents.item(i).getAttribute("tp") = TYPE_SERVICE Then
	    	    	    Response.Write "<BR /><A TITLE=""" & asDescriptors(627) & """ HREF=""subscriptions.asp?serviceID=" & oContents.item(i).getAttribute("id") & "&folderID=" & sFolderID & """><font color=""#000000"">" & asDescriptors(354) & "</font></A>" 'Descriptor: Subscriptions
	    	    	End If
	    	    	Response.Write "</font>"
	    	    	Response.Write "</TD>"
	    	    	Response.Write "</TR>"
	    	    	Response.Write "</TABLE>"
	    	    	If (i+1) Mod 2 = 1 Then
	    	    		Response.Write "</TD>"
	    	    		If i = (iNumContents-1) Then
	    	    			Response.Write "<TD></TD></TR></TABLE>"
	    	    		End If
	    	    	Else
	    	    		Response.Write "</TD></TR></TABLE><BR />"
	    	    	End If
	    	    Next
            Else
                'add handling
            End If
        End If

        Set oContentsDOM = Nothing
        Set oContents = Nothing

        RenderLargeIcon_Services = lErrNumber
        Err.Clear
	End Function
%>