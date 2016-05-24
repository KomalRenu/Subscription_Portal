<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/SubscriptionsCoLib.asp" -->
<%
Function ParseRequestForSubscriptions(oRequest, sServiceID, sFolderID, sDeliv_SortOrder, sDeliv_OrderBy, sRep_SortOrder, sRep_OrderBy)
'********************************************************
'*Purpose:
'*Inputs: oRequest
'*Outputs: sReqUserSubs
'*TO DO: Put SortOrder and OrderBy into constants?
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim sSubscriptionViewMode

	lErrNumber = NO_ERR

	sSubscriptionViewMode = ""
	sServiceID = ""
	sFolderID = ""
	sDeliv_SortOrder = ""
	sDeliv_OrderBy = ""
	sRep_SortOrder = ""
	sRep_OrderBy = ""

	sSubscriptionViewMode = Trim(CStr(oRequest("suvm")))
	sServiceID = Trim(CStr(oRequest("serviceID")))
	sFolderID = Trim(CStr(oRequest("folderID")))
	sDeliv_SortOrder = Trim(CStr(oRequest("dSortOrder")))
	sDeliv_OrderBy = Trim(CStr(oRequest("dOrderBy")))
	sRep_SortOrder = Trim(CStr(oRequest("rSortOrder")))
	sRep_OrderBy = Trim(CStr(oRequest("rOrderBy")))

	If Err.number <> NO_ERR Then
	    lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "ParseRequestForSubscriptions", "", "Error setting variables equal to Request variables", LogLevelError)
	End If

	If Len(CStr(sSubscriptionViewMode)) > 0 Then
        Call SetSubscriptionViewMode(sSubscriptionViewMode)
    End If

    If Len(sDeliv_OrderBy) = 0 Then
        sDeliv_OrderBy = GetDeliveryOrderBy()
    End If
    If Len(sDeliv_SortOrder) = 0 Then
        sDeliv_SortOrder = GetDeliverySortOrder()
    End If

    Call SetDeliverySorting(sDeliv_OrderBy, sDeliv_SortOrder)

    If Len(sRep_OrderBy) = 0 Then
        sRep_OrderBy = GetReportsOrderBy()
    End If

    If Len(sRep_SortOrder) = 0 Then
        sRep_SortOrder = GetReportsSortOrder()
    End If

    Call SetReportsSorting(sRep_OrderBy, sRep_SortOrder)

	ParseRequestForSubscriptions = lErrNumber
End Function

Function RenderList_Reports(sSortOrder, sOrderBy, sServiceID, sFolderID, sSubscriptionsXML, sGetAvailableSubscriptionsXML)
'********************************************************
'*Purpose: Renders list of subscriptions to Portal Address
'*Inputs: sSubscriptionsXML
'*Outputs:
'*TO DO: Add Error messages
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "RenderList_Reports"
    Const ACTIVE_COLOR = "#000000"
    Const INACTIVE_COLOR = "#aaaaaa"
    Dim oSubsDOM
	Dim oSubs
	Dim oCurrentSub
	Dim sSchedule
	Dim sService
	Dim oSchedules
	Dim oCurrentSchedule
	Dim sFontColor
	Dim lErrNumber
	Dim sSortIcon
	Dim sSortALT
	Dim sSortURL
	Dim oAvailSubsDOM
	Dim sBoldStart
	Dim sBoldEnd

	lErrNumber = NO_ERR

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sSubscriptionsXML, oSubsDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", PROCEDURE_NAME, "", "Error loading sSubscriptionsXML", LogLevelError)
		Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#cc0000""><B>" & asDescriptors(405) & "</B></FONT>" 'Descriptor: Error retrieving reports
	End If

    If lErrNumber = NO_ERR Then
        lErrNumber = ConvertXMLForSorting_Deliveries(aConnectionInfo, oSubsDOM)
    End If

    'Added for sorting
    If lErrNumber = NO_ERR Then
    	Dim asDictionary(1,1)
    	asDictionary(0,0) = "OrderBy"
        asDictionary(1,0) = sOrderBy
        asDictionary(0,1) = "SortOrder"
        asDictionary(1,1) = sSortOrder
        lErrNumber = AddInputsToXML(aConnectionInfo, oSubsDOM, asDictionary)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", PROCEDURE_NAME, "", "Error in call to AddInputsToXML function", LogLevelTrace)
            Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#cc0000""><B>" & asDescriptors(405) & "</B></FONT>" 'Descriptor: Error retrieving reports
        Else
            lErr = SortDeliveries(aConnectionInfo, "SortDeliveries_Subs.xsl", oSubsDOM)
        End If
    End If

    If lErrNumber = NO_ERR Then
        If Len(sServiceID) > 0 Then
            Set oSubs = oSubsDOM.selectNodes("//sub[@svid = '" & sServiceID & "' and @adid = '" & CStr(GetPortalAddress()) & "']")
        Else
            Set oSubs = oSubsDOM.selectNodes("//sub[@adid = '" & CStr(GetPortalAddress()) & "']")
        End If

        Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" WIDTH=""100%"">"
        Response.Write "<FORM METHOD=""POST"" ACTION=""confirm_delete.asp"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""formPage"" VALUE=""modify_subscription.asp"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""deleteType"" VALUE=""2"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""cancelButton"" VALUE=""subsCancel"" />"
    	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""action"" VALUE=""delete"" />"
    	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""serviceID"" VALUE=""" & sServiceID & """ />"
    	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""folderID"" VALUE=""" & sFolderID & """ />"
        Response.Write "<TBODY>"
        Response.Write "<TR BGCOLOR=""#cccc99"">"
		Response.Write "<TD>&nbsp;</TD>"
        Response.Write "<TD><IMG ALT="""" BORDER=""0"" HEIGHT=""10"" SRC=""Images/active_status.gif"" WIDTH=""19""></TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"">"
		    If StrComp(sOrderBy, "SERVICE", vbTextCompare) = 0 Then
		        If StrComp(sSortOrder, "ASC", vbTextCompare) = 0 Then
		            sSortIcon = "images/sort_asc.gif"
		            sSortALT = asDescriptors(108) 'Descriptor: Sort descending
		            sSortURL = "subscriptions.asp?rSortOrder=DESC&rOrderBy=SERVICE"
		            If Len(sServiceID) > 0 Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        Else
		            sSortIcon = "images/sort_desc.gif"
		            sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		            sSortURL = "subscriptions.asp?rSortOrder=ASC&rOrderBy=SERVICE"
		            If Len(sServiceID) > 0 Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        End If
		    Else
		        sSortIcon = "images/sort_row.gif"
		        sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		        sSortURL = "subscriptions.asp?rSortOrder=ASC&rOrderBy=SERVICE"
		        If Len(sServiceID) > 0 Then
		            sSortURL = sSortURL & "&serviceID=" & sServiceID
		        End If
		    End If
			Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """><FONT COLOR=""#000000""><B>" & asDescriptors(366) & "</B></FONT></FONT> <A HREF=""" & sSortURL & """><IMG SRC=""" & sSortIcon & """ WIDTH=""17"" HEIGHT=""8"" BORDER=""0"" ALT=""" & sSortALT & """ /></A>" 'Descriptor: Service
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"">"
		    If sOrderBy = "SCHEDULE" Then
		        If sSortOrder = "ASC" Then
		            sSortIcon = "images/sort_asc.gif"
		            sSortALT = asDescriptors(108) 'Descriptor: Sort descending
		            sSortURL = "subscriptions.asp?rSortOrder=DESC&rOrderBy=SCHEDULE"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        Else
		            sSortIcon = "images/sort_desc.gif"
		            sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		            sSortURL = "subscriptions.asp?rSortOrder=ASC&rOrderBy=SCHEDULE"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        End If
		    Else
		        sSortIcon = "images/sort_row.gif"
		        sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		        sSortURL = "subscriptions.asp?rSortOrder=ASC&rOrderBy=SCHEDULE"
		        If sServiceID <> "" Then
		            sSortURL = sSortURL & "&serviceID=" & sServiceID
		        End If
		    End If
			Response.Write "<FONT color=#000000 face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(351) & "</B></FONT> <A HREF=""" & sSortURL & """><IMG SRC=""" & sSortIcon & """ WIDTH=""17"" HEIGHT=""8"" BORDER=""0"" ALT=""" & sSortALT & """ /></A>" 'Descriptor: Schedule
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"">"
		    If sOrderBy = "TIME" Then
		        If sSortOrder = "ASC" Then
		            sSortIcon = "images/sort_asc.gif"
		            sSortALT = asDescriptors(108) 'Descriptor: Sort descending
		            sSortURL = "subscriptions.asp?rSortOrder=DESC&rOrderBy=TIME"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        Else
		            sSortIcon = "images/sort_desc.gif"
		            sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		            sSortURL = "subscriptions.asp?rSortOrder=ASC&rOrderBy=TIME"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        End If
		    Else
		        sSortIcon = "images/sort_row.gif"
		        sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		        sSortURL = "subscriptions.asp?rSortOrder=ASC&rOrderBy=TIME"
		        If sServiceID <> "" Then
		            sSortURL = sSortURL & "&serviceID=" & sServiceID
		        End If
		    End If
		    Response.Write "<FONT color=#000000 face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(34) & "</B></FONT> <A HREF=""" & sSortURL & """><IMG SRC=""" & sSortIcon & """ WIDTH=""17"" HEIGHT=""8"" BORDER=""0"" ALT=""" & sSortALT & """ /></A>" 'Descriptor: Modified
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
        Response.Write "<TD NOWRAP=""1"">&nbsp;&nbsp;</TD>"
        Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"" ALIGN=""CENTER"">"
		    If oSubs.length > 0 Then
	    		Response.Write "&nbsp;<input type=""SUBMIT"" VALUE=""" & asDescriptors(249) & """ name=""delSummary"" class=""buttonClass"">&nbsp;" 'Descriptor: Delete
	    	Else
	    		Response.Write "&nbsp;<input type=""BUTTON"" VALUE=""" & asDescriptors(249) & """ name=""delSummary"" class=""buttonClass"">&nbsp;" 'Descriptor: Delete
	    	End If
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD bgColor=#ffffff></TD>"
	    Response.Write "</TR>"

	    If oSubs.length > 0 Then
	        Call GetXMLDOM(aConnectionInfo, oAvailSubsDOM, sErrDescription)
	        oAvailSubsDOM.async = false
	        Call oAvailSubsDOM.loadXML(sGetAvailableSubscriptionsXML)

	    	For Each oCurrentSub in oSubs

	    		If StrComp(CStr(oCurrentSub.getAttribute("act")), "1", vbBinaryCompare) = 0 Then
	    		    sFontColor = ACTIVE_COLOR
	    		Else
	    		    sFontColor = INACTIVE_COLOR
	    		End If
	    		sService = oCurrentSub.getAttribute("svn")
				sSchedule = oCurrentSub.getAttribute("scn")

	    		Response.Write "<TR bgColor=#ffffff>"
	    			Response.Write "<TD colSpan=17 height=1>"
	    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	    			Response.Write "</TD>"
	    			Response.Write "<TD bgColor=#ffffff></TD>"
	    		Response.Write "</TR>"
	    		Response.Write "<TR bgColor=#ffffff>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD>"
	    		If StrComp(CStr(oCurrentSub.getAttribute("act")), "1", vbBinaryCompare) = 0 Then
	    		    Response.Write "<IMG SRC=""images/active.gif"" HEIGHT=""15"" WIDTH=""15"" BORDER=""0"" ALT=""" & asDescriptors(472) & """ />" 'Descriptor: Active
	    		Else
	    		    Response.Write "<IMG SRC=""images/inactive.gif"" HEIGHT=""15"" WIDTH=""15"" BORDER=""0"" ALT=""" & asDescriptors(528) & """ />" 'Descriptor: Inactive
	    		End If
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2>"
	    		If Not (oAvailSubsDOM.selectSingleNode("/mi/subs/sub[@id = '" & oCurrentSub.getAttribute("guid") & "']") Is Nothing) Then
	    		    sBoldStart = "<B>"
	    		    sBoldEnd = "</B>"
	    		    Response.Write "<A HREF=""reports.asp?subsId=" & oCurrentSub.getAttribute("guid") & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """><b>" & Server.HTMLEncode(oCurrentSub.getAttribute("svn")) & "</b></font></A>"
	    		Else
	    		    sBoldStart = ""
	    		    sBoldEnd = ""
	    		    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & sService & "</font>"
	    		End If
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2>"
	    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & sBoldStart & sSchedule & sBoldEnd & "</font>"
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2>"
	    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & sBoldStart & DisplayDateAndTime(oCurrentSub.getAttribute("mdt"), "") & sBoldEnd & "</font>"
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2 NOWRAP=""1"">"
	    		Response.Write "<A HREF=""subscribe.asp?serviceID=" & oCurrentSub.getAttribute("svid") & "&eSubID=" & oCurrentSub.getAttribute("id") & "&eSGUID=" & oCurrentSub.getAttribute("guid") & "&eAID=" & oCurrentSub.getAttribute("adid") & "&eSSID=" & oCurrentSub.getAttribute("sbstid") & "&serviceName=" & Server.URLEncode(sService) & "&enf=" & oCurrentSub.getAttribute("act") & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#000000"">" & asDescriptors(353) & "</font></A>" 'Descriptor: Edit
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD ALIGN=""CENTER""><INPUT TYPE=""CHECKBOX"" NAME=""delSubsGUID"" VALUE=""" & oCurrentSub.getAttribute("sbstid") & ";" & oCurrentSub.getAttribute("guid") & ";" & oCurrentSub.getAttribute("svid") & ";" & Server.HTMLEncode(sService) & " (" & Server.HTMLEncode(sSchedule) & ")" & """ /></TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "</TR>"
	    		Response.Write "<TR bgColor=#ffffff>"
	    			Response.Write "<TD colSpan=17 height=1>"
	    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	    			Response.Write "</TD>"
	    			Response.Write "<TD bgColor=#ffffff></TD>"
	    		Response.Write "</TR>"
	    		Response.Write "<TR bgColor=#cccccc>"
	    			Response.Write "<TD colSpan=17 height=1>"
	    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	    			Response.Write "</TD>"
	    			Response.Write "<TD bgColor=#ffffff></TD>"
	    		Response.Write "</TR>"
	    	Next
	    	Set oSchedules = Nothing
	    	Set oCurrentSchedule = Nothing
	    Else
	        'Lines for if there are no subscriptions
		    Response.Write "<TR bgColor=#ffffff>"
		    	Response.Write "<TD colSpan=17 height=5>"
		    		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
		    	Response.Write "</TD>"
		    	Response.Write "<TD bgColor=#ffffff></TD>"
		    Response.Write "</TR>"
		    Response.Write "<TR bgColor=#ffffff>"
		    	Response.Write "<TD colSpan=17 height=1>"
		    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=" & aFontInfo(N_SMALL_FONT) & "><b>&nbsp;&nbsp;" & asDescriptors(532) & "</b></font>" 'Descriptor: You do not have any reports.
		    	Response.Write "</TD>"
		    	Response.Write "<TD bgColor=#ffffff></TD>"
		    Response.Write "</TR>"
		    Response.Write "<TR bgColor=#ffffff>"
		    	Response.Write "<TD colSpan=17 height=5>"
		    		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
		    	Response.Write "</TD>"
		    	Response.Write "<TD bgColor=#ffffff></TD>"
		    Response.Write "</TR>"
		    Response.Write "<TR bgColor=#cccccc>"
		    	Response.Write "<TD colSpan=17 height=1>"
		    		Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
		    	Response.Write "</TD>"
		    	Response.Write "<TD bgColor=#ffffff></TD>"
		    Response.Write "</TR>"
	    End If

        Response.Write "</TBODY>"
        Response.Write "</form>"
        Response.Write "</TABLE>"
    End If

    Set oSubsDOM = Nothing
	Set oSubs = Nothing
	Set oCurrentSub = Nothing
	Set oAvailSubsDOM = Nothing

    RenderList_Reports = lErrNumber
    Err.Clear
End Function

Function RenderLargeIcons_Reports(sServiceID, sFolderID, sSubscriptionsXML, sGetAvailableSubscriptionsXML)
'********************************************************
'*Purpose: Renders list of subscriptions to Portal Address
'*Inputs: sSubscriptionsXML
'*Outputs:
'*TO DO: Add Error messages
'********************************************************
    On Error Resume Next
    Const ACTIVE_COLOR = "#000000"
    Const INACTIVE_COLOR = "#aaaaaa"
    Const S_TYPE_SERVICE = "19"
    Dim oSubsDOM
	Dim oSubs
	Dim oServices
	Dim oCurrentService
	Dim oCurrentSub
	Dim sSchedule
	Dim sService
	Dim oSchedules
	Dim oCurrentSchedule
	Dim sFontColor
	Dim lErrNumber
	Dim oAvailSubsDOM
	lErrNumber = NO_ERR

	Set oSubsDOM = Server.CreateObject("Microsoft.XMLDOM")
	oSubsDOM.async = False
	If oSubsDOM.loadXML(sSubscriptionsXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "RenderLargeIcons_Reports", "", "Error loading subscriptions XML", LogLevelError)
		'add error message
	End If

    If lErrNumber = NO_ERR Then
        If sServiceID <> "" Then
            Set oServices = oSubsDOM.selectNodes("//oi[@id = '" & sServiceID & "' and @tp = '" & S_TYPE_SERVICE & "']")
        Else
            Set oServices = oSubsDOM.selectNodes("//oi[@tp = '" & S_TYPE_SERVICE & "']")
        End If

        If oServices.length > 0 Then
            For Each oCurrentService In oServices
                sService = oCurrentService.getAttribute("n")
	            Set oSubs = oSubsDOM.selectNodes("//sub[@adid = '" & CStr(GetPortalAddress()) & "' and dst/@svid = '" & oCurrentService.getAttribute("id") & "']")

                If oSubs.length > 0 Then
                    Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
                    Response.Write "<TR>"
                    Response.Write "<TD VALIGN=TOP WIDTH=""1%""><IMG SRC=""images/report_big.gif"" WIDTH=""60"" HEIGHT=""57"" BORDER=""0"" ALT="""" /></TD>"
                    Response.Write "<TD VALIGN=TOP WIDTH=""98%"">"
                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & sService & "</b></font><BR />"
                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oCurrentService.getAttribute("des") & "<BR />" & asDescriptors(350) & "</font>" 'Descriptor: You are subscribed to the following schedules for this report:

                        Response.Write "<BR /><TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" WIDTH=""100%"">"
                        Response.Write "<form METHOD=""POST"" ACTION=""confirm_delete.asp"">"
	                	Response.Write "<input type=""HIDDEN"" NAME=""formPage"" VALUE=""modify_subscription.asp"" />"
	                	Response.Write "<input type=""HIDDEN"" NAME=""deleteType"" VALUE=""2"" />"
	                	Response.Write "<input type=""HIDDEN"" NAME=""cancelButton"" VALUE=""subsCancel"" />"
                    	Response.Write "<input type=""HIDDEN"" NAME=""action"" VALUE=""delete"" />"
                    	Response.Write "<input type=""HIDDEN"" NAME=""serviceID"" VALUE=""" & sServiceID & """ />"
                    	Response.Write "<input type=""HIDDEN"" NAME=""folderID"" VALUE=""" & sFolderID & """ />"
                        Response.Write "<TBODY>"
                        Response.Write "<TR bgColor=#cccc99>"
	                	Response.Write "<TD>&nbsp;</TD>"
                        Response.Write "<TD><IMG alt="""" border=0 height=10 src=""Images/active_status.gif"" width=19></TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD NOWRAP=""1"" WIDTH=""100%"">"
	                		Response.Write "<FONT color=#ffffff face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><FONT color=#000000><B>" & asDescriptors(351) & "</B></FONT></FONT>" 'Descriptor: Schedule
	                	Response.Write "</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD NOWRAP=""1"" WIDTH=""1%"">"
	                	    Response.Write "<FONT color=#000000 face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(34) & "</B></FONT>" 'Descriptor: Modified
	                	Response.Write "</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
                        Response.Write "<TD NOWRAP=""1"">&nbsp;&nbsp;</TD>"
                        Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD NOWRAP=""1"" ALIGN=""CENTER"">"
	                	    If oSubs.length > 0 Then
	                    		Response.Write "&nbsp;<input type=""SUBMIT"" VALUE=""" & asDescriptors(249) & """ name=""delSummary"" class=""buttonClass"">&nbsp;" 'Descriptor: Delete
	                    	Else
	                    		Response.Write "&nbsp;<input type=""BUTTON"" VALUE=""" & asDescriptors(249) & """ name=""delSummary"" class=""buttonClass"">&nbsp;" 'Descriptor: Delete
	                    	End If
	                	Response.Write "</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD bgColor=#ffffff></TD>"
	                    Response.Write "</TR>"

	                    If oSubs.length > 0 Then
	                        Call GetXMLDOM(aConnectionInfo, oAvailSubsDOM, sErrDescription)
	                        oAvailSubsDOM.async = false
	                        Call oAvailSubsDOM.loadXML(sGetAvailableSubscriptionsXML)

	                    	For Each oCurrentSub in oSubs
	                    		sSchedule = ""
	                    		Set oSchedules = oCurrentSub.selectNodes("dst")
	                    		If oSchedules.length > 0 Then
	                    			For Each oCurrentSchedule In oSchedules
	                    				sSchedule = sSchedule & oSubsDOM.selectSingleNode("/mi/in/oi[@id = '" & oCurrentSchedule.getAttribute("scid") & "']").getAttribute("n") & ", "
	                    			Next
	                    			sSchedule = Left(sSchedule, Len(sSchedule)-2)
	                    		Else
	                    			'Add error handling
	                    		End If

	                    		If oCurrentSub.getAttribute("act") = "1" Then
	                    		    sFontColor = ACTIVE_COLOR
	                    		Else
	                    		    sFontColor = INACTIVE_COLOR
	                    		End If

	                    		Response.Write "<TR bgColor=#ffffff>"
	                    			Response.Write "<TD colSpan=14 height=1>"
	                    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	                    			Response.Write "</TD>"
	                    			Response.Write "<TD bgColor=#ffffff></TD>"
	                    		Response.Write "</TR>"
	                    		Response.Write "<TR bgColor=#ffffff>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD>"
	                    		If oCurrentSub.getAttribute("act") = "1" Then
	                    		    Response.Write "<IMG SRC=""images/active.gif"" HEIGHT=""15"" WIDTH=""15"" BORDER=""0"" ALT=""" & asDescriptors(472) & """ />" 'Descriptor: Active
	                    		Else
	                    		    Response.Write "<IMG SRC=""images/inactive.gif"" HEIGHT=""15"" WIDTH=""15"" BORDER=""0"" ALT=""" & asDescriptors(528) & """ />" 'Descriptor: Inactive
	                    		End If
	                    		Response.Write "</TD>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD COLSPAN=2>"
	                    		If Not (oAvailSubsDOM.selectSingleNode("/mi/subs/sub[@id = '" & oCurrentSub.getAttribute("guid") & "']") Is Nothing) Then
	                    		    Response.Write "<A HREF=""reports.asp?subsId=" & oCurrentSub.getAttribute("guid") & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """><b>" & sSchedule & "</b></font></A>"
	                    		Else
	                    		    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & sSchedule & "</font>"
	                    		End If
	                    		Response.Write "</TD>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD COLSPAN=2>"
	                    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """><b>" & DisplayDateAndTime(oCurrentSub.getAttribute("mdt"), "") & "</b></font>"
	                    		Response.Write "</TD>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD COLSPAN=2 NOWRAP=""1"">"
	                    		Response.Write "<A HREF=""subscribe.asp?serviceID=" & oCurrentSub.selectSingleNode("dst").getAttribute("svid") & "&eSubID=" & oCurrentSub.getAttribute("id") & "&eSGUID=" & oCurrentSub.getAttribute("guid") & "&eAID=" & oCurrentSub.getAttribute("adid") & "&eSSID=" & oCurrentSub.getAttribute("sbstid") & "&serviceName=" & Server.URLEncode(sService) & "&enf=" & oCurrentSub.getAttribute("act") & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#000000"">" & asDescriptors(353) & "</font></A>" 'Descriptor: Edit
	                    		Response.Write "</TD>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD ALIGN=""CENTER""><INPUT TYPE=""CHECKBOX"" NAME=""delSubsGUID"" VALUE=""" & oCurrentSub.getAttribute("sbstid") & ";" & oCurrentSub.getAttribute("guid") & ";" & oCurrentService.getAttribute("id") & ";" & Server.HTMLEncode(sService) & " (" & Server.HTMLEncode(sSchedule) & ")" & """ /></TD>"
	    						Response.Write "<TD></TD>"
	                    		Response.Write "</TR>"
	                    		Response.Write "<TR bgColor=#ffffff>"
	                    			Response.Write "<TD colSpan=14 height=1>"
	                    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	                    			Response.Write "</TD>"
	                    			Response.Write "<TD bgColor=#ffffff></TD>"
	                    		Response.Write "</TR>"
	                    	    Response.Write "<TR bgColor=#cccccc>"
	                    	    	Response.Write "<TD colSpan=14 height=1>"
	                    	    		Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	                    	    	Response.Write "</TD>"
	                    	    	Response.Write "<TD bgColor=#ffffff></TD>"
	                    	    Response.Write "</TR>"
	                    	Next
	                    	Set oSchedules = Nothing
	                    	Set oCurrentSchedule = Nothing
	                    End If

                        Response.Write "</TBODY>"
                        Response.Write "</form>"
                        Response.Write "</TABLE>"

                    Response.Write "</TD>"
                    Response.Write "</TR>"
                    Response.Write "<TR><TD></TD><TD><A HREF=""subscribe.asp?serviceID=" & oCurrentService.getAttribute("id") & "&serviceName=" & Server.URLEncode(oCurrentService.getAttribute("n")) & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(349) & "</font></A></TD></TR>" 'Descriptor: Add subscription
                    Response.Write "</TABLE><BR />"
                End If
            Next
            If oSubsDOM.selectNodes("//sub[@adid = '" & CStr(GetPortalAddress()) & "']").length > 0 Then
                'Do nothing
            Else
	            'Lines for if there are no subscriptions
	            Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"
	            Response.Write "<TR bgColor=#ffffff>"
	            	Response.Write "<TD height=5>"
	            		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
	            	Response.Write "</TD>"
	            	Response.Write "<TD bgColor=#ffffff></TD>"
	            Response.Write "</TR>"
	            Response.Write "<TR bgColor=#ffffff>"
	            	Response.Write "<TD height=1>"
	            		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=" & aFontInfo(N_SMALL_FONT) & "><b>&nbsp;&nbsp;" & asDescriptors(532) & "</b></font>" 'Descriptor: You do not have any reports.
	            	Response.Write "</TD>"
	            	Response.Write "<TD bgColor=#ffffff></TD>"
	            Response.Write "</TR>"
	            Response.Write "<TR bgColor=#ffffff>"
	            	Response.Write "<TD height=5>"
	            		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
	            	Response.Write "</TD>"
	            	Response.Write "<TD bgColor=#ffffff></TD>"
	            Response.Write "</TR>"
	            Response.Write "</TABLE>"
            End If
        Else
            'If oSubsDOM.selectNodes("//sub[@adid = '" & CStr(GetPortalAddress()) & "']").length > 0 Then
                'Do nothing
            'Else
	            'Lines for if there are no subscriptions
	            Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"
	            Response.Write "<TR bgColor=#ffffff>"
	            	Response.Write "<TD height=5>"
	            		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
	            	Response.Write "</TD>"
	            	Response.Write "<TD bgColor=#ffffff></TD>"
	            Response.Write "</TR>"
	            Response.Write "<TR bgColor=#ffffff>"
	            	Response.Write "<TD height=1>"
	            		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=" & aFontInfo(N_SMALL_FONT) & "><b>&nbsp;&nbsp;" & asDescriptors(532) & "</b></font>" 'Descriptor: You do not have any reports.
	            	Response.Write "</TD>"
	            	Response.Write "<TD bgColor=#ffffff></TD>"
	            Response.Write "</TR>"
	            Response.Write "<TR bgColor=#ffffff>"
	            	Response.Write "<TD height=5>"
	            		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
	            	Response.Write "</TD>"
	            	Response.Write "<TD bgColor=#ffffff></TD>"
	            Response.Write "</TR>"
	            Response.Write "</TABLE>"
            'End If
        End If
    End If

    Set oSubsDOM = Nothing
	Set oSubs = Nothing
	Set oCurrentSub = Nothing
	Set oServices = Nothing
	Set oCurrentService = Nothing
	Set oAvailSubsDOM = Nothing

    RenderLargeIcons_Reports = lErrNumber
    Err.Clear
End Function

Function RenderList_Deliveries(sSortOrder, sOrderBy, sServiceID, sFolderID, sSubscriptionsXML)
'********************************************************
'*Purpose:
'*Inputs: sSubscriptionsXML
'*Outputs:
'*TO DO: Add Error messages
'********************************************************
    On Error Resume Next
    Const ACTIVE_COLOR = "#000000"
    Const INACTIVE_COLOR = "#aaaaaa"
    Dim oSubsDOM
	Dim oSubs
	Dim oCurrentSub
	Dim sSchedule
	Dim sAddress
	Dim oSchedules
	Dim oCurrentSchedule
	Dim sFontColor
	Dim lErrNumber
	Dim sSortIcon
	Dim sSortALT
	Dim sSortURL
	lErrNumber = NO_ERR

	Set oSubsDOM = Server.CreateObject("Microsoft.XMLDOM")
	oSubsDOM.async = False
	If oSubsDOM.loadXML(sSubscriptionsXML) = False Then
	'Temporary hard-coded to sample XML
	'If oSubsDOM.load(Server.MapPath("subscriptions.xml")) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "RenderList_Deliveries", "", "Error loading subscriptions XML", LogLevelError)
		'add error message
	End If

    If lErrNumber = NO_ERR Then
        lErrNumber = ConvertXMLForSorting_Deliveries(aConnectionInfo, oSubsDOM)
    End If

    'Added for sorting
    If lErrNumber = NO_ERR Then
    	Dim asDictionary(1,1)
    	asDictionary(0,0) = "OrderBy"
        asDictionary(1,0) = sOrderBy
        asDictionary(0,1) = "SortOrder"
        asDictionary(1,1) = sSortOrder
        lErrNumber = AddInputsToXML(aConnectionInfo, oSubsDOM, asDictionary)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "RenderList_Deliveries", "", "Error in call to AddInputsToXML function", LogLevelTrace)
        Else
            lErr = SortDeliveries(aConnectionInfo, "SortDeliveries_Subs.xsl", oSubsDOM)
            'oSubsDOM.save(server.MapPath("deltest.xml"))
        End If
    End If

    If lErrNumber = NO_ERR Then
        If sServiceID <> "" Then
            Set oSubs = oSubsDOM.selectNodes("//sub[@svid = '" & sServiceID & "' and @adid != '" & CStr(GetPortalAddress()) & "']")
        Else
            Set oSubs = oSubsDOM.selectNodes("//sub[@adid != '" & CStr(GetPortalAddress()) & "']")
        End If

        Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" WIDTH=""100%"">"
        Response.Write "<FORM METHOD=""POST"" ACTION=""confirm_delete.asp"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""formPage"" VALUE=""modify_subscription.asp"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""deleteType"" VALUE=""2"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""cancelButton"" VALUE=""subsCancel"" />"
    	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""action"" VALUE=""delete"" />"
    	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""serviceID"" VALUE=""" & sServiceID & """ />"
    	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""folderID"" VALUE=""" & sFolderID & """ />"
        Response.Write "<TBODY>"
        Response.Write "<TR bgColor=#cccc99>"
		Response.Write "<TD>&nbsp;</TD>"
        Response.Write "<TD><IMG alt="""" border=0 height=10 src=""Images/active_status.gif"" width=19></TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"">"
		    If sOrderBy = "SERVICE" Then
		        If sSortOrder = "ASC" Then
		            sSortIcon = "images/sort_asc.gif"
		            sSortALT = asDescriptors(108) 'Descriptor: Sort descending
		            sSortURL = "subscriptions.asp?dSortOrder=DESC&dOrderBy=SERVICE"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        Else
		            sSortIcon = "images/sort_desc.gif"
		            sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		            sSortURL = "subscriptions.asp?dSortOrder=ASC&dOrderBy=SERVICE"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        End If
		    Else
		        sSortIcon = "images/sort_row.gif"
		        sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		        sSortURL = "subscriptions.asp?dSortOrder=ASC&dOrderBy=SERVICE"
		        If sServiceID <> "" Then
		            sSortURL = sSortURL & "&serviceID=" & sServiceID
		        End If
		    End If
			Response.Write "<FONT face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><FONT color=#000000><B>" & asDescriptors(366) & "</B></FONT></FONT> <A HREF=""" & sSortURL & """><IMG SRC=""" & sSortIcon & """ WIDTH=""17"" HEIGHT=""8"" BORDER=""0"" ALT=""" & sSortALT & """ /></A>" 'Descriptor: Service
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"">"
            If sOrderBy = "SCHEDULE" Then
		        If sSortOrder = "ASC" Then
		            sSortIcon = "images/sort_asc.gif"
		            sSortALT = asDescriptors(108) 'Descriptor: Sort descending
		            sSortURL = "subscriptions.asp?dSortOrder=DESC&dOrderBy=SCHEDULE"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        Else
		            sSortIcon = "images/sort_desc.gif"
		            sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		            sSortURL = "subscriptions.asp?dSortOrder=ASC&dOrderBy=SCHEDULE"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        End If
		    Else
		        sSortIcon = "images/sort_row.gif"
		        sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		        sSortURL = "subscriptions.asp?dSortOrder=ASC&dOrderBy=SCHEDULE"
		        If sServiceID <> "" Then
		            sSortURL = sSortURL & "&serviceID=" & sServiceID
		        End If
		    End If
			Response.Write "<FONT color=#000000 face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(351) & "</B></FONT> <A HREF=""" & sSortURL & """><IMG SRC=""" & sSortIcon & """ WIDTH=""17"" HEIGHT=""8"" BORDER=""0"" ALT=""" & sSortALT & """ /></A>" 'Descriptor: Schedule
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"">"
            If sOrderBy = "ADDRESS" Then
		        If sSortOrder = "ASC" Then
		            sSortIcon = "images/sort_asc.gif"
		            sSortALT = asDescriptors(108) 'Descriptor: Sort descending
		            sSortURL = "subscriptions.asp?dSortOrder=DESC&dOrderBy=ADDRESS"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        Else
		            sSortIcon = "images/sort_desc.gif"
		            sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		            sSortURL = "subscriptions.asp?dSortOrder=ASC&dOrderBy=ADDRESS"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        End If
		    Else
		        sSortIcon = "images/sort_row.gif"
		        sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		        sSortURL = "subscriptions.asp?dSortOrder=ASC&dOrderBy=ADDRESS"
		        If sServiceID <> "" Then
		            sSortURL = sSortURL & "&serviceID=" & sServiceID
		        End If
		    End If
			Response.Write "<FONT color=#000000 face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(367) & "</B></FONT> <A HREF=""" & sSortURL & """><IMG SRC=""" & sSortIcon & """ WIDTH=""17"" HEIGHT=""8"" BORDER=""0"" ALT=""" & sSortALT & """ /></A>" 'Descriptor: Address
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"">"
            If sOrderBy = "TIME" Then
		        If sSortOrder = "ASC" Then
		            sSortIcon = "images/sort_asc.gif"
		            sSortALT = asDescriptors(108) 'Descriptor: Sort descending
		            sSortURL = "subscriptions.asp?dSortOrder=DESC&dOrderBy=TIME"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        Else
		            sSortIcon = "images/sort_desc.gif"
		            sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		            sSortURL = "subscriptions.asp?dSortOrder=ASC&dOrderBy=TIME"
		            If sServiceID <> "" Then
		                sSortURL = sSortURL & "&serviceID=" & sServiceID
		            End If
		        End If
		    Else
		        sSortIcon = "images/sort_row.gif"
		        sSortALT = asDescriptors(107) 'Descriptor: Sort ascending
		        sSortURL = "subscriptions.asp?dSortOrder=ASC&dOrderBy=TIME"
		        If sServiceID <> "" Then
		            sSortURL = sSortURL & "&serviceID=" & sServiceID
		        End If
		    End If
		    Response.Write "<FONT color=#000000 face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(34) & "</B></FONT> <A HREF=""" & sSortURL & """><IMG SRC=""" & sSortIcon & """ WIDTH=""17"" HEIGHT=""8"" BORDER=""0"" ALT=""" & sSortALT & """ /></A>" 'Descriptor: Modified
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"">&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD NOWRAP=""1"" ALIGN=""CENTER"">"
		    If oSubs.length > 0 Then
	    		Response.Write "&nbsp;<input type=""SUBMIT"" VALUE=""" & asDescriptors(249) & """ name=""delSummary"" class=""buttonClass"">&nbsp;" 'Descriptor: Delete
	    	Else
	    		Response.Write "&nbsp;<input type=""BUTTON"" VALUE=""" & asDescriptors(249) & """ name=""delSummary"" class=""buttonClass"">&nbsp;" 'Descriptor: Delete
	    	End If
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;</TD>"
		Response.Write "<TD bgColor=#ffffff></TD>"
	    Response.Write "</TR>"

	    If oSubs.length > 0 Then
	    	For Each oCurrentSub in oSubs
                If oCurrentSub.getAttribute("act") = "1" Then
                    sFontColor = ACTIVE_COLOR
                Else
                    sFontColor = INACTIVE_COLOR
                End If
				sSchedule = oCurrentSub.getAttribute("scn")
				sAddress = oCurrentSub.getAttribute("adn")

	    		Response.Write "<TR bgColor=#ffffff>"
	    			Response.Write "<TD colSpan=20 height=1>"
	    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	    			Response.Write "</TD>"
	    			Response.Write "<TD bgColor=#ffffff></TD>"
	    		Response.Write "</TR>"
	    		Response.Write "<TR bgColor=#ffffff>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD>"
	    		If oCurrentSub.getAttribute("act") = "1" Then
	    		    Response.Write "<IMG SRC=""images/active.gif"" HEIGHT=""15"" WIDTH=""15"" BORDER=""0"" ALT=""" & asDescriptors(472) & """ />" 'Descriptor: Active
	    		Else
	    		    Response.Write "<IMG SRC=""images/inactive.gif"" HEIGHT=""15"" WIDTH=""15"" BORDER=""0"" ALT=""" & asDescriptors(528) & """ />" 'Descriptor: Inactive
	    		End If
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2>"
	    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & oCurrentSub.getAttribute("svn") & "</font>"
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2>"
	    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & sSchedule & "</font>"
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2>"
	    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & sAddress & "</font>"
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2>"
	    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & DisplayDateAndTime(oCurrentSub.getAttribute("mdt"), "") & "</font>"
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2 NOWRAP=""1"">"
	    		Response.Write "<A HREF=""subscribe.asp?serviceID=" & oCurrentSub.getAttribute("svid") & "&eSubID=" & oCurrentSub.getAttribute("id") & "&eSGUID=" & oCurrentSub.getAttribute("guid") & "&eAID=" & oCurrentSub.getAttribute("adid") & "&eSSID=" & oCurrentSub.getAttribute("sbstid") & "&serviceName=" & Server.URLEncode(oCurrentSub.getAttribute("svn")) & "&enf=" & oCurrentSub.getAttribute("act") & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#000000"">" & asDescriptors(353) & "</font></A>" 'Descriptor: Edit
	    		Response.Write "</TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "<TD COLSPAN=2 ALIGN=""CENTER""><INPUT TYPE=""CHECKBOX"" NAME=""delSubsGUID"" VALUE=""" & oCurrentSub.getAttribute("sbstid") & ";" & oCurrentSub.getAttribute("guid") & ";" & oCurrentSub.getAttribute("svid") & ";" & Server.HTMLEncode(oCurrentSub.getAttribute("svn")) & " (" & Server.HTMLEncode(sSchedule) & ", " & Server.HTMLEncode(sAddress) & ")" &  """ /></TD>"
	    		Response.Write "<TD></TD>"
	    		Response.Write "</TR>"
	    		Response.Write "<TR bgColor=#ffffff>"
	    			Response.Write "<TD colSpan=20 height=1>"
	    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	    			Response.Write "</TD>"
	    			Response.Write "<TD bgColor=#ffffff></TD>"
	    		Response.Write "</TR>"
	    		Response.Write "<TR bgColor=#cccccc>"
	    			Response.Write "<TD colSpan=20 height=1>"
	    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	    			Response.Write "</TD>"
	    			Response.Write "<TD bgColor=#ffffff></TD>"
	    		Response.Write "</TR>"
	    	Next
	    	Set oSchedules = Nothing
	    	Set oCurrentSchedule = Nothing
	    Else
	        'Lines for if there are no subscriptions
		    Response.Write "<TR bgColor=#ffffff>"
		    	Response.Write "<TD colSpan=20 height=5>"
		    		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
		    	Response.Write "</TD>"
		    	Response.Write "<TD bgColor=#ffffff></TD>"
		    Response.Write "</TR>"
		    Response.Write "<TR bgColor=#ffffff>"
		    	Response.Write "<TD colSpan=20 height=1>"
		    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=" & aFontInfo(N_SMALL_FONT) & "><b>&nbsp;&nbsp;" & asDescriptors(358) & "</b></font>" 'Descriptor: You do not have any subscriptions.
		    	Response.Write "</TD>"
		    	Response.Write "<TD bgColor=#ffffff></TD>"
		    Response.Write "</TR>"
		    Response.Write "<TR bgColor=#ffffff>"
		    	Response.Write "<TD colSpan=20 height=5>"
		    		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
		    	Response.Write "</TD>"
		    	Response.Write "<TD bgColor=#ffffff></TD>"
		    Response.Write "</TR>"
		    Response.Write "<TR bgColor=#cccccc>"
		    	Response.Write "<TD colSpan=20 height=1>"
		    		Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
		    	Response.Write "</TD>"
		    	Response.Write "<TD bgColor=#ffffff></TD>"
		    Response.Write "</TR>"
	    End If

        Response.Write "</TBODY>"
        Response.Write "</form>"
        Response.Write "</TABLE>"
    End If

    Set oSubsDOM = Nothing
	Set oSubs = Nothing
	Set oCurrentSub = Nothing

    RenderList_Deliveries = lErrNumber
    Err.Clear
End Function

Function RenderLargeIcons_Deliveries(sServiceID, sFolderID, sSubscriptionsXML)
'********************************************************
'*Purpose: Renders list of subscriptions to Portal Address
'*Inputs: sSubscriptionsXML
'*Outputs:
'*TO DO: Add Error messages
'********************************************************
    On Error Resume Next
    Const ACTIVE_COLOR = "#000000"
    Const INACTIVE_COLOR = "#aaaaaa"
    Const S_TYPE_SERVICE = "19"
    Dim oSubsDOM
	Dim oSubs
	Dim oServices
	Dim oCurrentService
	Dim oCurrentSub
	Dim sSchedule
	Dim sAddress
	Dim oSchedules
	Dim oCurrentSchedule
	Dim sFontColor
	Dim lErrNumber
	lErrNumber = NO_ERR

	Set oSubsDOM = Server.CreateObject("Microsoft.XMLDOM")
	oSubsDOM.async = False
	If oSubsDOM.loadXML(sSubscriptionsXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "RenderLargeIcons_Deliveries", "", "Error loading subscriptions XML", LogLevelError)
		'add error message
	End If

    If lErrNumber = NO_ERR Then
        If sServiceID <> "" Then
            Set oServices = oSubsDOM.selectNodes("//oi[@id = '" & sServiceID & "' and @tp = '" & S_TYPE_SERVICE & "']")
        Else
            Set oServices = oSubsDOM.selectNodes("//oi[@tp = '" & S_TYPE_SERVICE & "']")
        End If

        If oServices.length > 0 Then
            For Each oCurrentService In oServices
                Set oSubs = oSubsDOM.selectNodes("//sub[@adid != '" & CStr(GetPortalAddress()) & "' and dst/@svid = '" & oCurrentService.getAttribute("id") & "']")

                If oSubs.length > 0 Then
                    Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"
                    Response.Write "<TR>"
                    Response.Write "<TD VALIGN=TOP WIDTH=""1%""><IMG SRC=""images/report_big.gif"" WIDTH=""60"" HEIGHT=""57"" BORDER=""0"" ALT="""" /></TD>"
                    Response.Write "<TD VALIGN=TOP WIDTH=""98%"">"
                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & oCurrentService.getAttribute("n") & "</b></font><BR />"
                    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oCurrentService.getAttribute("des") & "<BR />" & asDescriptors(350) & "</font>" 'Descriptor: You are subscribed to the following schedules for this report:

                        Response.Write "<BR /><TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" WIDTH=""100%"">"
                        Response.Write "<form METHOD=""POST"" ACTION=""confirm_delete.asp"">"
	                	Response.Write "<input type=""HIDDEN"" NAME=""formPage"" VALUE=""modify_subscription.asp"" />"
	                	Response.Write "<input type=""HIDDEN"" NAME=""deleteType"" VALUE=""2"" />"
	                	Response.Write "<input type=""HIDDEN"" NAME=""cancelButton"" VALUE=""subsCancel"" />"
                    	Response.Write "<input type=""HIDDEN"" NAME=""action"" VALUE=""delete"" />"
                    	Response.Write "<input type=""HIDDEN"" NAME=""serviceID"" VALUE=""" & sServiceID & """ />"
                    	Response.Write "<input type=""HIDDEN"" NAME=""folderID"" VALUE=""" & sFolderID & """ />"
                        Response.Write "<TBODY>"
                        Response.Write "<TR bgColor=#cccc99>"
	                	Response.Write "<TD>&nbsp;</TD>"
                        Response.Write "<TD><IMG alt="""" border=0 height=10 src=""Images/active_status.gif"" width=19></TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD NOWRAP=""1"" WIDTH=""100%"">"
	                		Response.Write "<FONT color=#ffffff face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><FONT color=#000000><B>" & asDescriptors(351) & "</B></FONT></FONT>" 'Descriptor: Schedule
	                	Response.Write "</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
		                Response.Write "<TD NOWRAP=""1"" WIDTH=""1%"">"
		                	Response.Write "<FONT color=#000000 face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(367) & "</B></FONT>" 'Descriptor: Address
		                Response.Write "</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD NOWRAP=""1"" WIDTH=""1%"">"
	                	    Response.Write "<FONT color=#000000 face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """><B>" & asDescriptors(34) & "</B></FONT>" 'Descriptor: Modified
	                	Response.Write "</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
                        Response.Write "<TD NOWRAP=""1"">&nbsp;&nbsp;</TD>"
                        Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD NOWRAP=""1"" ALIGN=""CENTER"">"
	                	    If oSubs.length > 0 Then
	                    		Response.Write "&nbsp;<input type=""SUBMIT"" VALUE=""" & asDescriptors(249) & """ name=""delSummary"" class=""buttonClass"">&nbsp;" 'Descriptor: Delete
	                    	Else
	                    		Response.Write "&nbsp;<input type=""BUTTON"" VALUE=""" & asDescriptors(249) & """ name=""delSummary"" class=""buttonClass"">&nbsp;" 'Descriptor: Delete
	                    	End If
	                	Response.Write "</TD>"
	                	Response.Write "<TD>&nbsp;&nbsp;</TD>"
	                	Response.Write "<TD bgColor=#ffffff></TD>"
	                    Response.Write "</TR>"

	                    If oSubs.length > 0 Then
	                    	For Each oCurrentSub in oSubs
	                    		sSchedule = ""
	                    		Set oSchedules = oCurrentSub.selectNodes("dst")
	                    		If oSchedules.length > 0 Then
	                    			For Each oCurrentSchedule In oSchedules
	                    				sSchedule = sSchedule & oSubsDOM.selectSingleNode("/mi/in/oi[@id = '" & oCurrentSchedule.getAttribute("scid") & "']").getAttribute("n") & ", "
	                    			Next
	                    			sSchedule = Left(sSchedule, Len(sSchedule)-2)
	                    		Else
	                    			'Add error handling
	                    		End If

	                    		If oCurrentSub.getAttribute("act") = "1" Then
	                    		    sFontColor = ACTIVE_COLOR
	                    		Else
	                    		    sFontColor = INACTIVE_COLOR
	                    		End If

								sAddress = oSubsDOM.selectSingleNode("/mi/in/oi[@id = '" & oCurrentSub.getAttribute("adid") & "']").getAttribute("n")

	                    		Response.Write "<TR bgColor=#ffffff>"
	                    			Response.Write "<TD colSpan=17 height=1>"
	                    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	                    			Response.Write "</TD>"
	                    			Response.Write "<TD bgColor=#ffffff></TD>"
	                    		Response.Write "</TR>"
	                    		Response.Write "<TR bgColor=#ffffff>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD>"
	                    		If oCurrentSub.getAttribute("act") = "1" Then
	                    		    Response.Write "<IMG SRC=""images/active.gif"" HEIGHT=""15"" WIDTH=""15"" BORDER=""0"" ALT=""" & asDescriptors(472) & """ />" 'Descriptor: Active
	                    		Else
	                    		    Response.Write "<IMG SRC=""images/inactive.gif"" HEIGHT=""15"" WIDTH=""15"" BORDER=""0"" ALT=""" & asDescriptors(528) & """ />" 'Descriptor: Inactive
	                    		End If
	                    		Response.Write "</TD>"

	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD COLSPAN=2>"
	                    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & sSchedule & "</font>"
	                    		Response.Write "</TD>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD COLSPAN=2>"
            		    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & sAddress & "</font>"
	                    		Response.Write "</TD>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD COLSPAN=2>"
	                    		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sFontColor & """>" & DisplayDateAndTime(oCurrentSub.getAttribute("mdt"), "") & "</font>"
	                    		Response.Write "</TD>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD COLSPAN=2 NOWRAP=""1"">"
	                    		Response.Write "<A HREF=""subscribe.asp?serviceID=" & oCurrentSub.selectSingleNode("dst").getAttribute("svid") & "&eSubID=" & oCurrentSub.getAttribute("id") & "&eSGUID=" & oCurrentSub.getAttribute("guid") & "&eAID=" & oCurrentSub.getAttribute("adid") & "&eSSID=" & oCurrentSub.getAttribute("sbstid") & "&serviceName=" & Server.URLEncode(oCurrentService.getAttribute("n")) & "&enf=" & oCurrentSub.getAttribute("act") & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""#000000"">" & asDescriptors(353) & "</font></A>" 'Descriptor: Edit
	                    		Response.Write "</TD>"
	                    		Response.Write "<TD></TD>"
	                    		Response.Write "<TD COLSPAN=2 ALIGN=""CENTER""><INPUT TYPE=""CHECKBOX"" NAME=""delSubsGUID"" VALUE=""" & oCurrentSub.getAttribute("sbstid") & ";" & oCurrentSub.getAttribute("guid") & ";" & oCurrentService.getAttribute("id") & ";" & Server.HTMLEncode(oCurrentService.getAttribute("n")) & " (" & Server.HTMLEncode(sSchedule) & ", " & Server.HTMLEncode(sAddress) & ")" & """ /></TD>"
	    						Response.Write "<TD></TD>"
	                    		Response.Write "</TR>"
	                    		Response.Write "<TR bgColor=#ffffff>"
	                    			Response.Write "<TD colSpan=17 height=1>"
	                    				Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	                    			Response.Write "</TD>"
	                    			Response.Write "<TD bgColor=#ffffff></TD>"
	                    		Response.Write "</TR>"
	                    	    Response.Write "<TR bgColor=#cccccc>"
	                    	    	Response.Write "<TD colSpan=17 height=1>"
	                    	    		Response.Write "<IMG alt="""" border=0 height=1 src=""Images/1ptrans.gif"" width=1>"
	                    	    	Response.Write "</TD>"
	                    	    	Response.Write "<TD bgColor=#ffffff></TD>"
	                    	    Response.Write "</TR>"
	                    	Next
	                    	Set oSchedules = Nothing
	                    	Set oCurrentSchedule = Nothing
	                    End If

                        Response.Write "</TBODY>"
                        Response.Write "</form>"
                        Response.Write "</TABLE>"

                    Response.Write "</TD>"
                    Response.Write "</TR>"
                    Response.Write "<TR><TD></TD><TD><A HREF=""subscribe.asp?serviceID=" & oCurrentService.getAttribute("id") & "&serviceName=" & Server.URLEncode(oCurrentService.getAttribute("n")) & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(349) & "</font></A></TD></TR>" 'Descriptor: Add subscription
                    Response.Write "</TABLE><BR />"
                End If
            Next
            If oSubsDOM.selectNodes("//sub[@adid != '" & CStr(GetPortalAddress()) & "']").length > 0 Then
                'Do nothing
            Else
	            'Lines for if there are no subscriptions
	            Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"
	            Response.Write "<TR bgColor=#ffffff>"
	            	Response.Write "<TD height=5>"
	            		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
	            	Response.Write "</TD>"
	            	Response.Write "<TD bgColor=#ffffff></TD>"
	            Response.Write "</TR>"
	            Response.Write "<TR bgColor=#ffffff>"
	            	Response.Write "<TD height=1>"
	            		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=" & aFontInfo(N_SMALL_FONT) & "><b>&nbsp;&nbsp;" & asDescriptors(358) & "</b></font>" 'Descriptor: You do not have any subscriptions.
	            	Response.Write "</TD>"
	            	Response.Write "<TD bgColor=#ffffff></TD>"
	            Response.Write "</TR>"
	            Response.Write "<TR bgColor=#ffffff>"
	            	Response.Write "<TD height=5>"
	            		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
	            	Response.Write "</TD>"
	            	Response.Write "<TD bgColor=#ffffff></TD>"
	            Response.Write "</TR>"
	            Response.Write "</TABLE>"
            End If
        Else
	        'Lines for if there are no subscriptions
	        Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"
	        Response.Write "<TR bgColor=#ffffff>"
	        	Response.Write "<TD height=5>"
	        		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
	        	Response.Write "</TD>"
	        	Response.Write "<TD bgColor=#ffffff></TD>"
	        Response.Write "</TR>"
	        Response.Write "<TR bgColor=#ffffff>"
	        	Response.Write "<TD height=1>"
	        		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=" & aFontInfo(N_SMALL_FONT) & "><b>&nbsp;&nbsp;" & asDescriptors(358) & "</b></font>" 'Descriptor: You do not have any subscriptions.
	        	Response.Write "</TD>"
	        	Response.Write "<TD bgColor=#ffffff></TD>"
	        Response.Write "</TR>"
	        Response.Write "<TR bgColor=#ffffff>"
	        	Response.Write "<TD height=5>"
	        		Response.Write "<IMG alt="""" border=0 height=5 src=""Images/1ptrans.gif"" width=1>"
	        	Response.Write "</TD>"
	        	Response.Write "<TD bgColor=#ffffff></TD>"
	        Response.Write "</TR>"
	        Response.Write "</TABLE>"
        End If
    End If

    Set oSubsDOM = Nothing
	Set oSubs = Nothing
	Set oCurrentSub = Nothing
	Set oServices = Nothing
	Set oCurrentService = Nothing

    RenderLargeIcons_Deliveries = lErrNumber
    Err.Clear
End Function

Function RenderPath_Subscriptions(sServiceID, sGetFolderContentsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
Dim lErrNumber
Dim oContentsDOM
Dim oFolder
Dim iNumFolders
Dim oRootFolder
Dim i

    On Error Resume Next
    lErrNumber = NO_ERR

    iNumFolders = 0

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetFolderContentsXML, oContentsDOM)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "RenderPath_Subscriptions", "", "Error loading folderContents xml", LogLevelError)
    Else

        Response.Write "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(26) & " " 'Descriptor: You are here:

        If Len(sServiceID) > 0 Then

            Set oRootFolder = oContentsDOM.selectSingleNode("//a/fd[@id='" & APP_ROOT_FOLDER & "']").parentNode
            iNumFolders = CInt(oRootFolder.selectNodes(".//a").length)

            Response.Write "<A HREF=""services.asp""><FONT COLOR=""#000000"">" & asDescriptors(362) & "</FONT></A>" 'Descriptor: Services

            If iNumFolders > 0 Then
                Set oFolder = oRootFolder
                For i = 1 To iNumFolders
                    Set oFolder = oFolder.selectSingleNode("a")
                    Response.Write " &gt; "
                    Response.Write "<A HREF=""services.asp?folderID=" & oFolder.selectSingleNode("fd").getAttribute("id") & """><FONT COLOR=""#0000"">" & Server.HTMLEncode(oFolder.selectSingleNode("fd").getAttribute("n")) & "</FONT></A>"
                Next
            End If

            Response.Write " &gt; <B>" & asDescriptors(354) & ": " & Server.HTMLEncode(oContentsDOM.selectSingleNode("/mi/fct/oi[@id = '" & sServiceID & "']").getAttribute("n")) & "</B>" 'Descriptor: Subscriptions

        Else
            Response.Write " &gt; <B>" & asDescriptors(354) & "</B>" 'Descriptor: Subscriptions

        End If

        Response.Write "</FONT>"


    End If

    Set oContentsDOM = Nothing
    Set oFolder = Nothing

    RenderPath_Subscriptions = lErrNumber
    Err.Clear

End Function

Function SortDeliveries(aConnectionInfo, sXSLFileName, oXMLRoot)
'******************************************************************************
'Purpose: Sort Inbox XML and return it as an object
'Inputs: aConnectionInfo, sXSLFileName, oXMLRoot
'Outputs: oXMLRoot, sErrDescription, lErrNumber
'******************************************************************************
	On Error Resume Next
	Dim oNewXML
	Dim oXSLRoot
	Dim lErrNumber
	Dim sErrDescription
	lErrNumber = NO_ERR
	lErrNumber = GetXMLDOM(aConnectionInfo, oXSLRoot, sErrDescription)
	If Err.number <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, Cstr(Err.number), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "SortDeliveries", "", "Error in call to  GetXMLDOM function", LogLevelTrace)
	End If
	If lErrNumber = NO_ERR Then
		lErrNumber = GetXMLDOM(aConnectionInfo, oNewXML, sErrDescription)
		If Err.number <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, Cstr(Err.number), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "SortDeliveries", "", "Error in call to  GetXMLDOM function", LogLevelTrace)
		End If
	End If
	If lErrNumber = NO_ERR Then
		oXSLRoot.Load (Server.MapPath(sXSLFileName))
		If Err.Number <> NO_ERR Then
			lErrNumber = Err.number
			Call LogErrorXML(aConnectionInfo, Cstr(Err.number), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "SortDeliveries", "", "Error loading XML", LogLevelError)
		Else
			Call oXMLRoot.transformNodeToObject(oXSLRoot, oNewXML)
			If Err.number <> NO_ERR Then
				lErrNumber = Err.number
			    Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "SortDeliveries", "", "Error oXMLRoot.transformNodeToObject", LogLevelError)
			Else
				Set oXMLRoot = oNewXML
			End If
		End If
	End If
	Set oNewXML = Nothing
	Set oXSLRoot = Nothing
	SortDeliveries = lErrNumber
	Err.Clear
End Function

Function ConvertXMLForSorting_Deliveries(aConnectionInfo, oSubsDOM)
'******************************************************************************
'Purpose:
'Inputs:
'Outputs:
'TO DO: add error handling!
'******************************************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oNewXML
    Dim oNewElement
    Dim oNewNode
    Dim oNewAtt
    Dim sErrDescription
    Dim i
    Dim oSubs
    Dim oSchedules
    Dim oCurrentSchedule
    Dim sSchedule
    lErrNumber = NO_ERR

    Set oSubs = oSubsDOM.selectNodes("//sub")
    If oSubs.length > 0 Then
        lErrNumber = GetXMLDOM(aConnectionInfo, oNewXML, sErrDescription)
	    If Err.number <> NO_ERR Then
	    	Call LogErrorXML(aConnectionInfo, Cstr(Err.number), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", "ConvertXMLForSorting_Deliveries", "", "Error in call to  GetXMLDOM function", LogLevelTrace)
	    End If

            If lErrNumber = NO_ERR Then
                oNewXML.async = False
                oNewXML.loadXML("<subs></subs>")

                For i = 0 To (oSubs.length - 1)
                    Set oNewNode = oNewXML.documentElement.SelectSingleNode("/subs")
                    Set oNewElement = oNewXML.createElement("sub")
                    Set oNewNode = oNewNode.appendChild(oNewElement)
                    oNewNode.setAttribute "id", oSubs.item(i).getAttribute("id")
                    oNewNode.setAttribute "guid", oSubs.item(i).getAttribute("guid")
                    oNewNode.setAttribute "sbstid", oSubs.item(i).getAttribute("sbstid")
                    oNewNode.setAttribute "przd", oSubs.item(i).getAttribute("przd")
                    oNewNode.setAttribute "act", oSubs.item(i).getAttribute("act")
                    oNewNode.setAttribute "srtt", DateTimeToString(oSubs.item(i).getAttribute("mdt"))
                    oNewNode.setAttribute "mdt", oSubs.item(i).getAttribute("mdt")
                    oNewNode.setAttribute "adid", oSubs.item(i).getAttribute("adid")
                    oNewNode.setAttribute "adn", oSubsDOM.selectSingleNode("/mi/in/oi[@id = '" & oSubs.item(i).getAttribute("adid") & "']").getAttribute("n")
                    oNewNode.setAttribute "svid", oSubs.item(i).selectSingleNode("dst").getAttribute("svid")
                    oNewNode.setAttribute "svn", oSubsDOM.selectSingleNode("/mi/in/oi[@id = '" & oSubs.item(i).selectSingleNode("dst").getAttribute("svid") & "']").getAttribute("n")
                    sSchedule = ""
		                Set oSchedules = oSubs.item(i).selectNodes("dst")
		                If oSchedules.length > 0 Then
		                	For Each oCurrentSchedule In oSchedules
		                		sSchedule = sSchedule & oSubsDOM.selectSingleNode("/mi/in/oi[@id = '" & oCurrentSchedule.getAttribute("scid") & "']").getAttribute("n") & ", "
		                	Next
		                	sSchedule = Left(sSchedule, Len(sSchedule)-2)
		                Else
		                	'Add error handling
		                End If
		                Set oCurrentSchedule = Nothing
		                Set oSchedules = Nothing
                    oNewNode.setAttribute "scn", sSchedule

                    If Err.Number <> 0 Then Exit For
                Next

                Set oNewElement = Nothing
                Set oNewNode = Nothing
                Set oNewAtt = Nothing

                'oNewXML.save(Server.MapPath("deltest.xml"))
                Set oSubsDOM = oNewXML
            End If

    Else
        'Do anything?
    End If

    Set oNewXML = Nothing
    Set oSubs = Nothing

    ConvertXMLForSorting_Deliveries = lErrNumber
    Err.Clear
End Function

Function cu_GetUserSubscriptions(sGetUserSubscriptionsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Add error handling.
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetUserSubscriptions"
	Dim lErrNumber
    Dim sSessionID
    Dim sChannelID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()
	sChannelID = GetCurrentChannel()

	lErrNumber = co_GetUserSubscriptions(sSessionID, sChannelID, sGetUserSubscriptionsXML)
	If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubscriptionsCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetUserSubscriptions", LogLevelTrace)
	End If

	cu_GetUserSubscriptions = lErrNumber
	Err.Clear
End Function
%>