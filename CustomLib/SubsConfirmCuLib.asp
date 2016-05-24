<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Const PAGE_NAME = "SubsConfirmCuLib.asp"

Function ParseRequestForSubsConfirm(oRequest, sSubGUID, sStatus)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber

	lErrNumber = NO_ERR

	sSubGUID = ""
	sStatus = ""

	sSubGUID = Trim(CStr(oRequest("subGUID")))
	sStatus = Trim(CStr(oRequest("status")))

	If Err.number <> NO_ERR Then
	    lErrNumber = Err.number
	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubsConfirmCuLib.asp", "ParseRequestForSubsConfirm", "", "Error setting variables equal to Request variables", LogLevelError)
	Else
	    If Len(sSubGUID) = 0 Then
	        lErrNumber = URL_MISSING_PARAMETER
	    End If
	End If

	ParseRequestForSubsConfirm = lErrNumber
	Err.Clear
End Function

Function ReadCacheVariables_SubsConfirm(sCacheXML, sFolderID, sServiceID, sServiceName, sScheduleName, sAddressID, sAddressName, sSubSetID, sPublicationID, sPersonalized)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
	Dim oCacheDOM
	Dim oSub
	Dim oCurrQO
	Dim sPersonalization

	lErrNumber = NO_ERR

	Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
	oCacheDOM.async = False
	If oCacheDOM.loadXML(sCacheXML) = False Then
		lErrNumber = ERR_XML_LOAD_FAILED
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubsConfirmCuLib.asp", "ReadCacheVariables_SubsConfirm", "", "Error loading sCacheXML", LogLevelError)
    Else
        Set oSub = oCacheDOM.selectSingleNode("/mi/sub")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SubsConfirmCuLib.asp", "ReadCacheVariables_SubsConfirm", "", "Error retrieving sub node", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
        sFolderID = oSub.getAttribute("fid")
        sServiceID = oSub.getAttribute("svcid")
        sServiceName = oSub.getAttribute("svn")
        sScheduleName = oSub.getAttribute("scn")
        sAddressID = oSub.getAttribute("adid")
        sAddressName = oSub.getAttribute("adn")
        sSubSetID = oSub.getAttribute("sbstid")
        sPublicationID = oSub.getAttribute("pubid")

        'If (oCacheDOM.selectNodes("//oi[@tp = '" & TYPE_QUESTION & "']").length > 0) Then
        If (oCacheDOM.selectNodes("//oi[@tp = '" & TYPE_QUESTION & "' $and$ @hidden='0']").length > 0) Then
			sPersonalized = asDescriptors(119) 	'Descriptor: Yes
			sPersonalization = ""
			For each oCurrQO in oCacheDOM.selectNodes("//oi[@tp = '" & TYPE_QUESTION & "' $and$ @hidden='0']")
				If Len(oCurrQO.selectSingleNode("answer").getAttribute("n")) > 0 Then
					sPersonalization = sPersonalization & oCurrQO.getAttribute("n") & ": " & oCurrQO.selectSingleNode("answer").getAttribute("n") & ", "
				End If
			Next
		    If Len(sPersonalization) > 0 Then
				sPersonalization = Left(sPersonalization, Len(sPersonalization) - 2)
            	sPersonalized = sPersonalized & " ( "	& sPersonalization	& " ) "
			End If
        Else
            sPersonalized = asDescriptors(118) 'Descriptor: No
        End If
    End If

	Set oCacheDOM = Nothing
	Set oSub = Nothing
	Set oCurrQO = Nothing

	ReadCacheVariables_SubsConfirm = lErrNumber
	Err.Clear
End Function

Function ChangeStatusFlagToEdit(sCacheXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetNamedSchedulesForService"
	Dim lErrNumber
	Dim oCacheDOM
	Dim oSub

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, PAGE_NAME, PROCEDURE_NAME, "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		Set oSub = oCacheDOM.selectSingleNode("/mi/sub")
		Call oSub.setAttribute("sf", "0")
		sCacheXML = oCacheDOM.xml
		lErrNumber = Err.number
	End If

	ChangeStatusFlagToEdit = lErrNumber
	Err.Clear
End Function
%>