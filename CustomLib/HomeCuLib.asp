<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/AddressCoLib.asp" -->
<%
Public Function cu_GetSiteInfo(sSitesXML, aSiteInfo)

    Dim aSitesDOM
    Dim oSite
    Dim lErr
    Dim oFolder

    Set aSitesDOM = Server.CreateObject("Microsoft.XMLDOM")
    aSitesDOM.async = False
    aSitesDOM.loadXML(sSitesXML)

    If Err.number <> NO_ERR Then
        lErr = ERR_XML_LOAD_FAILED
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "HomeCuLib.asp", "cu_GetSiteInfo", "", "Error loading sSitesXML", LogLevelError)
    End If

    If lErr = NO_ERR Then
        Set oSite = aSitesDOM.selectSingleNode("/mi/in/oi[@id = '" & GetCurrentChannel() & "']")

        If oSite Is Nothing Then
        	Set oSite = aSitesDOM.selectSingleNode("/mi/in/oi[0]")
        End If

		If oSite Is Nothing Then
			lErr = ERR_RETRIEVING_RESULTS
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "HomeCuLib.asp", "cu_GetSiteInfo", "", "Could not find current site Node", LogLevelError)
		Else
			Set oFolder = oSite.selectSingleNode("prs/pr[@id='serviceFolderID']")
			If oFolder Is Nothing Then
				lErr = ERR_RETRIEVING_RESULTS
				Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "HomeCuLib.asp", "cu_GetSiteInfo", "", "Could not find current site Node", LogLevelError)
			End If
		End If
    End If

    If lErr = NO_ERR Then
        aSiteInfo(0) = oSite.getAttribute("n")
        aSiteInfo(1) = oSite.getAttribute("des")
        aSiteInfo(2) = ofolder.getAttribute("v")
    End If

    Set aSitesDOM = Nothing
    Set oSite = Nothing
    Set oFolder = Nothing

    cu_GetSiteInfo = lErr
    Err.Clear

End Function
%>