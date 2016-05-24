<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function getInfSources(aInformationSources)
'********************************************************
'*Purpose:  Returns the XML of all the Information sources in a project
'*Inputs:
'*Outputs:  sISXML: The XML with the list
'********************************************************
CONST PROCEDURE_NAME = "getInfSources"
Dim lErr
Dim sErr

Dim oSiteInfo
Dim sISXML
Dim sAllISXML

Dim oDOM
Dim oAllDOM
Dim oAllISs
Dim oIS
Dim i

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    If lErr = NO_ERR Then
        sAllISXML = oSiteInfo.getAllInformationSources(Application.Value("SITE_ID"))
        lErr = checkReturnValue(sAllISXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "getAllInformationSources", "Error calling getAllInformationSources", LogLevelError)
    End If

    If lErr = NO_ERR Then
        Set oAllDOM = Server.CreateObject("Microsoft.XMLDOM")
        oAllDOM.async = False
        If oAllDOM.loadXML(sAllISXML) = False Then
            lErr = ERR_XML_LOAD_FAILED
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "loadXML", "Error loading sAllISXML", LogLevelError)
        End If
    End If

    If lErr = NO_ERR Then
        sISXML = oSiteInfo.getInformationSourcesForSite(Application.Value("SITE_ID"))
        lErr = checkReturnValue(sISXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "getInformationSourcesForSite", "Error calling getInformationSourcesForSite", LogLevelError)
    End If

    If lErr = NO_ERR Then
        Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
        oDOM.async = False
        If oDOM.loadXML(sISXML) = False Then
            lErr = ERR_XML_LOAD_FAILED
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "loadXML", "Error loading sISXML", LogLevelError)
        End If
    End If

    If lErr = NO_ERR Then

        Set oAllISs = Nothing
        'filter out the IS id for "User Details" and "Subscription Info" and "System Information" and "Delivery Notification". These are special purpose IS's.
        Set oAllISs = oAllDOM.selectNodes("//oi[@tp='6' && not(@id='56D4246394B211D4BE6600C04F0E93B7' || @id='6A99F49197E311d5BEBC00B0D02A22CD'|| @id='429AE7760D404BBEBC5270EB5E3C8A52'|| @id='F6D4246394B211D4BE6600C04F0E93B7')]")

        If oAllISs.length > 0 Then
            Redim aInformationSources(oAllISs.length - 1, 4)

            For i = 0 To oAllISs.length - 1
                aInformationSources(i, 0) = oAllISs.item(i).getAttribute("id")
                aInformationSources(i, 2) = oAllISs.item(i).getAttribute("n")
                aInformationSources(i, 3) = oAllISs.item(i).getAttribute("des")
                'Set oIS = oAllDOM.selectSingleNode("//oi[prs/pr[@id='ISM_physical' and @v=""" & aInformationSources(i, 0) & """]]")
                Set oIS = oAllDOM.selectSingleNode("//oi[@id=""" & aInformationSources(i, 0) & """]")
                If oIS Is Nothing Then
                    aInformationSources(i, 4) = ""
                Else
					aInformationSources(i, 4) = GetPropertyValue(oIS, "IS_ServerName")
                End If

                Set oIS = oDOM.selectSingleNode("//oi[prs/pr[@id='ISM_physical' and @v=""" & aInformationSources(i, 0) & """]]")
                If oIS Is Nothing Then
                    aInformationSources(i, 1) = "2"
                Else
                    aInformationSources(i, 1) = GetPropertyValue(oIS, "ISM_required")
                End If

            Next

        End If

    End If

    Set oSiteInfo = Nothing
    Set oAllDOM = Nothing
    Set oDOM = Nothing
    Set oIS = Nothing
    Set oAllISs = Nothing

    getInfSources = lErr
    Err.Clear

End Function

Function setInfSources(aInformationSources())
'********************************************************
'*Purpose:  Saves the data for InformationSources into MD
'*Inputs:   aInformationSources: An array with the IS to save:
'*Outputs:  sISXML: The XML with the list
'********************************************************
CONST PROCEDURE_NAME = "setInfSources"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo
Dim oDOM
Dim oISs

Dim sISXML
Dim sSiteId
Dim sISId

    On Error Resume Next
    lErr = NO_ERR


    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    sSiteId = Application.Value("SITE_ID")

    If lErr = NO_ERR Then
        sISXML = oSiteInfo.getInformationSourcesForSite(sSiteId)
        lErr = checkReturnValue(sISXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "getInformationSourcesForSite", "Error calling getInformationSourcesForSite", LogLevelError)
    End If

    If lErr = NO_ERR Then
        Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
        oDOM.async = False
        If oDOM.loadXML(sISXML) = False Then
            lErr = ERR_XML_LOAD_FAILED
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "loadXML", "Error loading sISXML", LogLevelError)
        End If
    End If

    If lErr = NO_ERR Then
        Set oISs = oDOM.selectNodes("//oi[@tp='6']")

        'DELETE ALL IS Info from MD:
        For i = 0 to  oISs.length - 1
            sReturn = oSiteInfo.deleteObject(sSiteId, oISs.item(i).getAttribute("id"))
            lErr = checkReturnValue(sReturn, sErr)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "deleteObject", "Error calling deleteObject", LogLevelError)
                Exit For
            End If
        Next

    End If


    If lErr = NO_ERR Then
        'Create the XML to update IS properties:
        For i = 0 to UBound(aInformationSources, 1)

            If aInformationSources(i, 0) <> "" Then
                If aInformationSources(i, 1) <> "2" Then
                    sISId = GetGUID()
                    sISXML = "<mi><in><oi tp='6'><prs>"
                    sISXML = sISXML & " <pr id='INF_SOURCE_DISPLAY' v='1' />"
                    sISXML = sISXML & " <pr id='INF_SOURCE_REQUIRED'  v='" & aInformationSources(i, 1) & "' />"
                    sISXML = sISXML & " <pr id='INF_SOURCE_PHYSICAL' v='" &  aInformationSources(i, 0) & "' />"
                    sISXML = sISXML & "</prs></oi></in></mi>"

                    sReturn = oSiteInfo.CreateObject(sSiteId, sSiteId, sISId, sISXML)
                    lErr = checkReturnValue(sReturn, sErr)
                    If lErr <> NO_ERR Then
                        Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ISCoLib.asp", PROCEDURE_NAME, "deleteObject", "Error calling deleteObject", LogLevelError)
                        Exit For
                    End If
                End If
            End If
        Next

    End If

    Set oSiteInfo = Nothing
    Set oDOM = Nothing
    Set oISs = Nothing

    setInformationSources = lErr
    Err.Clear

End Function


Function IS_getPropertyValue(oDOM, sPropertyId)
'********************************************************
'*Purpose:  Returns the value of a property (if exists)
'           searching inside the DOM Object
'*Inputs:   oDOM: A valid DOM object (probably from a getSiteProperties call)
'           sPropertyId: The id of the property we're looking for.
'*Outputs:  The property Value (if exists) or ""
'********************************************************
Dim oNode
Dim sValue

    On Error Resume Next

    Set oNode = Nothing
    Set oNode = oDom.selectSingleNode(".//pr[@n=""" & sPropertyId & """]")

    'By default Return an Empty Value, if the node exist, return its value:
    sValue = ""
    If Not oNode Is Nothing Then
        sValue = oNode.getAttribute("v")
    End If

    Set oNode = Nothing

    IS_getPropertyValue = sValue
    Err.Clear

End Function


%>