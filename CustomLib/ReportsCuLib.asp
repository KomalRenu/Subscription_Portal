<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/ReportsCoLib.asp" -->
<%


Function ParseRequestForReports(oRequest, aDocInfo)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Set folderID according to an "ROOT FOLDER" setting
'********************************************************
    On Error Resume Next

    aDocInfo(DOC_SUBS_ID) = Trim(CStr(oRequest("subsId")))
    aDocInfo(DOC_PORTAL_ADD) = GetPortalAddress()

    'Log Error if necessary:
    If Err.number <> NO_ERR Then Call LogErrorXML(aConnectionInfo, Err.number, Err.description, CStr(Err.source), "ReportsCuLib.asp", "ParseRequestForReports", "", "Error parsing request", LogLevelError)
    ParseRequestForReports = Err.number
    Err.Clear

End Function

Function GetSubscriptionContent(aDocInfo, sContentXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "GetSubscriptionContent"
    Dim lErrNumber
    Dim sSessionID
    Dim sSubscriptionID

    lErrNumber = NO_ERR
    sSessionID = GetSessionID()
    sSubscriptionID = aDocInfo(DOC_SUBS_ID)

    lErrNumber = co_GetSubscriptionContent(sSessionID, sSubscriptionID, sContentXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetSubscriptionContent", LogLevelTrace)
    End If

    GetSubscriptionContent = lErrNumber
    Err.Clear
End Function

Function GetUserSubscriptions_Reports(sSubsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "GetUserSubscriptions_Reports"
    Dim lErrNumber
    Dim sSessionID
    Dim sChannelID

    lErrNumber = NO_ERR
    sSessionID = GetSessionID()
    sChannelID = GetCurrentChannel()

    lErrNumber = co_GetUserSubscriptions(sSessionID, sChannelID, sSubsXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetUserSubscriptions", LogLevelTrace)
    End If

    GetUserSubscriptions_Reports = lErrNumber
    Err.Clear
End Function

Function GetDocumentContent(aDocInfo, sSubsXML, sGetAvailableSubscriptionsXML)
'********************************************************
'*Purpose: Gets the Information associated with a document that will be displayed
'*Inputs:
'   aDocInfo:    An aray where the information of a document will be stored
'   sContentXML: The document content in XML format.
'   sSubsXML:    The list of subscriptions in XML format
'*Outputs:
'   aDocInfo
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "GetDocumentContent"
    Dim lErr
    Dim sContentXML

    lErr = NO_ERR

    'If no SubsId, return all as information not available:
    If Len(aDocInfo(DOC_SUBS_ID)) = 0 Then
        aDocInfo(DOC_DATA) = asDescriptors(359) 'Descriptor: Information not available
        aDocInfo(DOC_LAST_UPDATE) = asDescriptors(359) 'Descriptor: Information not available
        aDocInfo(DOC_SVC_NAME) = asDescriptors(359) 'Descriptor: Information not available
    Else
        If lErr = NO_ERR Then
            lErr = co_GetSubscriptionContent(GetSessionID(), aDocInfo(DOC_SUBS_ID), sContentXML)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, "", "", "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetSubscriptionContent", LogLevelTrace)
                If lErr = ERR_DOC_BODY_NOT_FOUND Then
                    lErr = NO_ERR
                    aDocInfo(DOC_BODY) = asDescriptors(359) 'Descriptor: Information not available
                End If
            Else
                lErr = GetDocBody(aDocInfo, sContentXML)
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "", "Error calling GetDocContent", LogLevelTrace)
            End If
        End If

        If lErr = NO_ERR Then
            lErr = GetDocName(aDocInfo, sSubsXML)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "", "Error calling GetDocName", LogLevelTrace)
        End If

        If lErr = NO_ERR Then
            lErr = GetLastUpdate(aDocInfo, sGetAvailableSubscriptionsXML)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "", "Error calling GetLastUpdate", LogLevelTrace)
        End If
    End If

    GetDocumentContent = lErr
    Err.Clear
End Function

Function GetDocBody(aDocInfo, sContentXML)
'********************************************************
'*Purpose: Gets the Body associated with a document
'*Inputs:
'   aDocInfo:    An aray where the information of a document will be stored
'   sContentXML: The document content in XML format.
'*Outputs:
'   aDocInfo(DOC_BODY)
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "GetDocBody"
    Dim oContent
    Dim oDecoder
    Dim sData
    Dim bData
    Dim lErr

    lErr = NO_ERR

    '=--Create DOM Object
    If lErr = NO_ERR Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sContentXML, oContent)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error loading sContentXML", LogLevelError)
        End If
    End If

    '=--Retrieve the Node with the data:
    If lErr = NO_ERR Then
        sData = "" & oContent.documentElement.selectSingleNode("oi").Text
        If Len(sData) = 0 Then
            lErr = ERR_XML_LOAD_FAILED
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Emtpy data returned", LogLevelError)
        End If
    End If

    '=--The DATA comes in MIME format, decode it:
    If lErr = NO_ERR Then
        Set oDecoder = Server.CreateObject(PROGID_BASE64)
        If Err.Number <> NO_ERR Then
            lErr = Err.number
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error creating Decoder", LogLevelError)
        Else
            bData = oDecoder.Decode(sData)
            If Err.Number <> NO_ERR Then
                lErr = Err.number
                Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error decoding content", LogLevelError)
            End If
        End If
    End If

    'We finally got to the content:
    If lErr = NO_ERR Then
	    aDocInfo(DOC_BODY) = bData
    End If

    Set oContent = nothing
    Set oDecoder = nothing

    GetDocBody = lErr
    Err.Clear
End Function

Function GetDocName(aDocInfo, sSubsXML)
'********************************************************
'*Purpose: Gets the name associated with a document, currently the
'           name is the name of the service that it is subscribed.
'*Inputs:
'   aDocInfo: An aray where the information of a document will be stored
'   sSubsXML: The XML with the list of subscriptions for the portal address.
'*Outputs:
'   aDocInfo(DOC_SVC_NAME)
'********************************************************
    On Error Resume Next
	Const PROCEDURE_NAME = "GetDocName"
	Dim oSubsDOM
	Dim sSvcId
	Dim oSub
    Dim lErr

    lErr = NO_ERR

    'Load SubsXML
    If lErr = NO_ERR Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sSubsXML, oSubsDOM)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error loading sSubsXML", LogLevelError)
        End If
    End If

    If lErr = NO_ERR Then
        Set oSub = oSubsDOM.selectSingleNode("//sub[@guid='" & aDocInfo(DOC_SUBS_ID) & "']")
        If (Err.number <> NO_ERR) Or (oSub Is Nothing) Then
            lErr = ERR_RETRIEVING_RESULTS
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Could not find the subscription node", LogLevelError)
        Else
            sSvcId = oSub.selectSingleNode("dst").getAttribute("svid")
            aDocInfo(DOC_SVC_NAME) = oSubsDOM.selectSingleNode("//oi[@id='" & sSvcId & "']").getAttribute("n")
            If (Err.number <> NO_ERR) Then
                lErr = Err.number
                Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Could not find the name of the service", LogLevelError)
            End If
        End If
    End If

    If lErr <> NO_ERR Then
        aDocInfo(DOC_SVC_NAME) = asDescriptors(359) 'Descriptor: Information not available
    End If

    Set oSubsDOM = Nothing
    Set oSub = Nothing

    GetDocName = lErr
    Err.Clear
End Function

Function GetLastUpdate(aDocInfo, sGetAvailableSubscriptionsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "GetLastUpdate"
    Dim oContent
    Dim oDecoder
    Dim sUpdate
    Dim sExpires
    Dim lErr

    lErr = NO_ERR

    '=--Create DOM Object
    If lErr = NO_ERR Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sGetAvailableSubscriptionsXML, oContent)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error loading sGetAvailableSubscriptionsXML", LogLevelError)
        End If
    End If

    '=--Retrieve the Node with the data:
    If lErr = NO_ERR Then
        sUpdate = oContent.selectSingleNode("/mi/subs/sub[@id = '" & aDocInfo(DOC_SUBS_ID) & "']").getAttribute("exct")
        If Len(sUpdate) = 0 Then
            'add error logging?
            sUpdate = asDescriptors(359) 'Descriptor: Information not available
        End If
        sExpires = oContent.selectSingleNode("/mi/subs/sub[@id = '" & aDocInfo(DOC_SUBS_ID) & "']").getAttribute("expt")
        If Len(sExpires) = 0 Then
            'add error logging?
            sExpires = asDescriptors(359) 'Descriptor: Information not available
        End If
    End If

    'We finally got to the content:
    If lErr = NO_ERR Then
	    aDocInfo(DOC_LAST_UPDATE) = DisplayDateAndTime(CDate(sUpdate),CDate(sUpdate))
	    aDocInfo(DOC_EXPIRATION) = DisplayDateAndTime(CDate(sExpires),CDate(sExpires))
    End If

    Set oContent = nothing
    Set oDecoder = nothing

    GetLastUpdate = lErr
    Err.Clear
End Function

Function GetAttachmentInfo(aAttachmentInfo)
'********************************************************
'*Purpose:
'*Inputs:
'   aAttatchmentInfo: An aray where the information of a document will be stored
'   sContentXML: The document content in XML format.
'*Outputs:
'   aAttachmentInfo
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "GetAttachmentInfo"
    Dim lErr
    Dim sAttachXML
    Dim sSessionID
    Dim sDocumentID
    Dim iIndex

    lErr = NO_ERR
    sSessionID = GetSessionID()
    sDocumentID = aAttachmentInfo(ATT_DOC_ID)
    iIndex = aAttachmentInfo(ATT_INDEX)

    'Get XML:
    If lErr = NO_ERR Then
        lErr = co_GetAttachmentContent(sSessionID, sDocumentID, iIndex, sAttachXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetAttachmentContent", LogLevelTrace)
    End If

    'Parse the XML into the array:
    If lErr = NO_ERR Then
        lErr = ParseAttachmentInfo(aAttachmentInfo, sAttachXML)
    End If

    GetAttachmentInfo = lErr
    Err.Clear
End Function

Function ParseAttachmentInfo(aAttachmentInfo, sAttachXML)
'********************************************************
'*Purpose:
'   Reads the attachment information from the XML string and set it
'   into the aAttachmentInfo array
'*Inputs:
'   sAttchXML and XML string with the attachment information
'*Outputs:
'   None.
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "ParseAttachmentInfo"
    Dim lErr
    Dim oDoc
    Dim oDecoder
    Dim oAttachment
    Dim sData

    'Create a DOM object to read the results, they come in XML format:
    If lErr = NO_ERR Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sAttachXML, oDoc)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error loading sAttachXML", LogLevelError)
        End If
    End If

    'If the XML is succesful, the Content is returned in the DATA node:
    If lErr = NO_ERR Then
        Set oAttachment = Nothing
        Set oAttachment = oDoc.documentElement.selectSingleNode("atmt")
        If oAttachment Is Nothing Then
            lErr = ERR_RETRIEVING_RESULTS
            Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Empty data returned", LogLevelError)
        End If
    End If

    'Retrieve the type of the attachment as well:
    If lErr = NO_ERR Then
        sData = "" & oAttachment.Text
        aAttachmentInfo(ATT_TYPE) = oAttachment.getAttribute("tp")
    End If

    If lErr = NO_ERR Then
        'If the data must be returned encoded:
        If aAttachmentInfo(ATT_TYPE) = TYPE_ENCODED Then
            aAttachmentInfo(ATT_BODY) = sData
        ElseIf Len(sData) > 0 Then
            'Return the data decoded:
            Set oDecoder = Server.CreateObject(PROGID_BASE64)
            lErr = Err.number
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error creating Decoder Object", LogLevelError)
            Else
                aAttachmentInfo(ATT_BODY) = oDecoder.Decode(sData)
                lErr = Err.number
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error decoding data", LogLevelError)
            End if
        End If
    End If

    Set oDoc = Nothing
    Set oDecoder = nothing
    Set oAttachment = Nothing

    ParseAttachmentInfo = lErr

End Function

Function StreamAttachmentContent(aAttachmentInfo)
'********************************************************
'*Purpose:
'   Reads the attachment information and streams it directly into the browser
'     instead of reading it all in one block. This to improve memory consumption
'     in large attachments
'*Inputs:
'   aAttatchmentInfo: An aray where the information of a document will be stored
'*Outputs:
'   None.
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "StreamAttachmentContent"
    Dim lErr
    Dim sSessionID
    Dim sDocumentID
    Dim iIndex
    Dim sAttachXML

    Dim bIsBinary
    Dim oDocRepository
	Dim arrDataBlock
	Dim lBlockCounter
	Dim lCurrentSize

    lErr = NO_ERR
    sSessionID = GetSessionID()
    sDocumentID = aAttachmentInfo(ATT_DOC_ID)
    iIndex = aAttachmentInfo(ATT_INDEX)

    'First get the attachment type
    If lErr = NO_ERR Then
        lErr = co_GetAttachmentType(sSessionID, sDocumentID, iIndex, sAttachXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetAttachmentContent", LogLevelTrace)
    End If

    'Parse the AttachmentInfo into the array
    If lErr = NO_ERR Then
        lErr = ParseAttachmentInfo(aAttachmentInfo, sAttachXML)
    End If

    'Check if it's a binary attachment, and set the correct ContentType:
    If lErr = NO_ERR Then
        bIsBinary = False

        If StrComp(CStr(aAttachmInfo(ATT_TYPE)), "text/html", vbTextCompare) <> 0 and StrComp(CStr(aAttachmInfo(ATT_TYPE)), "text/plain", vbTextCompare) <> 0 and StrComp(CStr(aAttachmInfo(ATT_TYPE)), "text/xml", vbTextCompare) <> 0 Then
            Response.ContentType = aAttachmInfo(ATT_TYPE)
            bIsBinary = True
        End If

    End If

	' Retrieve and stream attachment in blocks.
	If lErr = NO_ERR Then
        Set oDocRepository = Server.CreateObject(PROGID_DOC_REPOSITORY)
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, CStr(Err.source), "ReportsCuLib.asp", PROCEDURE_NAME, "", "Error when creating the DocRepository", LogLevelError)
        Else
	        lBlockCounter = 1
	        lCurrentSize = 0

            Do
                arrDataBlock = oDocRepository.readAttachmentBlock(sSessionId, sDocumentID, iIndex, lBlockCounter)
                lCurrentSize = UBound(arrDataBlock)
                lBlockCounter = lBlockCounter + 1

                If lCurrentSize > 0 Then
                    If (bIsBinary) Then
            	        Response.BinaryWrite(arrDataBlock)
            	    Else
            	        Response.Write(arrDataBlock)
            	    End If
            	End If

            Loop While (lCurrentSize > 0)
	    End If
    End If

    Set oDocRepository  = Nothing

    StreamAttachmentContent = lErr
    Err.Clear
End Function

Function RenderSubscriptionContent(aDocInfo)
'********************************************************
'*Purpose: Present the content of a document.
'*Inputs:  aDocInfo: The document information.
'*Outputs: Nothing
'********************************************************


    'Check if the DocInfo has a valid object:
    If Len(CStr(aDocInfo(DOC_SUBS_ID))) = 0 Then

        'If not, no document was requested, show a fram asking the user
        'to select an element:
        Response.Write "<TABLE border=""0""  CELLPADDING=0 CELLSPACING=0 >"
        'Response.Write "    <TR>"
        'Response.Write "        <TD COLSPAN=""6"" HEIGHT=""20""><IMG SRC=""images/1pTRans.gif"" HEIGHT=""20""/></TD>"
        'Response.Write "    </TR>"
        Response.Write "    <TR>"
        Response.Write "        <TD COLSPAN=""5"" BGCOLOR=""#000000"" HEIGHT=""1""><IMG SRC=""images/1pTRans.gif"" HEIGHT=""1""/></TD>"
        Response.Write "        <TD ROWSPAN=""2"" WIDTH=""10"" HEIGHT=""1"" ><IMG SRC=""images/1pTRans.gif"" HEIGHT=""1""/></TD>"
        Response.Write "    </TR>"
        Response.Write "    <TR>"
        Response.Write "        <TD ROWSPAN=""3"" WIDTH=""1"" BGCOLOR=""#000000""><IMG SRC=""images/1pTRans.gif"" HEIGHT=""1""/></TD>"
        Response.Write "        <TD BGCOLOR=""FFFFCC"" COLSPAN=""3"" heigth=""10""><IMG SRC=""images/1pTRans.gif"" HEIGHT=""10"" /></TD>"
        Response.Write "        <TD ROWSPAN=""3"" WIDTH=""1"" BGCOLOR=""#000000""><IMG SRC=""images/1pTRans.gif"" /></TD>"
        Response.Write "    </TR>"
        Response.Write "    <TR>"
        Response.Write "        <TD BGCOLOR=""FFFFCC"" WIDTH=""10"">&nbsp;</TD>"
        Response.Write "        <TD BGCOLOR=""FFFFCC"">"
        Response.Write "          <b><FONT face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_MEDIUM_FONT) & """>" & asDescriptors(449) & "</FONT></b>" 'Descriptor: Click on a document name to the left to view that document here.
        Response.Write "        </TD>"
        Response.Write "        <TD BGCOLOR=""FFFFCC"" WIDTH=""10"">&nbsp;</TD>"
        Response.Write "        <TD ROWSPAN=""4"" WIDTH=""1"" BGCOLOR=""#CCCCCC""><IMG SRC=""images/1pTRans.gif""/></TD>"
        Response.Write "    </TR>"
        Response.Write "    <TR>"
        Response.Write "        <TD BGCOLOR=""FFFFCC"" COLSPAN=""3"" heigth=""10""><IMG SRC=""images/1pTRans.gif"" HEIGHT=""10"" /></TD>"
        Response.Write "    </TR>"
        Response.Write "    <TR>"
        Response.Write "        <TD COLSPAN=""5"" BGCOLOR=""#000000"" HEIGHT=""1""><IMG SRC=""images/1pTRans.gif"" /></TD>"
        Response.Write "    </TR>"
        Response.Write "    <TR>"
        Response.Write "        <TD COLSPAN=""2""><IMG SRC=""images/1pTRans.gif"" HEIGHT=""10""/></TD>"
        Response.Write "        <TD COLSPAN=""3"" BGCOLOR=""#CCCCCC""><IMG SRC=""images/1pTRans.gif"" HEIGHT=""10""/></TD>"
        Response.Write "    </TR>"
        Response.Write "</TABLE>"

    Else

        'If there is a document, show the Document Title
        Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
        Response.Write "  <TR>"
        Response.Write "    <TD COLSPAN=""2"" BGCOLOR=""#636563"" WIDTH=""11"" ALIGN=""LEFT"" VALIGN=""TOP""><IMG SRC=""images/loginUpperLeftCorner.gif"" WIDTH=""11"" HEIGHT=""11"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "    <TD BGCOLOR=""#636563"" WIDTH=""500"">"
        Response.Write "      <FONT face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#ffffff"" size=""" & aFontInfo(N_MEDIUM_FONT) & """>"
        Response.Write "       <B>&nbsp;" & aDocInfo(DOC_SVC_NAME) & "</B>"
        Response.Write "      </FONT>"
        Response.Write "    </TD>"
        Response.Write "    <TD BGCOLOR=""#636563"" ALIGN=""RIGHT"" VALIGN=""TOP"" NOWRAP=""1"">"
        Response.Write "      <A HREF=""printReport.asp?subsId=" & aDocInfo(DOC_SUBS_ID) & """ TARGET=""PrintWindow"" /><IMG SRC=""images/tb_preview.gif"" BORDER=""0"" ALT=""" & asDescriptors(61) & """></A>" 'Descriptor: Printable version
        Response.Write "    </TD>"
        Response.Write "    <TD BGCOLOR=""#636563"" WIDTH=""50""><IMG SRC=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "    <TD COLSPAN=""2"" BGCOLOR=""#636563"" ALIGN=""RIGHT"" VALIGN=""TOP""><IMG SRC=""images/loginUpperRightCorner.gif"" WIDTH=""11"" HEIGHT=""11"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "  </TR>"
        Response.Write "  <TR>"
        Response.Write "    <TD BGCOLOR=""#636563"" WIDTH=""1""><IMG SRC=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "    <TD BGCOLOR=""#CCCCCC"" WIDTH=""11"" ALIGN=""LEFT"" VALIGN=""TOP""><IMG SRC=""images/1ptrans.gif"" WIDTH=""11"" HEIGHT=""11"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "    <TD BGCOLOR=""#CCCCCC"" COLSPAN=""3"">"
        Response.Write "      <FONT face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """>"
        Response.Write "       <B>&nbsp;" & asDescriptors(38) & " " & "</B>" & aDocInfo(DOC_LAST_UPDATE) & "<BR />" 'Descriptor: Last Update:
        Response.Write "       <B>&nbsp;" & asDescriptors(406) & " " & "</B>" & aDocInfo(DOC_EXPIRATION) & "<BR />"
        Response.Write "      </FONT>"
        Response.Write "    </TD>"
        Response.Write "    <TD BGCOLOR=""#CCCCCC"" WIDTH=""11"" ALIGN=""LEFT"" VALIGN=""TOP""><IMG SRC=""images/1ptrans.gif"" WIDTH=""11"" HEIGHT=""11"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "    <TD BGCOLOR=""#636563"" WIDTH=""1""><IMG SRC=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "  <TR>"
        Response.Write "    <TD BGCOLOR=""#636563"" WIDTH=""1"" HEIGHT=""1"" COLSPAN=""8""><IMG SRC=""images/1ptrans.gif"" WIDTH=""1"" HEIGHT=""1"" ALT="""" BORDER=""0"" /></TD>"
        Response.Write "  </TR>"
        Response.Write "</TABLE><BR>"

        'Now show the content of the document:
        Response.Write aDocInfo(DOC_BODY)

    End If

    Err.Clear

End Function

Function RenderReportPath(aDocInfo)
'********************************************************
'*Purpose:  Shows the path of this document
'*Inputs:   aDocInfo: with the information of a document
'*Outputs:  none
'********************************************************

    'Start saying "You are here:"
    Response.Write "<FONT face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(26) & "</FONT>" 'Descriptor: You are here:

    If Len(CStr(aDocInfo(DOC_SUBS_ID))) = 0 Then
        'If no subsId, there was no request, so we're in the root, that is
        'in the reports tab:
        Response.Write " <B><FONT face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(360) & "</FONT></B>" 'Descriptor: Reports
    Else
        'If there is a document, for the moment, they all come
        'from the Reports Tab:
        Response.Write " <A HREF=""reports.asp""><FONT face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """>" & asDescriptors(360) & "</A> > </FONT>" 'Descriptor: Reports
        Response.Write "<B><FONT face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """>" & aDocInfo(DOC_SVC_NAME) & "</FONT></B>"
    End If

    Err.Clear

End Function

Function RenderPortalDocumentsList(sSubsXML, sGetAvailableSubscriptionsXML, sCurrentSubsId)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Remove the getDocumentId from here
'********************************************************
Dim oSubsDOM
Dim oSubs
Dim oSub
Dim sResult
Dim bReady
Dim sSubsId
Dim sSvcId
Dim sSchId
Dim sDocId

Dim oAvailableSubsDOM

Dim sLinkColor
Dim sBulletImage
Dim sBeginBold
Dim sEndBold
Dim sBeginLink
Dim sEndLink
Dim sStatus

    On Error Resume Next

    'Load SubsXML
    Set oSubsDOM = Server.CreateObject("Microsoft.XMLDOM")
    oSubsDOM.async = False
    If oSubsDOM.loadXML(sSubsXML) = False Then
        'RenderPortalDocumentsList = Err.number
        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#cc0000"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(405) & "</b></font>" 'Descriptor: Error retrieving reports
        Exit Function
    End If

    Response.Write "<TABLE CELLSPACING=0 CELLPADDING=2 BORDER=0>"

    'Check if there are any subscriptions at all
    Set oSubs = oSubsDOM.selectNodes("//sub[@adid='" & aDocInfo(DOC_PORTAL_ADD) & "']")
    If oSubs.length > 0 Then

        Call GetXMLDOM(aConnectionInfo, oAvailableSubsDOM, sErrDescription)
        oAvailableSubsDOM.async = False
        Call oAvailableSubsDOM.loadXML(sGetAvailableSubscriptionsXML)

        'We need to remove this from here:
        For Each oSub in oSubs

            'Get all necessary Ids:
            sSubsId = oSub.getAttribute("guid")
            sSvcId  = oSub.selectSingleNode("dst").getAttribute("svid")
            sSchId  = oSub.selectSingleNode("dst").getAttribute("scid")

            'Check if the subscriptions is ready, by checking if it is already in the DocRepository
            If (oAvailableSubsDOM.selectSingleNode("/mi/subs/sub[@id = '" & sSubsId & "']") Is Nothing) Then
                bReady = False
            Else
                bReady = True
            End If

            If bReady Then
                sBulletImage = "<img src=""images/bullet.gif"" BORDER=""0"" ALT="""">"
                sBeginLink = "<A HREF=""reports.asp?subsId=" & sSubsId & """ >"
                sEndLink = "</A>"
                sStatus = ""
            Else
                sBulletImage = "<img src=""images/bullet.gif"" BORDER=""0"" ALT="""">"
                sBeginLink = ""
                sEndLink = ""
                sStatus = "<BR><I>(" & asDescriptors(407) & ")</I>" 'Descriptor: Pending
            End If

		        If sCurrentSubsId = sSubsId then
		            sLinkColor = "#cc0000"
		            sBeginBold = "<B>"
		            sEndBold = "</B>"
		        Else
		            sLinkColor = "#000000"
		            sBeginBold = ""
		            sEndBold = ""
		        End if


		        Response.Write "<TR>" & chr(13) & chr(10)
		        Response.Write "<TD VALIGN=TOP>" & sBulletImage & "</TD>"
		        Response.Write "<TD>"
		        Response.Write sBeginLink
		        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sLinkColor & """>"
		        Response.Write "<B>" & oSubsDOM.selectSingleNode("//oi[@id='" & sSvcId & "']").getAttribute("n")  &  "</B>"
		        Response.Write "</font><br>"
		        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """ color=""" & sLinkColor & """>"
		        Response.Write oSubsDOM.selectSingleNode("//oi[@id='" & sSchId & "']").getAttribute("n")
		        Response.Write "</font>"
		        Response.Write sEndLink
		        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & sStatus & "</font>"
		        Response.Write "</TD>"
		        Response.Write "</TR>" & chr(13) & chr(10)
		        Response.Write "<TR><TD COLSPAN=2><IMG SRC=""images/1ptrans.gif"" HEIGHT=""3"" WIDTH=""1"" ALT="""" BORDER=""0""></TD></TR>"

		    Next

    Else
        Response.Write "<TR>"
        Response.Write "<TD VALIGN=TOP>"
        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#cc0000"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(408) & "</b></font>" 'Descriptor: There are no reports in this folder.
        Response.Write "</TR>"
        Response.Write "</TD>"

    End If

    Response.Write "</TABLE>"


    Set oSubsDOM = Nothing
    Set oSubs = Nothing
    Set oSub = Nothing
    Set oAvailableSubsDOM = Nothing

    Err.Clear

End Function

Function RenderPortalDocumentsIcons(sSubsXML, sCurrendSubsId)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Set folderID according to an "ROOT FOLDER" setting
'********************************************************
	Dim oSubsDOM
	Dim oSvcs
	Dim oSvc
	Dim oSubs
	Dim oSub
	Dim sSubsId
	Dim bReady

	Dim i
	Dim iCellCounter

	Dim sLinkColor
	Dim sBulletImage
	Dim sBeginLink
	Dim sEndLink
	Dim sPostFix
	Dim sStatus

	iCellCounter = 1

    'Load SubsXML
    Set oSubsDOM = Server.CreateObject("Microsoft.XMLDOM")
    oSubsDOM.async = False
    If oSubsDOM.loadXML(sSubsXML) = False Then
		'RenderPortalDocumentsIcons = Err.number
		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#cc0000"" size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & asDescriptors(405) & "</b></font>" 'Descriptor: Error retrieving reports
        Exit Function
    End If

	'Check if there are any subscriptions
	Set oSubs = oSubsDOM.selectNodes("//SUBSCRIPTION")
	If oSubs.length > 0 Then

		'Get all the services from the subscriptions:
		Set oSvcs = oSubsDOM.selectNodes("//SERVICES/SERVICE")
	    Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 WIDTH=""100%"">"

		For Each oSvc In oSvcs

		     'Check for all subscriptions of a service (ideally it should be only 1 per service, but you know...)
		     Set oSubs = oSvc.selectNodes("SUBSCRIPTIONS/SUBSCRIPTION")
		     i = 0
		     For Each oSub in oSubs
	            'Check if the subscriptions is ready with the available flag in the attribute
		        'of the subscription
                If oSub.Attributes.getNamedItem("AVAILABLE") Is Nothing Then
                    bReady = False
                Else
                    bReady = oSub.getAttribute("AVAILABLE") = "yes"
                End If

                If oSub.Attributes.getNamedItem("SUBSCRIPTION_GUID") Is Nothing Then
                    bReady = False
                Else
                    sSubsId = oSub.getAttribute("SUBSCRIPTION_GUID")
                End If

		        If bReady Then
		            sBulletImage = "<img src=""images/graph_big.gif"" BORDER=""0"" ALT="""">"
			        sLinkColor = "#cc0000"
			        sBeginLink = "<A HREF=""reports.asp?subsId=" & sSubsId & """ >"
			        sEndLink = "</A>"
			        sStatus = ""
                else
		            sBulletImage = "<img src=""images/graph_big.gif"" BORDER=""0"" ALT="""">"
			        sLinkColor = "#cc0000"
			        sBeginLink = ""
			        sEndLink = ""
			        sStatus = " (" & asDescriptors(407) & ")" 'Descriptor: Pending
		        End If

		        If i = 0 Then
		            sPostFix = ""
		        Else
		            sPostFix = " (" & CStr(i) & ")"
		        End if

		        If iCellCounter = 1 Then
			       Response.Write "<TR><TD WIDTH=""50%"">"
			    Else
			       Response.Write "<TD WIDTH=""50%"">"
		        End If

		        Response.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0><TR>"
		        Response.Write "<TD VALIGN=TOP>" & sBulletImage & "</TD>"
		        Response.Write "<TD VALIGN=TOP><font face=""" & aFontInfo(S_FAMILY_FONT) & """><font size=""" & aFontInfo(N_MEDIUM_FONT) & """ color=""" & sLinkColor & """>"
		        Response.Write "<b>" & sBeginLink & oSvc.getAttribute("SERVICE_NAME") & sPostFix & sEndLink & "</b></font>"
		        Response.Write "<BR /><font size=""" & aFontInfo(N_SMALL_FONT) & """>" & sStatus & "</font></font></TD>"
		        Response.Write "</TR></TABLE><BR />"' & chr(13) & chr(10)

		        If iCellCounter = 1 Then
				   	Response.Write "</TD>"
				   	iCellCounter = 2
				   Else
				   	Response.Write "</TD></TR>"
				   	iCellCounter = 1
		        End If

		        i = i + 1

		    Next
		Next

		If iCellCounter = 2 Then
			Response.Write "<TD WIDTH=""50%""></TD></TR>"
		End If

		Response.Write "</TABLE>"
	Else
		Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#cc0000"" size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & asDescriptors(408) & "</b></font>" 'Descriptor: There are no reports in this folder.
	End If

	Set oSubsDOM = Nothing
	Set oSvcs = Nothing
	Set oSvc = Nothing
	Set oSubs = Nothing
	Set oSub = Nothing

    Err.Clear

End Function
%>
