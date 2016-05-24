<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%

'*** aDocInfo vars:
Public Const DOC_SUBS_ID        = 0
Public Const DOC_SVC_NAME       = 1
Public Const DOC_LAST_UPDATE    = 2
Public Const DOC_BODY           = 3
Public Const DOC_PORTAL_ADD     = 4
Public Const DOC_EXPIRATION     = 5

'*** aAttachmentInfo vars:
Public Const ATT_DOC_ID         = 0
Public Const ATT_INDEX          = 1
Public Const ATT_TYPE           = 2
Public Const ATT_BODY           = 3

'=--File type constants:
Private Const TYPE_ENCODED = "application/octet-stream"


Function co_GetSubscriptionContent(sSessionID, sSubscriptionID, sContentXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetSubscriptionContent"
    Dim oDocRepository
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oDocRepository = Server.CreateObject(PROGID_DOC_REPOSITORY)
    If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "", "Error when creating the DocRepository", LogLevelError)
    Else
        sContentXML  = oDocRepository.readSubscription(sSessionID, sSubscriptionID)
        lErrNumber = checkReturnValue(sContentXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, sErr, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "readSubscription", "Error calling readSubscription", LogLevelError)
        End If
    End If

    Set oDocRepository = Nothing

    co_GetSubscriptionContent = lErrNumber
    Err.Clear
End Function


Function co_GetAttachmentContent(sSessionID, sDocumentID, iIndex, sContentXML)
'********************************************************
'*Purpose: Given a docId and Index, returns the content of the attachment
'*Inputs: sSessionID, sDocumentID, iIndex
'*Outputs: sContentXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetAttachmentContent"
    Dim oDocRepository
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oDocRepository = Server.CreateObject(PROGID_DOC_REPOSITORY)
    If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "", "Error when creating the DocRepository", LogLevelError)
    Else
        sContentXML  = oDocRepository.readAttachment(sSessionID, sDocumentID, iIndex)
        lErrNumber = checkReturnValue(sContentXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, sErr, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "ReadSubscription", "Error trying to read the attachment", LogLevelError)
        End If
    End If

    Set oDocRepository = Nothing

    co_GetAttachmentContent = lErrNumber
    Err.Clear
End Function


Function co_GetAttachmentType(sSessionID, sDocumentID, iIndex, sContentXML)
'********************************************************
'*Purpose: Given a docId and Index, returns the content type of the attachment.
'           Similar to co_GetAttachmentContent, except that it only returns types, not the content.
'*Inputs: sSessionID, sDocumentID, iIndex
'*Outputs: sContentXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetAttachmentType"
    Dim oDocRepository
    Dim lErrNumber
    Dim sErr

    lErrNumber = NO_ERR

    Set oDocRepository = Server.CreateObject(PROGID_DOC_REPOSITORY)
    If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "", "Error when creating the DocRepository", LogLevelError)
    Else
        sContentXML  = oDocRepository.getAttachmentType(sSessionID, sDocumentID, iIndex)
        lErrNumber = checkReturnValue(sContentXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, sErr, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "ReadSubscription", "Error trying to read the attachment type.", LogLevelError)
        End If
    End If

    Set oDocRepository = Nothing

    co_GetAttachmentType = lErrNumber
    Err.Clear
End Function


Function co_GetUserSubscriptions(sSessionID, sChannelID, sSubsXML)
'********************************************************
'*Purpose: Returns the Subscriptions of the user for the current channel.
'*Inputs:  none
'*Outputs: sSubsXML
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetUserSubscriptions"
    Dim lErrNumber
    Dim oSubscription
    Dim sErr

    lErrNumber = NO_ERR

    Set oSubscription = Server.CreateObject(PROGID_SUBSCRIPTION)
    If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "", "Error when creating the Subscription object", LogLevelError)
    Else
        sSubsXML = oSubscription.getUserSubscriptions(sSessionID, sChannelID)
        lErrNumber = checkReturnValue(sSubsXML, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErrNumber, sErr, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "getUserSubscriptions", "Error when requesting the subscriptions", LogLevelError)
        End If
    End If

    Set oSubscription = Nothing

    co_GetUserSubscriptions = lErrNumber
    Err.Clear

End Function

Function co_getDocumentId(sSubscriptionId, sDocumentId)
'********************************************************
'*Purpose:  Returns the document Id for a given subscription.
'*Inputs:
'*Outputs:  sEngine: machine name where the communication server is located
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_getDocumentId"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oDocRep
    Dim oDOM

    lErr = NO_ERR

    If Len(sSubscriptionId) > 0 Then

        sDocumentId = ""

        'Change site configuration on subscriptionPortal.properties file:
        If lErr = NO_ERR Then
            Set oDocRep = Server.CreateObject(PROGID_DOC_REPOSITORY)

            sReturn = oDocRep.getDocumentID(aConnectionInfo(S_TOKEN_CONNECTION), sSubscriptionId)
            lErr = checkReturnValue(sReturn, sErr)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "getDocumentID", "Error calling getDocumentID", LogLevelError)
        End If

        If lErr = NO_ERR Then
            lErr = LoadXMLDOMFromString(aConnectionInfo, sReturn, oDOM)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, Err.description, CStr(Err.source), "ReportsCoLib.asp", PROCEDURE_NAME, "loadXML", "Error loading XML from getDocumentID", LogLevelError)
            End If
        End If

        If lErr = NO_ERR Then
            If Not oDOM.selectSingleNode("mi/in") Is Nothing Then
                sDocumentId = oDOM.selectSingleNode("mi/in").getAttribute("id")
            End If
        End If

    End If

    Set oDocRep = Nothing
    Set oDOM = Nothing

    co_getDocumentId = lErr
    Err.Clear
End Function

%>
