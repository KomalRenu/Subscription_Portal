<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'
Private Const O_INBOXXML_JOBERROR = 0
Private Const S_MSG_JOBERROR = 1
Private Const S_FOLDER_ID_JOBERROR = 2
Private Const S_OK_JOBERROR = 3
Private Const L_ERROR_NUMBER_JOBERROR = 4
Private Const S_ERROR_DESCRIPTION_JOBERROR = 5
Private Const S_REQUEST_JOBERROR = 6
Private Const S_PREVIOUS_PAGE_JOBERROR = 7
Private Const O_RESULTSET_JOBERROR = 8

Private Const MAX_REQUEST_JOBERROR = 8

Function ReceiveJobErrorRequest(oRequest, aConnectionInfo, aJobErrorRequest, aObjectInfo, aFolderInfo, sErrDescription)
'******************************************************************************
'Purpose: Read the variables from the request object
'Inputs:  oRequest
'Outputs: aJobErrorRequest, aConnectionInfo, aObjectInfo, aFolderInfo , sErrDescription
'******************************************************************************
	On Error Resume Next
	ReDim Preserve aJobErrorRequest(MAX_REQUEST_JOBERROR)
	ReDim aObjectInfo(MAX_OBJECT_INFO)
	ReDim aFolderInfo(MAX_OBJECT_INFO)

	If Not IsObject(aJobErrorRequest(O_RESULTSET_JOBERROR)) Then
		lErrNumber = CreateResultSetHelperObject(aConnectionInfo, aJobErrorRequest(O_RESULTSET_JOBERROR), sErrDescription)
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "JobErrorCuLib.asp", "ReceiveJobErrorRequest", "", "Error after calling CreateResultSetHelperObject", LogLevelTrace)
		End If
	End If
	If Not IsObject(aJobErrorRequest(O_INBOXXML_JOBERROR)) Then
		lErrNumber = GetXMLDOM(aConnectionInfo, aJobErrorRequest(O_INBOXXML_JOBERROR), sErrDescription)
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "JobErrorCuLib.asp", "ReceiveJobErrorRequest", "", "Error after calling GetXMLDOM", LogLevelTrace)
		End If
	End If
	If lErrNumber <> NO_ERR Then
		ReceiveRebuildRequest = lErrNumber
		Exit Function
	End If
	aObjectInfo(S_OBJECT_ID_OBJECT) = CStr(oRequest("ReportID").Item)
	If Len(aObjectInfo(S_OBJECT_ID_OBJECT)) > 0 Then
		aObjectInfo(L_TYPE_OBJECT) = DssXmlTypeReportDefinition
	Else
		aObjectInfo(S_OBJECT_ID_OBJECT) = CStr(oRequest("DocumentID").Item)
		aObjectInfo(L_TYPE_OBJECT) = DssXmlTypeDocumentDefinition
	End If
	aObjectInfo(S_TARGET_PAGE_OBJECT) = "Folder.asp"
	aFolderInfo(S_TARGET_PAGE_OBJECT) = "Folder.asp"
	aFolderInfo(N_NUMBER_OF_FOLDERS_TO_SHOW_OBJECT) = 0

	If Len(aJobErrorRequest(O_RESULTSET_JOBERROR).MessageID) = 0 Then
		aJobErrorRequest(O_RESULTSET_JOBERROR).MessageID = CStr(oRequest("MsgID").Item)
	End If
	If IsEmpty(aJobErrorRequest(S_MSG_JOBERROR)) Then
		aJobErrorRequest(S_MSG_JOBERROR) = CStr(oRequest("Msg").Item)
	End If
	If IsEmpty(aJobErrorRequest(S_PREVIOUS_PAGE_JOBERROR)) Then
		aJobErrorRequest(S_PREVIOUS_PAGE_JOBERROR) = CStr(oRequest("PreviousPage").Item)
	End If
	If IsEmpty(aJobErrorRequest(S_OK_JOBERROR)) Then
		aJobErrorRequest(S_OK_JOBERROR) = CStr(oRequest("Ok").Item)
	End If
	If IsEmpty(aJobErrorRequest(L_ERROR_NUMBER_JOBERROR)) Then
		aJobErrorRequest(L_ERROR_NUMBER_JOBERROR) = CStr(oRequest("ErrNum").Item)
	End If
	If Len(aJobErrorRequest(L_ERROR_NUMBER_JOBERROR)) > 0 Then
		aJobErrorRequest(L_ERROR_NUMBER_JOBERROR) = CLng(aJobErrorRequest(L_ERROR_NUMBER_JOBERROR))
	Else
		aJobErrorRequest(L_ERROR_NUMBER_JOBERROR) = 0
	End If
	If IsEmpty(aJobErrorRequest(S_ERROR_DESCRIPTION_JOBERROR)) Then
		aJobErrorRequest(S_ERROR_DESCRIPTION_JOBERROR) = CStr(oRequest("ErrDesc").Item)
	End If
	If IsEmpty(aJobErrorRequest(S_REQUEST_JOBERROR)) Then
		aJobErrorRequest(S_REQUEST_JOBERROR) = CStr(oRequest("Request").Item)
	End If
	Err.Clear
End Function

Function GetErrorMessageForMsgID(aJobErrorRequest, aConnectionInfo, sErrorMessage, sErrDescription)
'******************************************************************************
'Purpose: Get the Error Message from Inbox
'Inputs:   aJobErrorRequest, aConnectionInfo
'Outputs: sErrorMessage, lErrNumber, sErrDescription
'******************************************************************************
	On Error Resume Next
	Dim sInboxXML
	Dim oInbox
	Dim lErrNumber

	sErrorMessage = ""
	sErrDescription = ""
	lErrNumber = NO_ERR
	Set oInbox = aJobErrorRequest(O_RESULTSET_JOBERROR).InboxObject
	lErrNumber = Err.number
	If Not IsObject(oInbox) Or lErrNumber <> NO_ERR Then
		sErrDescription = asDescriptors(153) 'Descriptor: Error accessing user's history list.
		Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, Err.source, "JobErrorCuLib.asp", "GetErrorMessageForMsgID", "", "Error accessing user's history list", LogLevelTrace)
	Else
		'*** Get my Inbox ***'
		sInboxXML = oInbox.GetMessages()
		lErrNumber = Err.number
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "JobErrorCuLib.asp", "GetErrorMessageForMsgID", "InboxObject.GetMessages", "Error in call to InboxObject.GetMessages() function", LogLevelError)
			sErrDescription = asDescriptors(154) 'Descriptor: Your history list messages could not be retrieved from MicroStrategy Server.
		Else
			aJobErrorRequest(O_INBOXXML_JOBERROR).loadXML(sInboxXML)
			lErrNumber = Err.number
			If lErrNumber <> NO_ERR Then
				sErrDescription = asDescriptors(271) 'Descriptor: Error loading XML data.
				Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, Err.source, "JobErrorCuLib.asp", "GetErrorMessageForMsgID", "", "Error loading XML", LogLevelError)
			Else
				'*** Parse XML for Job ID ***'
				sErrorMessage = aJobErrorRequest(O_INBOXXML_JOBERROR).SelectSingleNode("/mi/ic/im[@mid='" & oInbox.MessageID & "']").firstChild.text
			End If
		End If
	End If
	Set oInbox = Nothing
	GetErrorMessageForMsgID = lErrNumber
	Err.Clear
End Function

Function CancelMessage(aJobErrorRequest, aConnectionInfo, sErrDescription)
'******************************************************************************
'Purpose: Delete the message from History List
'Inputs:   aJobErrorRequest, aConnectionInfo
'Outputs:  lErrNumber, sErrDescription
'******************************************************************************
	On Error Resume Next
	Dim oInbox
	Dim lErrNumber
	lErrNumber = NO_ERR
	Set oInbox = aJobErrorRequest(O_RESULTSET_JOBERROR).InboxObject
	lErrNumber = Err.number
	If Not IsObject(oInbox) Or lErrNumber <> NO_ERR Then
		sErrDescription = asDescriptors(153) 'Descriptor: Error accessing user's history list.
		Call LogErrorXML(aConnectionInfo, lErrNumber, Err.description, Err.source, "JobErrorCuLib.asp", "CancelMessage", "", "Error accessing user's history list", LogLevelTrace)
	Else
		Call oInbox.Remove()
		lErrNumber = Err.number
		If lErrNumber <> NO_ERR Then
			sErrDescription = asDescriptors(155) 'Descriptor: Your history list messages could not be removed from MicroStrategy Server.
			Call LogErrorXML(aConnectionInfo, Err.number, Err.description, Err.source, "JobErrorCuLib.asp", "CancelMessage()", "InboxObject.Remove", "Error in call InboxObject.Remove()", LogLevelError)
		End If
	End If
	CancelMessage = lErrNumber
	Set oInbox = Nothing
	Err.Clear
End Function
%>