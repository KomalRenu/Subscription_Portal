<%

Function co_GetProfileNamesForQuestions(sSessionID, aQuestionIDs, pDef)
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetDetailsForQuestions"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr
	Dim tempArray()
	Dim i
	
	lErrNumber = NO_ERR
		
	Redim tempArray(Ubound(aQuestionIDs))
	
	For i=0 to ubound(aQuestionIDs)
		tempArray(i) = aQuestionIDs(i)
	Next
	
	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        pDef = oPersonalizationInfo.getProfileNamesForQuestions(sSessionID, tempArray)
        lErrNumber = checkReturnValue(pDef, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.getDetailsForQuestions", "Error while calling getDetailsForQuestions", LogLevelError)
        End If
	End If
		
	Set oPersonalizationInfo = Nothing
	
	co_GetProfileNamesForQuestions = lErrNumber
	Err.Clear

End Function

Function co_SavePreferenceObject(sessionID,sPreferenceObjectID,ansXML)
    On Error Resume Next
    Const PROCEDURE_NAME = "co_SavePreferenceObject"
	Dim oPersonalizationInfo
	Dim lErrNumber
	Dim sErr
	Dim tempArray1(0)
	Dim tempArray2(0)
	
	lErrNumber = NO_ERR
		
	tempArray1(0) = sPreferenceObjectID
	tempArray2(0) = ansXML

	Set oPersonalizationInfo = Server.CreateObject(PROGID_PERSONALIZATION_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        pDef = oPersonalizationInfo.savePreferenceObjects(sessionID,tempArray1,tempArray2)
        lErrNumber = checkReturnValue(pDef, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.getDetailsForQuestions", "Error while calling getDetailsForQuestions", LogLevelError)
        End If
	End If
		
	Set oPersonalizationInfo = Nothing
	
	co_GetProfileNamesForQuestions = lErrNumber
	Err.Clear
End Function


Function co_GetInformationSourceDefinition(sessionID,ISID,ISDefn)
    On Error Resume Next
    Const PROCEDURE_NAME = "co_GetInformationSourceDefinition"
	Dim oSystemInfo
	Dim lErrNumber
	Dim sErr
	
	lErrNumber = NO_ERR
		
	Set oSystemInfo = Server.CreateObject(PROGID_SYSTEM_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        ISDefn = oSystemInfo.getInformationSourceDefinition(sessionID,ISID)
        lErrNumber = checkReturnValue(ISDefn, sErr)
        If lErrNumber <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.getDetailsForQuestions", "Error while calling getDetailsForQuestions", LogLevelError)
        End If
	End If
		
	Set oSystemInfo = Nothing
	
	co_GetInformationSourceDefinition = lErrNumber
	Err.Clear
End Function
%>