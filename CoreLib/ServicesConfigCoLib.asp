<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%

Function co_getObjectsParentInfo(sSiteId, aObjectsId, sObjectsXML)
'********************************************************
'*Purpose: Return all objects of the MD which match the given type
'*Inputs:  sSiteId: A valid siteId for the MD; nObjectType: the type of object to search
'*Outputs: sObjectsForSiteXML: An XML string with the list of objects found.
'********************************************************
Const PROCEDURE_NAME = "co_getObjectsParentInfo"
Dim lErr
Dim sErr

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call GetObjectsForSite
    If lErr = NO_ERR Then
        sObjectsXML = oSiteInfo.getObjectParentInfo(sSiteId, aObjectsId)
        'sObjectsXML = "<mi><as>"
        'sObjectsXML = sObjectsXML &  "<a tp='2'  id='56D4248694B211D4BE6600C04F0E93B7' n='Services' ct='01/31/2001 01:00:00 AM' mdt='01/31/2003 01:00:00 AM' />"
        'sObjectsXML = sObjectsXML &  "<a tp='19' id='9D1F4D36FC7711D48D96009027DCD594' n='fportalService'  ct='01/31/2001 01:00:00 AM' mdt='01/31/2003 01:00:00 AM' />"
        'sObjectsXML = sObjectsXML &  "<a tp='19' id='66C269F4F6C511D48D94009027DCD594' n='salesService'  ct='01/31/2001 01:00:00 AM' mdt='01/31/2003 01:00:00 AM' />"
        'sObjectsXML = sObjectsXML &  "</as>"
        'sObjectsXML = sObjectsXML &  "<in>"
        'sObjectsXML = sObjectsXML &  "<oi tp='19' id='9D1F4D36FC7711D48D96009027DCD594' pid='56D4248694B211D4BE6600C04F0E93B7' />"
        'sObjectsXML = sObjectsXML &  "<oi tp='19' id='66C269F4F6C511D48D94009027DCD594' pid='56D4248694B211D4BE6600C04F0E93B7' />"
        'sObjectsXML = sObjectsXML &  "</in></mi>"

        lErr = checkReturnValue(sObjectsXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "getObjectsForSite", "Error calling getObjectsForSite", LogLevelError)
    End If


    Set oSiteInfo = Nothing

    co_getObjectsParentInfo =  lErr
    Err.Clear

End Function

Function co_createObject(sSiteId, sParentId, sObjectId, sPropertiesXML)
'********************************************************
'*Purpose: Creates a MD object
'*Inputs:  sSiteId: A valid siteId for the MD;
'           sParentId: The parent of this object, if the parent is the root, use the siteId
'           sObjectId: The GUID of the object (Note: Objects must have unique Ids across different sites of the same MD
'           sPropertiesXML: The properties of the object.
'*Outputs: none
'********************************************************
Const PROCEDURE_NAME = "co_createObject"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call GetObjectsForSite
    If lErr = NO_ERR Then
        sReturn = oSiteInfo.createObject(sSiteId, sParentId, sObjectId, sPropertiesXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "createObject", "Error calling createObject", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_createObject =  lErr
    Err.Clear

End Function

Function co_createMappingDefinition(sSiteId, sId, sName, sDefinitionXML)
'********************************************************
'*Purpose: Creates a Mapping definition
'*Inputs:  sSiteId: A valid siteId for the MD;
'          sMapId: the parent to which this object is the definition
'          sName: Map name
'          sDefinitionXML: The defintion
'*Outputs: none
'********************************************************
Const PROCEDURE_NAME = "co_createMappingDefinition"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call GetObjectsForSite
    If lErr = NO_ERR Then
        sReturn = oSiteInfo.createMappingDefinition(sSiteId, sId, sName, sDefinitionXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "createMappingDefinition", "Error calling createMappingDefinition", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_createMappingDefinition =  lErr
    Err.Clear

End Function

Function co_deleteMappingDefinition(sSiteId, sObjectId)
'********************************************************
'*Purpose: Delete a mapping definition from MD
'*Inputs:  sSiteId: A valid siteId for the MD;
'           sObjectId: The GUID of the map to delete
'*Outputs: none
'********************************************************
Const PROCEDURE_NAME = "co_deleteMappingDefinition"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call deleteObject
    If lErr = NO_ERR Then
        sReturn = oSiteInfo.deleteMappingDefinition(sSiteId, sObjectId)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "updateObjectProperties", "Error calling updateObjectProperties", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_deleteMappingDefinition =  lErr
    Err.Clear

End Function

Function co_deleteObject(sSiteId, sObjectId)
'********************************************************
'*Purpose: Delete an object from MD
'*Inputs:  sSiteId: A valid siteId for the MD;
'           sObjectId: The GUID of the object to delete
'*Outputs: none
'********************************************************
Const PROCEDURE_NAME = "co_deleteObject"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call deleteObject
    If lErr = NO_ERR Then
        sReturn = oSiteInfo.deleteObject(sSiteId, sObjectId)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "updateObjectProperties", "Error calling updateObjectProperties", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_deleteObject =  lErr
    Err.Clear

End Function

Function co_updateObjectProperties(sSiteId, sObjectId, sPropertiesXML)
'********************************************************
'*Purpose: Updates a MD object
'*Inputs:  sSiteId: A valid siteId for the MD;
'           sObjectId: The GUID of the object
'           sPropertiesXML: The properties of the object.
'*Outputs: none
'********************************************************
Const PROCEDURE_NAME = "co_updateObjectProperties"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call updateObjectProperties
    If lErr = NO_ERR Then
        sReturn = oSiteInfo.updateObjectProperties(sSiteId, sObjectId, sPropertiesXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "updateObjectProperties", "Error calling updateObjectProperties", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_updateObjectProperties =  lErr
    Err.Clear

End Function

Function co_getSubscriptionSetsForService(sSiteId, sServiceId, sSubsSetsXML)
'********************************************************
'*Purpose: Gets the subscrition sets of a given service
'*Inputs:  sSiteId: A valid siteId for the MD;
'           sServiceId: The service to retrieve the subscription sets (from Project Repository)
'*Outputs: sSubsSetsXML: The XML returned by backend
'********************************************************
Const PROCEDURE_NAME = "co_getSubscriptionSetsForService"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call updateObjectProperties
    If lErr = NO_ERR Then
        sSubsSetsXML = oSiteInfo.getSubscriptionSetsForService(sSiteId, sServiceId)
        'sSubsSetsXML = oSiteInfo.getNamedSchedulesForService(sSiteId, sServiceId, False)
        lErr = checkReturnValue(sSubsSetsXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "getSubscriptionSetsForService", "Error calling getSubscriptionSetsForService", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_getSubscriptionSetsForService =  lErr
    Err.Clear

End Function

Function co_getQuestionsForService(sSiteId, sServiceId, sQuestionsXML)
'********************************************************
'*Purpose: Gets the Questions of a given service
'*Inputs:  sSiteId: A valid siteId for the MD;
'          sServiceId: The service to retrieve the Questions (from Project Repository)
'*Outputs: sQuestionsXML: The XML returned by backend
'********************************************************
Const PROCEDURE_NAME = "co_getQuestionsForService"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call updateObjectProperties
    If lErr = NO_ERR Then
        sQuestionsXML = oSiteInfo.getQuestionsForService(sSiteId, sServiceId)
        lErr = checkReturnValue(sQuestionsXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "getQuestionsForService", "Error calling getQuestionsForService", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_getQuestionsForService =  lErr
    Err.Clear

End Function


Function co_getMappingDefinition(sSiteId, sMapId, sMapXML)
'********************************************************
'*Purpose: Gets the Questions of a given service
'*Inputs:  sSiteId: A valid siteId for the MD;
'          sServiceId: The service to retrieve the Questions (from Project Repository)
'*Outputs: sQuestionsXML: The XML returned by backend
'********************************************************
Const PROCEDURE_NAME = "co_getMappingDefinition"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call updateObjectProperties
    If lErr = NO_ERR Then
        sMapXML = oSiteInfo.getMappingDefinition(sSiteId, sMapId)
        'sMapXML = Left(sMapXML, Len(sMapXML) - Len("</pr></prs></oi></in></mi>"))
        lErr = checkReturnValue(sMapXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "getMappingDefinition calling getMappingDefinition", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_getMappingDefinition =  lErr
    Err.Clear

End Function


Function co_getMappingObjects(sSiteId, sObjectId, sMapsXML)
'********************************************************
'*Purpose: Gets the Questions of a given service
'*Inputs:  sSiteId: A valid siteId for the MD;
'          sServiceId: The service to retrieve the Questions (from Project Repository)
'*Outputs: sQuestionsXML: The XML returned by backend
'********************************************************
Const PROCEDURE_NAME = "co_getMappingObjects"
Dim lErr
Dim sErr
Dim sReturn

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call updateObjectProperties
    If lErr = NO_ERR Then
        sMapsXML = oSiteInfo.getMappingObjects(sSiteId, sObjectId, TYPE_QUESTION_CONFIG)
        lErr = checkReturnValue(sMapsXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "getMappingObjects", "Error calling getMappingObjects", LogLevelError)
    End If

    Set oSiteInfo = Nothing

    co_getMappingObjects =  lErr
    Err.Clear

End Function

Function co_getColumns(sDBAlias, sOwner, sTableName, sColumnsXML)
'********************************************************
'*Purpose: Returns the columsn of a given table
'*Inputs:  sDBAlias: The dbalias where the tables are.
'          sTableName: The tables of which we want the columns
'*Outputs: sColumnsXML: The returned XML
'********************************************************
Const PROCEDURE_NAME = "co_getColumns"
Dim lErr
Dim sErr
Dim sReturn

Dim oAdmin

    On Error Resume Next
    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)

    '=--Call updateObjectProperties
    If lErr = NO_ERR Then
        sColumnsXML = oAdmin.getColumns(sDBAlias, "", sOwner, sTableName, "")
        lErr = checkReturnValue(sColumnsXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "getColumns", "Error calling getColumns", LogLevelError)
    End If

    Set oAdmin = Nothing

    co_getColumns =  lErr
    Err.Clear

End Function

Function co_getTables(sDBAlias, sOwner, sFilter, sTablesXML)
'********************************************************
'*Purpose: Returns the Tables of a given Database
'*Inputs:  sDBAlias: The dbalias where the tables are.
'          sFilter: A valid SQL expression to restrict the search
'*Outputs: sTablesXML: The returned XML
'********************************************************
Const PROCEDURE_NAME = "co_getTables"
Dim lErr
Dim sErr
Dim sReturn


Dim oAdmin

    On Error Resume Next
    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)

    '=--Call updateObjectProperties
    If lErr = NO_ERR Then
        sTablesXML = oAdmin.getTables(sDBAlias, "", sOwner, sFilter)
        lErr = checkReturnValue(sTablesXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCoLib.asp", PROCEDURE_NAME, "getTables", "Error calling getTables", LogLevelError)
    End If

    Set oAdmin = Nothing

    co_getTables =  lErr
    Err.Clear

End Function


Function co_GetDetailsForQuestions(sSiteId, asQuestionObjectID, sGetDetailsForQuestionsXML)
'********************************************************
'*Purpose: Given a asQuestionObjectID, returns details of that question
'*Inputs: sSessionID, asQuestionObjectID
'*Outputs: sGetDetailsForQuestionsXML
'********************************************************
Const PROCEDURE_NAME = "co_GetDetailsForQuestions"
Dim oSiteInfo
Dim lErrNumber
Dim sErr

    On Error Resume Next
	lErrNumber = NO_ERR

	Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
	If Err.number <> NO_ERR Then
        lErrNumber = Err.number
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_PERSONALIZATION_INFO, LogLevelError)
    Else
        sGetDetailsForQuestionsXML = oSiteInfo.getDetailsForQuestions(sSiteId, asQuestionObjectID)
        lErrNumber = checkReturnValue(sGetDetailsForQuestionsXML, sErr)
        If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), sErr, CStr(Err.source), "PrePromptCoLib.asp", PROCEDURE_NAME, "PersonalizationInfo.getDetailsForQuestions", "Error while calling getDetailsForQuestions", LogLevelError)
	End If

	Set oSiteInfo = Nothing

	co_GetDetailsForQuestions = lErrNumber
	Err.Clear

End Function


%>
