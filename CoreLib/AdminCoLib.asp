<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
'Root folder types
Const ROOT_APP_FOLDER_TYPE = 3
Const ROOT_DEVICE_FOLDER_TYPE = 14

Function co_getObjectsForSite(sSiteId, nObjectType, sObjectsForSiteXML)
'********************************************************
'*Purpose: Return all objects of the MD which match the given type
'*Inputs:  sSiteId: A valid siteId for the MD; nObjectType: the type of object to search
'*Outputs: sObjectsForSiteXML: An XML string with the list of objects found.
'********************************************************
Const PROCEDURE_NAME = "co_getObjectsForSite"
Dim lErr
Dim sErr

Dim oSiteInfo

    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call GetObjectsForSite
    If lErr = NO_ERR Then
        sObjectsForSiteXML = oSiteInfo.getObjectsForSite(sSiteId, nObjectType)
        'sObjectsForSiteXML = "<mi><in>"
 	      'sObjectsForSiteXML = sObjectsForSiteXML & "<oi tp='1017' id='sGUID1' n='' des='' phid='9D1F4D36FC7711D48D96009027DCD594' />"
 	      'sObjectsForSiteXML = sObjectsForSiteXML & "<oi tp='1017' id='sGUID4' n='' des='' phid='66C269F4F6C511D48D94009027DCD594' />"
 	      'sObjectsForSiteXML = sObjectsForSiteXML & "</in></mi>"

        lErr = checkReturnValue(sObjectsForSiteXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "AdminCoLib.asp", PROCEDURE_NAME, "getObjectsForSite", "Error calling getObjectsForSite", LogLevelError)
    End If


    Set oSiteInfo = Nothing

    co_getObjectsForSite =  lErr
    Err.Clear

End Function

Function co_getObjectProperties(sSiteId, sObjectId, sObjectPropsXML)
'********************************************************
'*Purpose: Return the object properties of an object.
'*Inputs:  sSiteId: A valid siteId for the MD; sObjectId: the object Id to retrieve properties from
'*Outputs: sObjectPropsXML: An XML string with the properties of the object.
'********************************************************
Const PROCEDURE_NAME = "co_getObjectProperties"
Dim lErr
Dim sErr

Dim oSiteInfo


    On Error Resume Next
    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)

    '=--Call GetObjectsForSite
    If lErr = NO_ERR Then
        sObjectPropsXML = oSiteInfo.getObjectProperties(sSiteId, sObjectId)
        'sObjectPropsXML = ""
        'sObjectPropsXML = sObjectPropsXML + "<mi><in><oi tp='1017' id='sGUID1'>"
        'sObjectPropsXML = sObjectPropsXML +   "<prs>"
        'sObjectPropsXML = sObjectPropsXML +     "<pr id='DEFAULT_SLICING_ANSWER' v='" & ANSWER_USER_ID & "' />"
        'sObjectPropsXML = sObjectPropsXML +     "<pr id='PHYSICAL_ID' v='9D1F4D36FC7711D48D96009027DCD594' />"
        'sObjectPropsXML = sObjectPropsXML +     "<pr id='NAME' v='Sales Service with personalization using an object prompt' />"
        'sObjectPropsXML = sObjectPropsXML +   "</prs>"
        'sObjectPropsXML = sObjectPropsXML +   "<mi>"
        'sObjectPropsXML = sObjectPropsXML +     "<in>"
        'sObjectPropsXML = sObjectPropsXML +       "<oi tp='1013' id='sGUID2'>"
        'sObjectPropsXML = sObjectPropsXML +         "<prs>"
        'sObjectPropsXML = sObjectPropsXML +           "<pr id='PHYSICAL_ID' v='EADC7003B1A211D48FBE00C04F58EBC8' />"
        'sObjectPropsXML = sObjectPropsXML +           "<pr id='NAME' v='somename' />"
        'sObjectPropsXML = sObjectPropsXML +         "</prs>"
        'sObjectPropsXML = sObjectPropsXML +			    "<mi>"
        'sObjectPropsXML = sObjectPropsXML +     	    "<in>"
        'sObjectPropsXML = sObjectPropsXML +			        "<oi tp='1011' id='sGUID3'>"
        'sObjectPropsXML = sObjectPropsXML +			          "<prs>"
        'sObjectPropsXML = sObjectPropsXML +			            "<pr id='PHYSICAL_ID' v='FADC7003B1A211D48FBE00C04F58EBC8' />"
        'sObjectPropsXML = sObjectPropsXML +			            "<pr id='NAME' v='Object prompt QO' />"
        'sObjectPropsXML = sObjectPropsXML +					        "<pr id='QUESTION_TYPE' v='0' />"
        'sObjectPropsXML = sObjectPropsXML +			            "<pr id='IS_SLICING' v='NO' />"
        'sObjectPropsXML = sObjectPropsXML +			            "<pr id='IS_SHOWN' v='YES' />"
        'sObjectPropsXML = sObjectPropsXML +			          "</prs>"
        'sObjectPropsXML = sObjectPropsXML +			 	      "</oi>"
        'sObjectPropsXML = sObjectPropsXML +			      "</in>"
        'sObjectPropsXML = sObjectPropsXML +			    "</mi>"
        'sObjectPropsXML = sObjectPropsXML + 	    "</oi>"
        'sObjectPropsXML = sObjectPropsXML +     "</in>"
        'sObjectPropsXML = sObjectPropsXML +   "</mi>"
        'sObjectPropsXML = sObjectPropsXML + "</oi></in></mi>"

        lErr = checkReturnValue(sObjectPropsXML, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "AdminCoLib.asp", PROCEDURE_NAME, "getObjectProperties", "Error calling getObjectProperties", LogLevelError)

    End If


    Set oSiteInfo = Nothing

    co_getObjectProperties =  lErr
    Err.Clear

End Function

Function co_ResetSubscriptionEngine()
'********************************************************
'*Purpose:  Resets the Subscription Engine so values take place
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_ResetSubscriptionEngine"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "AdminCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn = oAdmin.resetSubscriptionEngines()
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "AdminCoLib.asp", PROCEDURE_NAME, "resetSubscriptionEngines", "Error calling oAdmin.resetSubscriptionEngines", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_ResetSubscriptionEngine = lErr
    Err.Clear
End Function

Function co_ConnectToSubscriptionEngines()
'********************************************************
'*Purpose:  Resets the Subscription Engine so values take place
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_ConnectToSubscriptionEngine"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "AdminCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn = oAdmin.connectToSubscriptionEngines()
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "AdminCoLib.asp", PROCEDURE_NAME, "connectToSubscriptionEngines", "Error calling oAdmin.connectToSubscriptionEngines", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_ResetSubscriptionEngine = lErr
    Err.Clear
End Function


Function co_generateSubscriptionSetSQL(sSiteId, sServiceId, sSubsSetId)
'********************************************************
'*Purpose: Generates the necessary SQL for a given subscription set
'*Inputs:  sSiteId: A valid siteId for the MD;
'          sSubsSetId: a valid repository object Id
'*Outputs: None
'********************************************************
Const PROCEDURE_NAME = "co_generateSubscriptionSetSQL"
Dim lErr
Dim sErr
Dim sReturn

Dim oAdmin

    On Error Resume Next
    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)

    '=--Call updateObjectProperties
    If lErr = NO_ERR Then
        sReturn = oAdmin.generateSubscriptionSetSQL(sSiteId, sServiceId, sSubsSetId, True)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "AdminCoLib.asp", PROCEDURE_NAME, "generateSubscriptionSetSQL", "Error calling generateSubscriptionSetSQL", LogLevelError)
    End If

    Set oAdmin = Nothing

    co_generateSubscriptionSetSQL =  lErr
    Err.Clear

End Function

Function co_getFolderXML(sFolderId, nRootType, sFolderXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Const PROCEDURE_NAME = "co_getFolderXML"
    Dim lErr
    Dim sErr
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), Err.source, "AdminCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sFolderXML = oSiteInfo.getFolderContents(Application.Value("SITE_ID"), sFolderId, nRootType)
        lErr = checkReturnValue(sFolderXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "AdminCoLib.asp", PROCEDURE_NAME, "getFolderContents", "Error calling getFolderContents", LogLevelError)
        End If
    End If

    If sFolderXML = "<mi><as/><fct/></mi>" Then lErr = ERR_INACTIVE_FOLDER_ANCESTOR

    Set oSiteInfo = Nothing

    co_GetFolderXML = lErr
    Err.Clear
End Function

%>