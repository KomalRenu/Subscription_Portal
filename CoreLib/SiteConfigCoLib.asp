<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%

'Number of engines to keep track of:
Const MRU_COUNT = 5

Function CO_deleteDBAlias(aDBAlias)
'********************************************************
'*Purpose:  Wrapper function of SDK to createDBAlias which creates a DB Alias in the Subscription Engine
'*Inputs:   aDBAlias (string array)
'*Outputs:  Error Code
'********************************************************
CONST PROCEDURE_NAME = "CO_deleteDBAlias"
Dim lErr
Dim sErr

Dim oAdmin
Dim sResultXML

    On Error Resume Next
    lErr = NO_ERR

    If lErr = NO_ERR Then
        Set oAdmin = Server.CreateObject(PROGID_ADMIN)

        sResultXML = oAdmin.deleteDBAlias(CStr(aDBAlias(0)))
        lErr = checkReturnValue(sResultXML, sErr)

        If lErr <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "CO_deleteDBAlias", sErr, LogLevelError)
		End If

    End If

    Set oAdmin = Nothing

    CO_deleteDBAlias = lErr
    Err.Clear

End Function

Function co_createMDTables()
'********************************************************
'*Purpose:  Calls createMDTables to create the tables in the MD
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_createMDTables"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating" & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturn = oSiteInfo.createMetadataTables("admin")
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "createMetadataTables", "Error calling createMetadataTables", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_createMDTables = lErr
    Err.Clear
End Function

Function co_CheckMDTables(sReturnXML)
'********************************************************
'*Purpose:  Calls CheckMDTables to see if the tables are ready.
'*Inputs:
'*Outputs:  sReturnXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_CheckMDTables"
    Dim lErr
    Dim sErr
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturnXML = oSiteInfo.checkMetadataTables("admin")
        lErr = checkReturnValue(sReturnXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "checkMetadataTables", "Error calling oSiteInfo.checkMetadataTables", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_CheckMDTables = lErr
    Err.Clear
End Function

Function co_editDBAlias(sAliasName,sAlias)
'********************************************************
'*Purpose:  Wrapper function of SDK to edit DBAlias which creates a DB Alias in the Subscription Engine
'*Inputs:   sAlias (string array)
'			sAliasName
'*Outputs:  Error Code
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_editDBAlias"
    Dim lErr
    Dim sErr
    Dim oAdmin
    Dim oSLAPI
    Dim sResultXML

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
		Set oSLAPI = Server.CreateObject("MSTRSequeLAPI.cSequeLAPI")
		If Err.number <> NO_ERR Then
		    lErr = Err.number
		    Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & MSTRSequeLAPI.cSequeLAPI, LogLevelError)
		Else
			lErr = oSLAPI.SetDBAlias(CStr(sAlias),nothing)
			If lErr <> NO_ERR Then
				Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "createDBAlias", "Error calling SetDBAlias", LogLevelError)
			Else
				lErr = oSLAPI.ValidateDataSource(nothing)
				If lErr <> NO_ERR Then
					Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "createDBAlias", "Error calling ValidateDataSource", LogLevelError)
				Else
					sResultXML = oAdmin.updateDBAlias(sAliasName,CStr(sAlias))
					lErr = checkReturnValue(sResultXML, sErr)
					If lErr <> NO_ERR Then
						Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "editDBAlias", "Error calling updateDBAlias", LogLevelError)
					End If
				End if
			End If
	    End if
    End If

    Set oAdmin = Nothing

    co_editDBAlias = lErr
    Err.Clear
End Function

Function co_createDBAlias(sAlias)
'********************************************************
'*Purpose:  Wrapper function of SDK to createDBAlias which creates a DB Alias in the Subscription Engine
'*Inputs:   aDBAlias (string array)
'*Outputs:  Error Code
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "CO_createDBAlias"
    Dim lErr
    Dim sErr
    Dim oAdmin
    Dim oSLAPI
    Dim sResultXML

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
		Set oSLAPI = Server.CreateObject("MSTRSequeLAPI.cSequeLAPI")
		If Err.number <> NO_ERR Then
		    lErr = Err.number
		    Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & MSTRSequeLAPI.cSequeLAPI, LogLevelError)
		Else
			lErr = oSLAPI.SetDBAlias(CStr(sAlias),nothing)
			If lErr <> NO_ERR Then
				Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "createDBAlias", "Error calling SetDBAlias", LogLevelError)
			Else
				lErr = oSLAPI.ValidateDataSource(nothing)
				If lErr <> NO_ERR Then
					Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "createDBAlias", "Error calling ValidateDataSource", LogLevelError)
				Else
					sResultXML = oAdmin.createDBAlias(CStr(sAlias))
					lErr = checkReturnValue(sResultXML, sErr)
					If lErr <> NO_ERR Then
						Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "createDBAlias", "Error calling createDBAlias", LogLevelError)
					End If
				End if
			End If
	    End if
    End If

    Set oAdmin = Nothing
    Set oSLAPI = Nothing

    CO_createDBAlias = lErr
    Err.Clear
End Function

Function co_GetDBAliases(sAliasesXML)
'********************************************************
'*Purpose:  Returns a list of Db Aliases found in the Subscription Engine
'*Inputs:
'*Outputs:  sAliasesXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_GetDBAliases"
    Dim lErr
    Dim sErr
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sAliasesXML = oAdmin.getDBAliases()
        lErr = checkReturnValue(sAliasesXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "getDBAliases", "Error calling oAdmin.getDBAliases", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_GetDBAliases = lErr
    Err.Clear
End Function

Function co_checkDBAlias(sDBAlias, sPrefix, lRepositoryType)
'********************************************************
'*Purpose:  Wrapper function of SDK to checkDBTables which checks that tables are created
'			in Database and have the correct version
'*Inputs:   sDBAlias, sPrefix, lRepositoryType
'*Outputs:  Error Code
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "CO_checkDBAlias"
    Dim lErr
    Dim sErr
    Dim oAdmin
    Dim sAlias
    Dim sResultXML

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sResultXML = oAdmin.checkDBAlias(sDBAlias, sPrefix, lRepositoryType)
        lErr = checkReturnValue(sResultXML, sErr)
        If lErr <> NO_ERR Then
	    	Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "checkDBAlias", "Error calling checkDBAlias", LogLevelError)
	    End If
    End If

    Set oAdmin = Nothing

    CO_checkDBAlias = lErr
    Err.Clear
End Function

Function co_GetAllSites(sSitesXML)
'********************************************************
'*Purpose:  Returns a list of all the sites stored in the current MD,
'           including id, name and description
'*Inputs:
'*Outputs:  sSitesXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_GetAllSites"
    Dim lErr
    Dim sErr
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sSitesXML = oSiteInfo.getAllSites()
        lErr = checkReturnValue(sSitesXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "getAllSites", "Error calling oSiteInfo.getAllSites", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_GetAllSites = lErr
    Err.Clear
End Function

Function co_DeleteSite(sSiteId)
'********************************************************
'*Purpose:  Deletes the given siteId.
'*Inputs:   sName, sDescription> figure it out ;)
'*Outputs:  sSiteId: The Id of the site we just created.
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_DeleteSite"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturn = oSiteInfo.deleteSite(sSiteId)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "deleteSite", "Error calling oSiteInfo.deleteSite", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_DeleteSite = lErr
    Err.Clear
End Function

Function co_CreateSite(sSiteID, sConfigXML)
'********************************************************
'*Purpose:  Creates a new Site on the Hydra MD, and by default
'           makes it the selected one.
'*Inputs:   sName, sDescription> figure it out ;)
'*Outputs:  sSiteId: The Id of the site we just created.
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_CreateSite"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturn = oSiteInfo.createSite(sSiteID, sConfigXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "createSite", "Error calling oSiteInfo.createSite", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_CreateSite = lErr
    Err.Clear
End Function


Function co_createConnection(sSiteID, sParentID, sConnectionID, sPropertiesXML)
'********************************************************
'*Purpose:  Creates a connection object
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_createConnection"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturn = oSiteInfo.createObject(sSiteID, sParentID, sConnectionID, sPropertiesXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "createSite", "Error calling oSiteInfo.createConnection", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_createConnection = lErr
    Err.Clear

End Function

Function co_SetSite(sConfigXML)
  '********************************************************
'*Purpose:  Selects a site from the Hydra MD, to be used by this Portal.
'*Inputs:   sSiteId: SiteID
'           aSiteProperties: The site properties, name and description should have valid values
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_SetSite"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn =  oAdmin.saveSiteConfigurationProperties(sConfigXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "saveSiteConfigurationProperties", "Error calling saveSiteConfigurationProperties", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_SetSite = lErr
    Err.Clear
End Function

Function co_GetMDConn(sMDXML)
'********************************************************
'*Purpose:  Returns the values of the current MDConn as used by the engine
'*Inputs:
'*Outputs:  sMDXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_GetMDConn"
    Dim lErr
    Dim sErr
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sMDXML = oAdmin.getMetadataConnectionProperties()
        lErr = checkReturnValue(sMDXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "getMetadataConnectionProperties", "Error calling oAdmin.getMetadataConnectionProperties", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_GetMDConn = lErr
    Err.Clear
End Function

Function co_SetMDConn(sConfigXML)
'********************************************************
'*Purpose:  sets the value of the MDConnection. This call invalidates
'           the local siteId and DBConn. The value is stored on the subscriptionAPI
'           file, but is not updated in the engine until a new site
'           is selected
'*Inputs:   MDDBAlias: the DBConnection the SE will use for the MD
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_SetMDConn"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn = oAdmin.saveMetadataConnectionProperties(sConfigXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "saveMetadataConnectionProperties", "Error calling oAdmin.saveMetadataConnectionProperties", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_SetMDConn = lErr
    Err.Clear
End Function

Function co_GetMRUEngines(sConfigName, sConfigXML)
'********************************************************
'*Purpose:  Returns a list of the MRU SubscriptinEngines of this portal
'*Inputs:   sConfigName
'*Outputs:  sConfigXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_GetMRUEngines"
    Dim lErr
    Dim sErr
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sConfigXML = oAdmin.getSiteConfigurationProperties(sConfigName)
        lErr = checkReturnValue(sConfigXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "getSiteConfigurationProperties", "Error calling oAdmin.getSiteConfigurationProperties", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_GetMRUEngines = lErr
    Err.Clear
End Function

Function co_SetMRUEngines(sConfigXML)
'********************************************************
'*Purpose:  Selects a site from the Hydra MD, to be used by this Portal.
'*Inputs:
'*Outputs:  sConfigXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_SetMRUEngines"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn =  oAdmin.saveSiteConfigurationProperties(sConfigXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "saveSiteConfigurationByName", "Error calling oAdmin.saveSiteConfigurationProperties", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_SetMRUEngines = lErr
    Err.Clear
End Function

Function co_GetLocales(sSiteID, sLocalesXML)
'********************************************************
'*Purpose:  Returns a list of Locales from the MD
'*Inputs: sSiteID
'*Outputs:  aLocales (string array)
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_GetLocales"
    Dim lErr
    Dim sErr
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sLocalesXML = oSiteInfo.getAllLocales(sSiteID)
        lErr = checkReturnValue(sLocalesXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "getAllLocales", "Error calling oSiteInfo.getAllLocales", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_GetLocales = lErr
    Err.Clear
End Function

Function co_GetSubscriptionEngine(sReturnXML)
'********************************************************
'*Purpose:  Returns the current value of the Subscription Engine
'*Inputs:
'*Outputs:  sReturnXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_GetSubscriptionEngine"
    Dim lErr
    Dim sErr
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturnXML = oAdmin.getSubscriptionEngineLocation()
        lErr = checkReturnValue(sReturnXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "getSubscriptionEngineLocation", "Error calling oAdmin.getSubscriptionEngineLocation", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_GetSubscriptionEngine = lErr
    Err.Clear
End Function

Function co_SetSubscriptionEngine(sNewEngine)
'********************************************************
'*Purpose:  Sets the new location of the SE.
'           It invalidates the value of the MD, Site and DBConnections.
'*Inputs:   sNewEngine: New location
'*Outputs:
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_SetSubscriptionEngine"
    Dim lErr
    Dim sErr
    Dim oAdmin
    Dim sReturn

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturn = oAdmin.saveSubscriptionEngineLocation(sNewEngine)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "saveSubscriptionEngineLocation", "Error calling oAdmin.saveSubscriptionEngineLocation", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_SetSubscriptionEngine = lErr
    Err.Clear
End Function

Function co_CreateDefaultChannel(sSiteId, sChannelId, sChannelXML)
'********************************************************
'*Purpose:  Reads the defaultSiteProperties.xml file and creates into
'           MD all the channels found in it.
'*Inputs:   sSiteId:
'*Outputs:  None.
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_CreateDefaultChannel"
    Dim lErr
    Dim sErr
    Dim sReturn
    Dim oSiteInfo

    lErr = NO_ERR

    Set oSiteInfo = Server.CreateObject(PROGID_SITE_INFO)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_SITE_INFO, LogLevelError)
    Else
        sReturn = oSiteInfo.createChannel(sSiteId, sChannelId, sChannelXML)
        lErr = checkReturnValue(sReturn, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "createChannel", "Error calling oSiteInfo.createChannel", LogLevelError)
        End If
    End If

    Set oSiteInfo = Nothing

    co_CreateDefaultChannel = lErr
    Err.Clear
End Function

Function co_FindBackupPropertyFiles(sReturnXML)
'********************************************************
'*Purpose:  Returns the current value of the Subscription Engine
'*Inputs:
'*Outputs:  sReturnXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_FindBackupPropertyFiles"
    Dim lErr
    Dim sErr
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturnXML = oAdmin.findBackupPropertyFiles()
        lErr = checkReturnValue(sReturnXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "findBackupPropertyFiles", "Error calling oAdmin.findBackupPropertyFiles", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_FindBackupPropertyFiles = lErr
    Err.Clear
End Function


Function co_RestoreBackupPropertyFiles()
'********************************************************
'*Purpose:  Returns the current value of the Subscription Engine
'*Inputs:
'*Outputs:  sReturnXML
'********************************************************
    On Error Resume Next
    CONST PROCEDURE_NAME = "co_RestoreBackupPropertyFiles"
    Dim sReturnXML
    Dim lErr
    Dim sErr
    Dim oAdmin

    lErr = NO_ERR

    Set oAdmin = Server.CreateObject(PROGID_ADMIN)
    If Err.number <> NO_ERR Then
        lErr = Err.number
        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "", "Error creating " & PROGID_ADMIN, LogLevelError)
    Else
        sReturnXML = oAdmin.restoreBackupPropertyFiles()
        lErr = checkReturnValue(sReturnXML, sErr)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, CStr(Err.source), "SiteConfigCoLib.asp", PROCEDURE_NAME, "restoreBackupPropertyFiles", "Error calling oAdmin.restoreBackupPropertyFiles", LogLevelError)
        End If
    End If

    Set oAdmin = Nothing

    co_RestoreBackupPropertyFiles = lErr
    Err.Clear
End Function


%>