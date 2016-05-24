<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>



<!-- #include file="../CoreLib/SiteConfigCoLib.asp" -->



<%

Const REPOSITORY_MD = 0

Const REPOSITORY_AUREP = 1

Const REPOSITORY_SBREP = 2

Const REPOSITORY_WAREHOUSE = 4



Const OBJECT_REP_CONN = 0

Const SUSB_BOOK_REP_CONN = 1



Const CONN_ID = 0

Const CONN_DSN = 1

Const CONN_PREFIX = 2

Const MAX_CONN_INFO = 2



Const DBALIAS_NAME = 0

Const DBALIAS_SERVER = 1

Const DBALIAS_DBNAME = 2

Const DBALIAS_USER = 3

Const DBALIAS_PASSWORD = 4

Const DBALIAS_CLASSNAME = 5

Const DBALIAS_PLATFORM = 6

Const DBALIAS_CONFIRM = 7

Const DBALIAS_DECODED_NAME = 8

Const DBALIAS_SYSTEM_NAME = 9

Const DBALIAS_PREFIX = 10

Const DBALIAS_PORT_NUMBER = 11

Const DBALIAS_POOL_SIZE = 12

Const MAX_DBALIAS_INFO = 12





Function CreateDefaultChannels(sSiteId)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "createDefaultChannels"

	Dim lErrNumber

    Dim sFileName

    Dim oDOM

    Dim oChannels

    Dim oChannel

    Dim i

    Dim sChannelId

    Dim sChannelXML



	lErrNumber = NO_ERR

    sFileName = Server.MapPath("./") & "\defaultSiteProperties.xml"



    lErrNumber = LoadXMLDOMFromFile(aConnectionInfo, sFileName, oDOM)

    If lErrNumber <> NO_ERR Then

        Call LogErrorXML(aConnectionInfo, lErrNumber, Cstr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromFile", LogLevelTrace)

    Else

        Set oChannels = oDOM.selectNodes("//oi[@tp='2']")



        For i = 0 To (oChannels.length - 1)



            Set oChannel = oChannels.item(i)



            sChannelId = GetGUID()

            If lErrNumber = NO_ERR Then

                sChannelXML = "<mi><in><oi tp='" & TYPE_CHANNEL & "'><prs>"

                sChannelXML = sChannelXML & " <pr id='NAME' v='" & GetPropertyValue(oChannel, "CHANNEL_NAME") & "' />"

                sChannelXML = sChannelXML & " <pr id='DESC' v='" & GetPropertyValue(oChannel, "CHANNEL_DESC") & "' />"

                sChannelXML = sChannelXML & " <pr id='CHANNEL_PUBLISHED' v='" & GetPropertyValue(oChannel, "CHANNEL_PUBLISHED") & "' />"

                sChannelXML = sChannelXML & " <pr id='SERVICE_FOLDER_ID' v='" & GetPropertyValue(oChannel, "CHANNEL_ROOT_FOLDER_ID") & "' />"

                sChannelXML = sChannelXML & "</prs></oi></in></mi>"

            End If



	        If lErrNumber = NO_ERR Then

	            lErrNumber = co_CreateDefaultChannel(sSiteId, sChannelId, sChannelXML)

	            If lErrNumber <> NO_ERR Then

	            	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateDefaultChannel", LogLevelTrace)

	            End If

	        End If



            If lErrNumber <> NO_ERR Then Exit For



        Next

    End If



    Set oDOM = Nothing

    Set oChannel = Nothing

    Set oChannels = Nothing



	createDefaultChannels = lErrNumber

	Err.Clear

End Function





Function CreateDefaultDeviceTypes(sSiteId)

'*******************************************************************************

'Purpose:

'Inputs:

'Outputs:

'TO DO: Add handling for portalDevice and defaultDevice

'*******************************************************************************

Const PROCEDURE_NAME = "createDefaultDeviceTypes"

Dim lErrNumber

Dim sAdminPath

Dim sFilePath

Dim oDefaultDeviceTypesDOM

Dim oDefaultDeviceTypes

Dim oCurrentDefaultDeviceType

Dim sNewGUID



	On Error Resume Next

	lErrNumber = NO_ERR



	sFilePath = GetFilePath()

	sAdminPath = sFilePath & "admin\"



    Set oDefaultDeviceTypesDOM = Server.CreateObject("Microsoft.XMLDOM")

	oDefaultDeviceTypesDOM.async = False

	If oDefaultDeviceTypesDOM.load(sAdminPath & "defaultDeviceTypes.xml") = False Then

	    lErrNumber = ERR_XML_LOAD_FAILED

	    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error loading defaultDeviceTypes.xml file", LogLevelError)

	Else

	    Set oDefaultDeviceTypes = oDefaultDeviceTypesDOM.selectNodes("//devicetype")

	    If oDefaultDeviceTypes.length > 0 Then

	        For Each oCurrentDefaultDeviceType in oDefaultDeviceTypes

	            sNewGUID = GetGUID()

	            lErrNumber = cu_CreateDeviceType(sSiteId, sNewGUID, oCurrentDefaultDeviceType.selectSingleNode("name").text)

	            If lErrNumber <> NO_ERR Then

	                Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling cu_CreateDeviceType", LogLevelTrace)

	                Exit For

	            Else

	                oCurrentDefaultDeviceType.selectSingleNode("devicetypeID").text = sNewGUID



	                lErrNumber = cu_CreateDeviceTypeDefinitions(sSiteId, sNewGUID, oCurrentDefaultDeviceType.xml)

	                If lErrNumber <> NO_ERR Then

	                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "DeviceTypesCuLib.asp", PROCEDURE_NAME, "", "Error calling cu_CreateDeviceTypeDefinitions", LogLevelTrace)

	                    Exit For

	                End If

	            End If

	        Next

	    End If

	End If



    Set oDefaultDeviceTypesDOM = Nothing

    Set oDefaultDeviceTypes = Nothing

    Set oCurrentDefaultDeviceType = Nothing



	createDefaultDeviceTypes = lErrNumber

	Err.Clear



End Function



Function GetMDConn(aMDConn)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs: aMDConn

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "getMDConn"

	Dim lErrNumber

    Dim sMDXML

    Dim oDOM



	lErrNumber = NO_ERR



	lErrNumber = co_GetMDConn(sMDXML)

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetMDConn", LogLevelTrace)

	End If



    If lErrNumber = NO_ERR Then

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sMDXML, oDOM)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        Redim aMDConn(1)



        aMDConn(0) = GetPropertyValue(oDOM, "GROUP_DSN")

        aMDConn(1) = GetPropertyValue(oDOM, "GROUP_PREFIX")



    End If



    Set oDOM = Nothing



	getMDConn = lErrNumber

	Err.Clear

End Function



Function SetMDConn(aMDConn)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "setMDConn"

	Dim lErrNumber

	Dim sConfigXML

    Dim aSiteProperties()

    Redim aSiteProperties(MAX_SITE_PROP)



	lErrNumber = NO_ERR



    'Create the XML necessary for the call:

    sConfigXML = "<mi><oi id='admin'><prs>"

    sConfigXML = sConfigXML & "<pr id='GROUP_DSN'    v='" & aMDConn(0) & "' />"

    sConfigXML = sConfigXML & "<pr id='GROUP_PREFIX' v='" & aMDConn(1) & "' />"

    sConfigXML = sConfigXML & "</prs></oi></mi>"



	If lErrNumber = NO_ERR Then

	    lErrNumber = co_SetMDConn(sConfigXML)

	    If lErrNumber <> NO_ERR Then

	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SetMDConn", LogLevelTrace)

	    End If

	End If



    'Set the App variable to this new value:

    If lErrNumber = NO_ERR Then

        Application.Value("MD_CONN") = aMDConn(0)

    End If



    'If succeded to change the MD, set the new site to be an empty one

    If lErrNumber = NO_ERR Then

        lErrNumber = SetSite("")

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling SetSite", LogLevelTrace)

        End If

    End If



    'Finally, we need to reset all the engines to take the

    'new Connection values:

    If lErrNumber = NO_ERR Then

        lErrNumber = ResetSubscriptionEngine()

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ResetSubscriptionEngine", LogLevelTrace)

        End If

    End If



	setMDConn = lErrNumber

	Err.Clear

End Function



Function GetMRUEngines(aRUEngines)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs: aRUEngines

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "getMRUEngines"

	Dim lErrNumber

	Dim sConfigName

	Dim sConfigXML

	Dim oDOM

	Dim i



	lErrNumber = NO_ERR

	sConfigName = GetVirtualDirectoryName()



	lErrNumber = co_GetMRUEngines(sConfigName, sConfigXML)

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetMRUEngines", LogLevelTrace)

	End If



    If lErrNumber = NO_ERR Then

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sConfigXML, oDOM)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        Redim aRUEngines(MRU_COUNT)



        For i = 0 to MRU_COUNT

            aRUEngines(i) = GetPropertyValue(oDOM, "site." & sConfigName & ".MRU" & i)

            If Len(CStr(aRUEngines(i))) = 0 Then Exit For

        Next

    End If



    Set oDOM = Nothing



	getMRUEngines = lErrNumber

	Err.Clear

End Function



Function SetMRUEngines(aRUEngines, sNewEngine)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "setMRUEngines"

	Dim lErrNumber

    Dim sConfigName

    Dim sConfigXML

    Dim i

    Dim nCount



    lErrNumber = NO_ERR

    sConfigName = GetVirtualDirectoryName()

    nCount = 1



    'Create the XML necessary for the call:

    sConfigXML = "<mi><oi id='' n='" & sConfigName & "'><prs>"

    sConfigXML = sConfigXML & "<pr id='MRU0' v='" & sNewEngine & "' />"



    For i = 0 To MRU_COUNT

        If StrComp(CStr(aRUEngines(i)), CStr(sNewEngine), vbBinaryCompare) <> 0 Then

            sConfigXML = sConfigXML & "<pr id='MRU" & CStr(nCount) & "' v='" & aRUEngines(i) & "' />"

            nCount = nCount + 1

        End If



        If nCount > MRU_COUNT Then Exit For

    Next



	sConfigXML = sConfigXML & "</prs></oi></mi>"



	lErrNumber = co_SetMRUEngines(sConfigXML)

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SetMRUEngines", LogLevelTrace)

	End If



	setMRUEngines = lErrNumber

	Err.Clear

End Function



Function GetSubscriptionEngine(sEngine)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "getSubscriptionEngine"

	Dim lErrNumber

	Dim sReturnXML

	Dim oDOM



	lErrNumber = NO_ERR



    'Use the value on the application server variables, if empty, call

    'to read from the backend

    sEngine = CStr(Application.Value("SE"))

    If Len(sEngine) = 0 Then



	    lErrNumber = co_GetSubscriptionEngine(sReturnXML)

	    If lErrNumber <> NO_ERR Then

	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetSubscriptionEngine", LogLevelTrace)

	    End If



        If lErrNumber = NO_ERR Then

            lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sReturnXML, oDOM)

            If lErrNumber <> NO_ERR Then

                Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

            End If

        End If



        If lErrNumber = NO_ERR Then

            sEngine = UCase(GetPropertyValue(oDOM, "SUBSCRIPTION_ENGINE"))

        End If



    End If



    Set oDOM = Nothing



	getSubscriptionEngine = lErrNumber

	Err.Clear

End Function



Function SetSubscriptionEngine(sNewEngine)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "setSubscriptionEngine"

	Dim lErrNumber

    Dim aRUEngines()

    Dim i

    Dim nCount



	lErrNumber = NO_ERR



    'For convenience, use only upper case names:

    sNewEngine = Replace(Server.HTMLEncode(UCase(sNewEngine)), "'", "&apos;")



    If lErrNumber = NO_ERR Then

        lErrNumber = getMRUEngines(aRUEngines)

        If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getMRUEngines", LogLevelTrace)

    End If



    If lErrNumber = NO_ERR Then

	    lErrNumber = co_SetSubscriptionEngine(sNewEngine)

	    If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SetSubscriptionEngine", LogLevelTrace)

    End If



    'Call Connect to detect if the new engine is valid:

    If lErrNumber = NO_ERR Then

	    lErrNumber = ConnectToSubscriptionEngines()

	    If lErrNumber <> NO_ERR Then

	        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ConnectToSubscriptionEngines", LogLevelTrace)

	        Call co_SetSubscriptionEngine(Application.Value("SE"))

        End If

    End If





    'If succeded to change the Engine, set the new values at the application level vars,

    If lErrNumber = NO_ERR Then

        Application.Value("SE") = CStr(sNewEngine)

    End If



    'Save the MRU Engines List:

    If lErrNumber = NO_ERR Then

        lErrNumber = setMRUEngines(aRUEngines, sNewEngine)

        If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling setMRUEngines", LogLevelTrace)

    End If



    'Change site configuration on subscriptionPortal.properties file to save MRU list:

    If lErrNumber = NO_ERR Then

        'Since we just changed the engine, the current MD

        'values are not valid, reset them to empty:

        'This call will also reset the Engine:

        lErrNumber = setMDConn(Array("", ""))

        If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling setMDConn", LogLevelTrace)

    End If



	setSubscriptionEngine = lErrNumber

	Err.Clear

End Function



Function CreateMDTables()

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "CreateMDTables"

	Dim lErrNumber



	lErrNumber = NO_ERR



	lErrNumber = co_createMDTables()

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_createMDTables", LogLevelTrace)

	End If



	CreateMDTables = lErrNumber

	Err.Clear

End Function



Function CheckMDTables(bTablesReady)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "checkMDTables"

	Dim lErrNumber

	Dim sReturnXML

    Dim oDOM



	lErrNumber = NO_ERR



	lErrNumber = co_CheckMDTables(sReturnXML)

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_CheckMDTables", LogLevelTrace)

	End If



    If lErrNumber = NO_ERR Then

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sReturnXML, oDOM)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        bTablesReady = CBool(GetPropertyValue(oDOM, "exists") = "1")

    End If



    Set oDOM = Nothing



	checkMDTables = lErrNumber

	Err.Clear

End Function



Function GetAllSites(aSites, nCount)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "getAllSites"

	Dim lErrNumber

	Dim sSitesXML

	Dim oDOM

	Dim oSites

	Dim i



	lErrNumber = NO_ERR



	lErrNumber = co_GetAllSites(sSitesXML)

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetAllSites", LogLevelTrace)

	End If



    If lErrNumber = NO_ERR Then

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sSitesXML, oDOM)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        Set oSites = oDOM.selectNodes("//oi[@tp='" & TYPE_SITE & "']")

        If (Not (oSites Is Nothing)) Then



            nCount = oSites.length



            Redim aSites(nCount - 1, 3)

            For i = 0 to (nCount - 1)

                aSites(i, 0) = oSites(i).getAttribute("id")

                aSites(i, 1) = oSites(i).getAttribute("n")

                aSites(i, 2) = oSites(i).getAttribute("des")

				aSites(i, 3) = oSites(i).getAttribute("default")
            Next



        Else

            nCount = 0

        End If

    End If



    Set oDOM = Nothing

    Set oSites = Nothing



	getAllSites = lErrNumber

	Err.Clear

End Function



Function CreateSite(aSiteProperties)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

    On Error Resume Next

    Const PROCEDURE_NAME = "CreateSite"

    Dim lErrNumber

    Dim sConfigXML



    lErrNumber = NO_ERR



    'Generate new GUIDs

    If lErrNumber = NO_ERR Then

        aSiteProperties(SITE_PROP_ID) = GetGUID()

    End If



    'Create the XML necessary to save ALL properties:

    If lErrNumber = NO_ERR Then

        lErrNumber = GenerateSitePropertiesXML(aSiteProperties, &HFFFF, sConfigXML)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling GenerateSitePropertiesXML", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        lErrNumber = co_CreateSite(aSiteProperties(SITE_PROP_ID), sConfigXML)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_CreateSite", LogLevelTrace)

        End If

    End If



    CreateSite = lErrNumber

    Err.Clear

End Function



Function DeleteSite(sSiteId)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

    On Error Resume Next

    Const PROCEDURE_NAME = "DeleteSite"

    Dim lErr

    Dim aSiteProperties()

    Redim aSiteProperties(MAX_SITE_PROP)

    Dim aPortals

    Dim nCount

    Dim i

    lErr = NO_ERR



    lErr = co_DeleteSite(sSiteId)

    If lErr <> NO_ERR Then

        Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_DeleteSite", LogLevelTrace)

    Else

        'If successful, make sure we're not using it anymore for this VD:

        If Trim(sSiteId) = Application.Value("SITE_ID") Then

            Application.Value("SITE_ID")    = ""

            Application.Value("SITE_NAME")  = sSiteName

            Application.Value("AUREP_CONN") = sAuRep

            Application.Value("SBREP_CONN") = sSBRep

        End If



        lErr = ResetSubscriptionEngine()

        If lErr <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ResetSubscriptionEngine", LogLevelTrace)

        End If

    End If



    'Reset the site to nothing for all the portals pointing to that site

    If lErr = NO_ERR Then

	    lErr = cu_GetAllPortals(aPortals,nCount)

	    If lErr <> NO_ERR Then

		    Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GellAllPortals", LogLevelTrace)

		End If



		For i=0 to (nCount - 1)

			If Trim(sSiteId) = Trim(aPortals(i, 3)) Then

		        lErr = Delete_SetSite(aPortals(i, 0), aSiteProperties)

				If lErr <> NO_ERR Then

					Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling Delete_SetSite", LogLevelTrace)

				End If

			End If

		Next

    End If



    DeleteSite = lErr

    Err.Clear



End Function



Function Delete_SetSite(sPortalName,aSiteProperties)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "Delete_SetSite"

	Dim lErrNumber

    Dim sConfigXML

    Dim sSiteId

    Dim sSiteName

    Dim sAuRep

    Dim sSBRep



	lErrNumber = NO_ERR



    sSiteId   = ""

    sSiteName = ""

    sAuRep    = ""

    sSBRep    = ""



    'Create the XML necessary for the call:

    sConfigXML = "<mi><oi id='" & sSiteId & "' n=""" & Server.HTMLEncode(sPortalName) & """><prs>"

    sConfigXML = sConfigXML & "<pr id='SITE_NAME'  v='" & sSiteName & "' />"

    sConfigXML = sConfigXML & "<pr id='AUREP_CONN'  v='" & sAuRep & "' />"

    sConfigXML = sConfigXML & "<pr id='SBREP_CONN'  v='" & sSBRep & "' />"

    sConfigXML = sConfigXML & "</prs></oi></mi>"



	lErrNumber = co_SetSite(sConfigXML)

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SetSite", LogLevelTrace)

	End If



	Delete_SetSite = lErrNumber

	Err.Clear

End Function



Function SetSite(sSiteId)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

Const PROCEDURE_NAME = "SetSite"

Dim lErrNumber

Dim sConfigName

Dim sConfigXML

Dim sSiteName

Dim sAuRep

Dim sSBRep

Dim aConnectionsInfo

Dim aSiteProperties

Redim aSiteProperties(MAX_SITE_PROP)



	On Error Resume Next

	lErrNumber = NO_ERR



    If Len(sSiteId) > 0 Then

        If lErrNumber = NO_ERR Then

            aSiteProperties(SITE_PROP_ID) = sSiteId

            lErrNumber = GetSiteProperties(aSiteProperties)

            If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SetSite", LogLevelTrace)

        End If



        If lErrNumber = NO_ERR Then

            lErrNumber = GetSiteConnections(aSiteProperties(SITE_PROP_ID), aConnectionsInfo)

            If lErrNumber <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SetSite", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        sConfigName = GetVirtualDirectoryName()

        sSiteId   = aSiteProperties(SITE_PROP_ID)

        sSiteName = aSiteProperties(SITE_PROP_NAME)

        sAuRep = aConnectionsInfo(OBJECT_REP_CONN, CONN_DSN)

        sSBRep = aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_DSN)





        'Create the XML necessary for the call:

        sConfigXML = "<mi><oi id='" & sSiteId & "' n=""" & Server.HTMLEncode(sConfigName) & """><prs>"

        sConfigXML = sConfigXML & "<pr id='SITE_NAME'  v='" & Replace(Server.HTMLEncode(sSiteName), "'", "&apos;") & "' />"

        sConfigXML = sConfigXML & "<pr id='AUREP_CONN'  v='" & sAuRep & "' />"

        sConfigXML = sConfigXML & "<pr id='SBREP_CONN'  v='" & sSBRep & "' />"

        sConfigXML = sConfigXML & "</prs></oi></mi>"



	    lErrNumber = co_SetSite(sConfigXML)

	    If lErrNumber <> NO_ERR Then

	    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SetSite", LogLevelTrace)

	    End If

    End If



    'If successful, set the new values at the application level vars,

    'invalidating the rest:

    If lErrNumber = NO_ERR Then

        Application.Value("SITE_ID")    = sSiteId

        Application.Value("SITE_NAME")  = sSiteName

        Application.Value("AUREP_CONN") = sAuRep

        Application.Value("SBREP_CONN") = sSBRep

    End If





    'Now that the site has succesfully been configured, we need to regenerate the custom SQL

    'for dynamic subscriptions, this is because the SQL resides on the web server, and it

    'is not for sure that we have it

    If lErrNumber = NO_ERR Then

        If Len(sSiteId) > 0 Then

            Call ResetSubscriptionEngine()

            Call GenerateSiteDynamicSQL(sSiteId)

        End If

    End If





	SetSite = lErrNumber

	Err.Clear



End Function





Function GenerateSiteDynamicSQL(sSiteId)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

Const PROCEDURE_NAME = "GenerateSiteDynamicSQL"

Dim lErrNumber



Dim sSiteObjectsXML

Dim oSiteObjectsDOM



Dim sPropsXML

Dim oPropsDOM



Dim oServices

Dim oService



Dim oSubsSets

Dim oSubsSet



    On Error Resume Next

    lErrNumber = NO_ERR



    'Get all services sets configured for this site:

    If lErr = NO_ERR Then

        lErr = co_getObjectsForSite(sSiteId, TYPE_SERVICE_CONFIG, sSiteObjectsXML)

        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectsForSite", LogLevelTrace)



        If lErr = NO_ERR Then

            lErr = LoadXMLDOMFromString(aConnectionInfo, sSiteObjectsXML, oSiteObjectsDOM)

            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sSiteObjectsXML", LogLevelTrace)

        End If

    End If

    'Get all subscription sets configured for this service:



    If lErr = NO_ERR Then

        Set oServices = oSiteObjectsDOM.selectNodes("//oi[@tp='" & TYPE_SERVICE_CONFIG & "']")

        For Each oService In oServices



            lErr = co_getObjectProperties(sSiteId, oService.getAttribute("id"), sPropsXML)

            If lErr <> NO_ERR Then

                Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectProperties", LogLevelTrace)

            Else

                lErr = LoadXMLDOMFromString(aConnectionInfo, sPropsXML, oPropsDOM)

                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sPropsXML", LogLevelTrace)

            End If



            'Exit if error:

            If lErr <> NO_ERR Then Exit For



            Set oSubsSets = oPropsDOM.selectNodes("//oi[@tp='" & TYPE_SUBSSET_CONFIG & "']")

            For Each oSubsSet In oSubsSets

                Call co_generateSubscriptionSetSQL(sSiteId, oService.getAttribute("id"), oSubsSet.getAttribute("id"))

            Next



            'Exit if error:

            If lErr <> NO_ERR Then Exit For



        Next



        If oServices.length > 0 Then Call ResetSubscriptionEngine()



    End If



    Set oPropsDOM = Nothing

    Set oService = Nothing

    Set oServices = Nothing

    Set oSiteObjectsDOM = Nothing

    Set oSubsSet = Nothing

    Set oSubsSets = Nothing



    GenerateSiteDynamicSQL = lErrNumber

    Err.Clear



End Function





Function ParseRequestForDBAlias(oRequest, aDBAliasInfo, lRepositoryType)

'********************************************************

'*Purpose:  This function loads the addDBAlias settings.

'			Also, aPageInfo is set to corresponding values based upon

'			repository type.

'*Inputs:   oRequest (request object)

'*Outputs:  Error Code

'			aPageInfo, lRepositoryType, sErrorMessage

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "getDBAliasSettings"

	Dim lErr

	Dim sLink



	lErr = NO_ERR



	'Setting repository type to MD

	If Len(oRequest("rep")) = 0 Then

		lRepositoryType = REPOSITORY_MD

	Else

		lRepositoryType = Clng(oRequest("rep"))

	End If



	'Gather all Alias information

	Redim aDBAliasInfo(MAX_DBALIAS_INFO)

	aDBAliasInfo(DBALIAS_NAME) = oRequest("tAliasName")

	aDBAliasInfo(DBALIAS_SERVER) = oRequest("tServerName")

	aDBAliasInfo(DBALIAS_DBNAME) = oRequest("tDBName")

	aDBAliasInfo(DBALIAS_USER) = oRequest("tUser")

	aDBAliasInfo(DBALIAS_PASSWORD) = oRequest("tPassword")

	aDBAliasInfo(DBALIAS_CLASSNAME) = oRequest("tClassName")

	aDBAliasInfo(DBALIAS_PLATFORM) = oRequest("tDatabase")

	aDBAliasInfo(DBALIAS_CONFIRM) = CBool(Len(oRequest("tValidate")) > 0)

	aDBAliasInfo(DBALIAS_DECODED_NAME) = oRequest("tDecodeAlias")

	aDBAliasInfo(DBALIAS_SYSTEM_NAME) = oRequest("tSystemName")

	aDBAliasInfo(DBALIAS_PREFIX) = oRequest("tPrefix")

	aDBAliasInfo(DBALIAS_PORT_NUMBER) = oRequest("tPortNumber")

	aDBAliasInfo(DBALIAS_POOL_SIZE) = oRequest("tPoolSize")

	getDBAliasSettings = lErr

	Err.Clear



End Function





Function CheckSiteRepositories(sSiteId)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

Const PROCEDURE_NAME = "CheckSiteRepositories"

Dim lErr

Dim aConnectionsInfo



	On Error Resume Next

	lErr = NO_ERR



    'Retrieve the connections for this site:

    lErr = GetSiteConnections(sSiteId, aConnectionsInfo)



    'Check for object repository:

    If lErr = NO_ERR Then



        If Len(aConnectionsInfo(OBJECT_REP_CONN, CONN_DSN)) > 0 Then

            lErr = CheckDBAlias(aConnectionsInfo(OBJECT_REP_CONN, CONN_DSN), aConnectionsInfo(OBJECT_REP_CONN, CONN_PREFIX), REPOSITORY_AUREP)

        End If



    End If



    'Check for SBR:

    If lErr = NO_ERR Then



        If Len(aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_DSN)) > 0 Then

            lErr = CheckDBAlias(aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_DSN), aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_PREFIX), REPOSITORY_SBREP)

        End If



	End If



	CheckSiteRepositories = lErr

	Err.Clear



End Function



Function ProccessEditDBAlias(aDBAliasInfo, lRepositoryType)

'********************************************************

'*Purpose:  Proccess all information submitted by user in order to update

'			a DB Alias.

'			wrapper function.

'*Inputs:   aDBAliasInfo: array form values

'			oRequest: request object

'*Outputs:  Error Code

'********************************************************

Const PROCEDURE_NAME = "ProccessEditDBAlias"

Dim lErr

Dim sMsg



    On Error Resume Next

    lErr = NO_ERR



    'Edit DB Alias

    lErr = EditDBAliases(aDBAliasInfo)

    If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, Err.Description, Err.Source, "SiteConfigCuLib.asp", PROCEDURE_NAME, "Error calling EditDBAliases", sMsg, LogLevelTrace)



    'If the alias was updated succesfully:

    If lErr = NO_ERR Then



    	'Validate DBAlias if requested:

    	If aDBAliasInfo(DBALIAS_CONFIRM) Then



    		'Check if DB alias properties are correct.

    		lErr = CheckDBAlias(aDBAliasInfo(DBALIAS_NAME), "", lRepositoryType)

            If lErr <> NO_ERR Then

                Call LogErrorXML(aConnectionInfo, lErr, Err.Description, Err.Source, "SiteConfigCuLib.asp", PROCEDURE_NAME, "Error calling checkDBAlias", sMsg, LogLevelTrace)



    		    'If we get an error of tables not found, that's ok, since we don't have the prefix:

                'For any other error, delete the dbalias we just created, and return the original error

                If lErr = ERR_NO_TABLES_EXIST Then

                    lErr = NO_ERR

                Else

                    Call DeleteDBAlias(aDBAliasInfo)

                End If

            End If



        End If



    End If



	ProccessEditDBAlias = lErr

	Err.Clear



End Function



Function ProccessCreateDBAlias(aDBAliasInfo, lRepositoryType)

'********************************************************

'*Purpose:  Proccess all information submitted by user in order to create

'			a DB Alias.

'			wrapper function.

'*Inputs:   aDBAliasInfo: array form values

'			oRequest: request object

'*Outputs:  Error Code

'********************************************************

Const PROCEDURE_NAME = "ProccessCreateDBAlias"

Dim lErr

Dim sMsg



    On Error Resume Next

    lErr = NO_ERR



    'Create new DB Alias

    lErr = CreateDBAliases(aDBAliasInfo)

    If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, Err.Description, Err.Source, "SiteConfigCuLib.asp", PROCEDURE_NAME, "Error calling CreateDBAliases", sMsg, LogLevelTrace)



    'If the alias was created succesfully:

    If lErr = NO_ERR Then



    	'Validate DBAlias if requested:

    	If aDBAliasInfo(DBALIAS_CONFIRM) Then



    		'Check if DB alias properties are correct.

    		lErr = CheckDBAlias(aDBAliasInfo(DBALIAS_NAME), "", lRepositoryType)

            If lErr <> NO_ERR Then

                Call LogErrorXML(aConnectionInfo, lErr, Err.Description, Err.Source, "SiteConfigCuLib.asp", PROCEDURE_NAME, "Error calling checkDBAlias", sMsg, LogLevelTrace)



    		    'If we get an error of tables not found, that's ok, since we don't have the prefix:

                'For any other error, delete the dbalias we just created, and return the original error

                If lErr = ERR_NO_TABLES_EXIST Then

                    lErr = NO_ERR

                Else

                    Call DeleteDBAlias(aDBAliasInfo)

                End If

            End If



        End If



    End If



	ProccessCreateDBAlias = lErr

	Err.Clear



End Function



Function DeleteDBAlias(aDBAliasInfo)

'********************************************************

'*Purpose:  Create new DB alias. Encapsulates the function call to

'			wrapper function.

'*Inputs:   aDBAliasInfo (string array)

'*Outputs:  Error Code

'********************************************************

Dim lErr



    On Error Resume Next

    lErr = NO_ERR



    'Call wrapper function

    lErr = CO_deleteDBAlias(aDBAliasInfo)



    If lErr = NO_ERR Then

        Call ResetSubscriptionEngine()

    End If



    deleteDBAlias = lErr

    Err.Clear



End Function



Function CheckDBAlias(sDBAlias, sPrefix, lRepositoryType)

'********************************************************

'*Purpose:  Check whether tables are created in DB alias with a specific prefix.

'			Also, checks if the version tables is up-to-date.

'			wrapper function.

'*Inputs:   sDBAlias, sPrefix, lRepositoryType

'*Outputs:  Error Code, sErrorMessage

'********************************************************

    On Error Resume Next

    Const PROCEDURE_NAME = "checkDBAlias"

    Dim lErrNumber



    lErrNumber = NO_ERR



    lErrNumber = CO_checkDBAlias(sDBAlias, sPrefix, lRepositoryType)

    If lErrNumber <> NO_ERR Then

        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling CO_checkDBAlias", LogLevelTrace)

    End If



    checkDBAlias = lErrNumber

    Err.Clear

End Function



Function GetDBAliases(aDBAliases)

'********************************************************

'*Purpose:  Check whether tables are created in DB alias with a specific prefix.

'			Also, checks if the version tables is up-to-date.

'			wrapper function.

'*Inputs:   sDBAlias, sPrefix, lRepositoryType

'*Outputs:  Error Code, sErrorMessage

'********************************************************

    On Error Resume Next

    Const PROCEDURE_NAME = "getDBAliases"

    Dim lErrNumber

    Dim sAliasesXML

    Dim oDOM

    Dim oAliases

    Dim i

	Dim oObj

    Dim sEncoded



    lErrNumber = NO_ERR



    lErrNumber = co_GetDBAliases(sAliasesXML)

    If lErrNumber <> NO_ERR Then

        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetDBAliases", LogLevelTrace)

    End If



	'Create encoder/decoder object

    If lErrNumber = NO_ERR Then

        lErrNumber = CreateEncoderDecoderObject(oObj)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling CreateEncoderDecoderObject", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sAliasesXML, oDOM)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        Set oAliases = oDOM.selectNodes("//oi[@tp='" & TYPE_DBALIAS & "']")

        If (Not (oAliases Is Nothing)) Then



            Redim aDBAliases(oAliases.length - 1, 1)

            For i = 0 to (oAliases.length - 1)

                aDBAliases(i,0) = oAliases(i).getAttribute("n")

				sEncoded = GetPropertyValue(oAliases(i), "encoded")

				If Len(sEncoded) > 0 Then

					lErrNumber = DecodeDBAlias(oObj, aDBAliases(i,0), aDBAliases(i,1))

				End If

				If lErrNumber <> NO_ERR Then

					Exit For

				End If

            Next



        End If

    End If

    If lErr = NO_ERR Then
        lErr = ResetSubscriptionEngine()
    End If


    Set oDOM = Nothing

    Set oAliases = Nothing



    getDBAliases = lErrNumber

    Err.Clear

End Function



Function GetDBAliasProperties(aDBAlias,sDBName,sJServerName,sODBCName,sUserName,sPassword,sDatabaseType, sSystemName, sDefaultPrefix, sPortNumber, sPoolSize)

'********************************************************

'*Purpose:  Retrieves all properties for a DB Alias

'*Inputs:   sDBAlias

'*Outputs:  sDBName,sJServerName,sODBCName,sUserName,sPassword

'********************************************************

    On Error Resume Next

    Const PROCEDURE_NAME = "getDBAliasesProperties"

    Dim lErrNumber

    Dim sAliasesXML

    Dim oDOM

    Dim oAliases

    Dim i

	Dim oNode



    lErrNumber = NO_ERR



    lErrNumber = co_GetDBAliases(sAliasesXML)

    If lErrNumber <> NO_ERR Then

        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetDBAliases", LogLevelTrace)

    End If



    If lErrNumber = NO_ERR Then

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sAliasesXML, oDOM)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        Set oAliases = oDOM.selectSingleNode("//oi[@n='" & aDBAlias & "']")



        If (Not (oAliases Is Nothing)) Then

			Set oNode = oAliases.selectSingleNode("prs/pr[@id='databaseName']")

			sDBName = oNode.getAttribute("v")



			Set oNode = oAliases.selectSingleNode("prs/pr[@id='serverName']")

			sJServerName = oNode.getAttribute("v")



			Set oNode = oAliases.selectSingleNode("prs/pr[@id='user']")

			sUserName = oNode.getAttribute("v")



			Set oNode = oAliases.selectSingleNode("prs/pr[@id='databaseType']")

			sDatabaseType = oNode.getAttribute("v")



			Set oNode = oAliases.selectSingleNode("prs/pr[@id='systemName']")

			sSystemName = oNode.getAttribute("v")



			Set oNode = oAliases.selectSingleNode("prs/pr[@id='defaultPrefix']")

			sDefaultPrefix = oNode.getAttribute("v")



			Set oNode = oAliases.selectSingleNode("prs/pr[@id='portNumber']")

			sPortNumber = oNode.getAttribute("v")


			Set oNode = oAliases.selectSingleNode("prs/pr[@id='poolSize']")

			sPoolSize = oNode.getAttribute("v")


        End If

    End If



    Set oDOM = Nothing

    Set oAliases = Nothing



    getDBAliaseProperties = lErrNumber

    Err.Clear

End Function



Function EditDBAliases(aDBAliasInfo)

'********************************************************

'*Purpose:  Update DB alias. Encapsulates the function call to

'			wrapper function.

'*Inputs:   aDBAliasInfo (string array)

'*Outputs:  Error Code

'********************************************************

Const PROCEDURE_NAME = "EditDBAliases"

Dim lErr

Dim sAlias

Dim sName

Dim oObj



    On Error Resume Next

    lErr = NO_ERR



	If Len(aDBAliasInfo(DBALIAS_DECODED_NAME)) > 0 Then

		'Create encoder/decoder object

		If lErr = NO_ERR Then

		    lErr = CreateEncoderDecoderObject(oObj)

		    If lErr <> NO_ERR Then

		        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling CreateEncoderDecoderObject", LogLevelTrace)

		    End If

		End If



		If lErr = NO_ERR Then

		    lErr = EncodeDBAlias(oObj,aDBAliasInfo(DBALIAS_NAME),sName)

		    If lErr <> NO_ERR Then

		        Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error encoding the DBAlias Name", LogLevelTrace)

		    End If

		End If

	Else

		sName = aDBAliasInfo(DBALIAS_NAME)

	End IF



	If lErr = NO_ERR Then

		'Creating XML string needed by backend to create DB alias

		sAlias = "<mi><in><oi id='' tp='1003' n=""" & Server.HTMLEncode(sName) & """><prs>" & _

				 "<pr id='class' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_CLASSNAME)) & """/>" & _

				 "<pr id='serverName' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_SERVER)) & """/>" & _

				 "<pr id='databaseName' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_DBNAME)) & """/>" & _

				 "<pr id='user' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_USER)) & """/>" & _

				 "<pr id='databaseType' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_PLATFORM)) & """/>" & _

				 "<pr id='password' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_PASSWORD)) & """/>"



		If Strcomp(aDBAliasInfo(DBALIAS_PLATFORM), "db2") = 0 Then

			sAlias = sAlias & "<pr id=""characterSet"" v=""Unicode""/>"

		Else

		    sAlias = sAlias & "<pr id=""characterSet"" v=""""/>"

		End If



		If Len(aDBAliasInfo(DBALIAS_DECODED_NAME)) > 0 Then

			sAlias = sAlias & "<pr id=""encoded"" v=""true""/>"

		End If



		If Len(aDBAliasInfo(DBALIAS_SYSTEM_NAME)) > 0 Then

			sAlias = sAlias & "<pr id=""systemName"" v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_SYSTEM_NAME)) & """/>"

		End If



		If Len(aDBAliasInfo(DBALIAS_PREFIX)) > 0 Then

			sAlias = sAlias & "<pr id=""defaultPrefix"" v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_PREFIX)) & """/>"

		End If



		If Len(aDBAliasInfo(DBALIAS_PORT_NUMBER)) > 0 Then

			sAlias = sAlias & "<pr id=""portNumber"" v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_PORT_NUMBER)) & """/>"

		End If


		If Len(aDBAliasInfo(DBALIAS_POOL_SIZE)) > 0 Then

			sAlias = sAlias & "<pr id=""poolSize"" v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_POOL_SIZE)) & """/>"

		End If


		sAlias = sAlias & "</prs></oi></in></mi>"

	End If



	If lErr = NO_ERR Then

		lErr = CO_editDBAlias(sName,sAlias)

		If lErr <> NO_ERR Then

		    Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling CO_editDBAlias", LogLevelTrace)

		Else

		    Call ResetSubscriptionEngine()

		End If

	End If



    EditDBAliases = lErr

    Err.Clear



End Function



Function CreateDBAliases(aDBAliasInfo)

'********************************************************

'*Purpose:  Create new DB alias. Encapsulates the function call to

'			wrapper function.

'*Inputs:   aDBAliasInfo (string array)

'*Outputs:  Error Code

'********************************************************

Const PROCEDURE_NAME = "CreateDBAliases"

Dim lErr

Dim sAlias



    On Error Resume Next

    lErr = NO_ERR



    'Creating XML string needed by backend to create DB alias

    sAlias = "<mi><in><oi id='' tp='1003' n=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_NAME)) & """><prs>" & _

			 "<pr id='class' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_CLASSNAME)) & """/>" & _

			 "<pr id='serverName' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_SERVER)) & """/>" & _

			 "<pr id='databaseName' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_DBNAME)) & """/>" & _

			 "<pr id='user' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_USER)) & """/>" & _

			 "<pr id='databaseType' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_PLATFORM)) & """/>" & _

			 "<pr id='password' v=""" & Server.HTMLEncode(aDBAliasInfo(DBALIAS_PASSWORD)) & """/>" & _

			 "<pr id='portNumber' v=""19996""/>" & _

			 "<pr id='poolSize' v=""" &Server.HTMLEncode(aDBAliasInfo(DBALIAS_POOL_SIZE)) & """/>"


	If Strcomp(aDBAliasInfo(DBALIAS_PLATFORM), "db2") = 0 Then

		sAlias = sAlias & "<pr id=""characterSet"" v=""Unicode""/>"

	Else

	    sAlias = sAlias & "<pr id=""characterSet"" v=""""/>"

	End If



	sAlias = sAlias & "</prs></oi></in></mi>"



    lErr = CO_createDBAlias(sAlias)

    If lErr <> NO_ERR Then

        Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling CO_createDBAlias", LogLevelTrace)

    Else

        Call ResetSubscriptionEngine()

    End If



    CreateDBAliases = lErr

    Err.Clear



End Function



Function GetLocales(aLocales)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "getLocales"

	Dim lErrNumber

    Dim sLocalesXML

    Dim oDOM

    Dim oLocales

    Dim i



	lErrNumber = NO_ERR



	lErrNumber = co_GetLocales(CStr(Application.Value("SITE_ID")), sLocalesXML)

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetLocales", LogLevelTrace)

	End If



    If lErrNumber = NO_ERR Then

        lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sLocalesXML, oDOM)

        If lErrNumber <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

        End If

    End If



    If lErrNumber = NO_ERR Then

        Set oLocales = oDOM.selectNodes("/mi/in/oi[@tp='" & TYPE_LOCALE & "' $and$ @plid!='']")



        If oLocales.length > 0 Then

            Redim aLocales(oLocales.length - 1, 1)



            For i = 0 to (oLocales.length - 1)

                aLocales(i, 0) = oLocales(i).getAttribute("id")

                aLocales(i, 1) = oLocales(i).getAttribute("n")

            Next

		Else

			Redim aLocales(0, 1)

			aLocales(0, 0) = SYSTEM_LOCALE_ID

			aLocales(0, 1) = asDescriptors(315) 	 'Descriptor: Default

        End If

    End If



    Set oDOM = Nothing

    Set oLocales = Nothing



	getLocales = lErrNumber

	Err.Clear

End Function





Function GetSiteConnections(sSiteId, aConnectionsInfo)

'********************************************************

'*Purpose:  Reads the defaultSiteProperties.xml file and return

'           the values from it

'*Inputs:   sSiteId = The siteId

'*Outputs:  aConnectionsInfo: The information about the connections of a site.

'********************************************************

CONST PROCEDURE_NAME = "GetSiteConnections"

Dim lErr

Dim oDOM

Dim sObjectsForSiteXML

Dim oConnection



    On Error Resume Next

    lErr = NO_ERR



    lErr = co_getObjectProperties(sSiteId, sSiteId, sObjectsForSiteXML)

	If lErr <> NO_ERR Then

	    Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectsForSite", LogLevelTrace)

	Else

        lErr = LoadXMLDOMFromString(aConnectionInfo, sObjectsForSiteXML, oDOM)

        If lErr <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

        End If

    End If



    If lErr = NO_ERR Then

        Redim aConnectionsInfo(1, MAX_CONN_INFO)



        Set oConnection = oDOM.selectSingleNode("//oi[@tp = '" & TYPE_DBALIAS & "' and prs/pr[@id='GROUP' and @v='AUREP']]")

        If Not oConnection Is Nothing Then

            aConnectionsInfo(OBJECT_REP_CONN, CONN_ID) = oConnection.getAttribute("id")

            aConnectionsInfo(OBJECT_REP_CONN, CONN_DSN) = oConnection.selectSingleNode("prs/pr[@id='GROUP_DSN']").getAttribute("v")

            aConnectionsInfo(OBJECT_REP_CONN, CONN_PREFIX) = oConnection.selectSingleNode("prs/pr[@id='GROUP_PREFIX']").getAttribute("v")

        End If



        Set oConnection = oDOM.selectSingleNode("//oi[@tp = '" & TYPE_DBALIAS & "' and prs/pr[@id='GROUP' and @v='SBREP']]")

        If Not oConnection Is Nothing Then

            aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_ID) = oConnection.getAttribute("id")

            aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_DSN) = oConnection.selectSingleNode("prs/pr[@id='GROUP_DSN']").getAttribute("v")

            aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_PREFIX) = oConnection.selectSingleNode("prs/pr[@id='GROUP_PREFIX']").getAttribute("v")

        End If



    End If



    GetSiteConnections = lErr

    Err.Clear



End Function



Function SetSiteConnections(sSiteId, aConnectionsInfo)

'********************************************************

'*Purpose:  Set the site connections

'*Inputs:   sSiteId: the site to modify; aConnectionsInfo: the new connection information

'*Outputs:  None

'********************************************************

CONST PROCEDURE_NAME = "SetSiteConnections"

Dim lErr

Dim sPropertiesXML



    On Error Resume Next

    lErr = NO_ERR



    sPropertiesXML = "<mi><in><oi tp=""1001""><prs></prs>"



    sPropertiesXML = sPropertiesXML & "<mi><in>"

    sPropertiesXML = sPropertiesXML & "<oi id=""" & aConnectionsInfo(OBJECT_REP_CONN, CONN_ID) & """ tp=""" & TYPE_DBALIAS & """><prs>"

    sPropertiesXML = sPropertiesXML & " <pr id=""GROUP"" v=""AUREP"" />"

    sPropertiesXML = sPropertiesXML & " <pr id=""GROUP_DSN"" v=""" & Server.HTMLEncode(aConnectionsInfo(OBJECT_REP_CONN, CONN_DSN)) & """ />"

    sPropertiesXML = sPropertiesXML & " <pr id=""GROUP_PREFIX"" v=""" & Server.HTMLEncode(aConnectionsInfo(OBJECT_REP_CONN, CONN_PREFIX)) & """ />"

    sPropertiesXML = sPropertiesXML & "</prs></oi>"



    sPropertiesXML = sPropertiesXML & "<oi id=""" & aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_ID) & """ tp=""" & TYPE_DBALIAS & """><prs>"

    sPropertiesXML = sPropertiesXML & " <pr id=""GROUP"" v=""SBREP"" />"

    sPropertiesXML = sPropertiesXML & " <pr id=""GROUP_DSN"" v=""" & Server.HTMLEncode(aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_DSN)) & """ />"

    sPropertiesXML = sPropertiesXML & " <pr id=""GROUP_PREFIX"" v=""" & Server.HTMLEncode(aConnectionsInfo(SUSB_BOOK_REP_CONN, CONN_PREFIX)) & """ />"

    sPropertiesXML = sPropertiesXML & "</prs></oi>"



    sPropertiesXML = sPropertiesXML & "</in></mi>"



    sPropertiesXML = sPropertiesXML & "</oi></in></mi>"



    lErr = co_SetSiteProperties(sSiteId, sPropertiesXML)

    If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_SetSiteProperties", LogLevelTrace)



    SetSiteConnections = lErr

    Err.Clear



End Function



Function CreateDefaultConnections(sSiteId)

'********************************************************

'*Purpose:  Set the site connections

'*Inputs:   sSiteId: the site to modify; aConnectionsInfo: the new connection information

'*Outputs:  None

'********************************************************

CONST PROCEDURE_NAME = "CreateDefaultConnections"

Dim lErr

Dim sPropertiesXML



    On Error Resume Next

    lErr = NO_ERR



    If lErr = NO_ERR Then

        sPropertiesXML = "<mi><in>"

        sPropertiesXML = sPropertiesXML & "<oi tp=""" & TYPE_DBALIAS & """><prs>"

        sPropertiesXML = sPropertiesXML & " <pr id=""GROUP"" v=""AUREP"" />"

        sPropertiesXML = sPropertiesXML & " <pr id=""GROUP_DSN"" v="""" />"

        sPropertiesXML = sPropertiesXML & " <pr id=""GROUP_PREFIX"" v="""" />"

        sPropertiesXML = sPropertiesXML & "</prs></oi></in></mi>"

        lErr = co_CreateConnection(sSiteId, sSiteId, GetGUID(), sPropertiesXML)

        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_InsertSiteProperties", LogLevelTrace)

    End If



    If lErr = NO_ERR Then

        sPropertiesXML = "<mi><in>"

        sPropertiesXML = sPropertiesXML & "<oi id=""" & GetGUID() & """ tp=""" & TYPE_DBALIAS & """><prs>"

        sPropertiesXML = sPropertiesXML & " <pr id=""GROUP"" v=""SBREP"" />"

        sPropertiesXML = sPropertiesXML & " <pr id=""GROUP_DSN"" v="""" />"

        sPropertiesXML = sPropertiesXML & " <pr id=""GROUP_PREFIX"" v="""" />"

        sPropertiesXML = sPropertiesXML & "</prs></oi></in></mi>"

        lErr = co_CreateConnection(sSiteId, sSiteId, GetGUID(), sPropertiesXML)

        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_InsertSiteProperties", LogLevelTrace)

    End If



    CreateDefaultConnections = lErr

    Err.Clear



End Function



Function GetDefaultSiteProperties(aSiteProperties)

'********************************************************

'*Purpose:  Reads the defaultSiteProperties.xml file and return

'           the values from it

'*Inputs:

'*Outputs:  aSiteProperties: an array with Site properties filled out

'           with the default values.

'********************************************************

    On Error Resume Next

    CONST PROCEDURE_NAME = "getDefaultSiteProperties"

    Dim lErr

    Dim sErr

    Dim sFileName

    Dim oDOM

    Dim oObject

    Dim sSiteCache



    lErr = NO_ERR



    sFileName = Server.MapPath("./") & "\defaultSiteProperties.xml"



    lErr = LoadXMLDOMFromFile(aConnectionInfo, sFileName, oDOM)

	If lErr <> NO_ERR Then

	    Call LogErrorXML(aConnectionInfo, lErr, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromFile", LogLevelTrace)

	    'Error Handling: the file is not found, should we hardcode some values?

	Else



        aSiteProperties(SITE_PROP_NEW_USERS) = GetPropertyValue(oDOM, "SITE_ALLOW_NEW_USERS")

        aSiteProperties(SITE_PROP_NEW_EXPIRE) = GetPropertyValue(oDOM, "SITE_DEFAULT_EXPIRE")

        aSiteProperties(SITE_PROP_EXPIRE_VALUE) = GetPropertyValue(oDOM, "SITE_DEFAULT_EXPIRE_VALUE")

        aSiteProperties(SITE_PROP_GUI_LANG) = GetPropertyValue(oDOM, "SITE_GUI_LANG")

        aSiteProperties(SITE_PROP_USE_DHTML) = GetPropertyValue(oDOM, "SITE_DHTML")

        aSiteProperties(SITE_PROP_TMP_DIR) = GetPropertyValue(oDOM, "SITE_TMP_DIR")



        sSiteCache = GetPropertyValue(oDOM, "SITE_PROMPT_CACHE")

        If sSiteCache = "FILE" Then

			aSiteProperties(SITE_PROP_PROMPT_CACHE) = "1"

        Else

			aSiteProperties(SITE_PROP_PROMPT_CACHE) = "2"

		End If



        aSiteProperties(SITE_PROP_SUMMARY_PAGE) = GetPropertyValue(oDOM, "SITE_SUMMARY_PAGE")

        aSiteProperties(SITE_PROP_EMAIL) = GetPropertyValue(oDOM, "SITE_ADMIN_EMAIL")

        aSiteProperties(SITE_PROP_PHONE) = GetPropertyValue(oDOM, "SITE_ADMIN_PHONE")

        aSiteProperties(SITE_PROP_DEFAULT_ANSWER) = GetPropertyValue(oDOM, "SITE_DEFAULT_ANSWER")



        aSiteProperties(SITE_LOGIN_MODE) = "NC_NORMAL"

        aSiteProperties(SITE_ELEMENT_PROMPT_BLOCK_COUNT) = "30"

        aSiteProperties(SITE_OBJECT_PROMPT_BLOCK_COUNT) = "30"

        aSiteProperties(SITE_PROP_STREAM_ATTACHMENTS) = "1"





        Set oObject = oDOM.selectSingleNode("//oi[@id='DefaultLocale']")

        If Not oObject Is Nothing Then

            aSiteProperties(SITE_PROP_NEW_LOCALE) = GetPropertyValue(oObject , "LOCALE_ID")

        End If



        Set oObject = oDOM.selectSingleNode("//oi[@id='DefaultDevice']")

        If Not oObject Is Nothing Then

            aSiteProperties(SITE_PROP_DEFAULT_DEV_ID)    = GetPropertyValue(oObject, "DEVICE_ID")

            aSiteProperties(SITE_PROP_DEFAULT_DEV_NAME)  = GetPropertyValue(oObject, "DEVICE_NAME")

            aSiteProperties(SITE_PROP_DEFAULT_FOLDER_ID) = GetPropertyValue(oObject, "DEVICE_FOLDER_ID")

            aSiteProperties(SITE_PROP_DEFAULT_DEV_VALIDATION) = GetPropertyValue(oObject, "DEVICE_VALIDATION")

        End If



        Set oObject = oDOM.selectSingleNode("//oi[@id='PortalDevice']")

        If Not oObject Is Nothing Then

            aSiteProperties(SITE_PROP_PORTAL_DEV_ID)    = GetPropertyValue(oObject , "DEVICE_ID")

            aSiteProperties(SITE_PROP_PORTAL_DEV_NAME)  = GetPropertyValue(oObject , "DEVICE_NAME")

            aSiteProperties(SITE_PROP_PORTAL_FOLDER_ID) = GetPropertyValue(oObject , "DEVICE_FOLDER_ID")

        End If



    End If



    Set oDOM = Nothing

    Set oObject = Nothing



    getDefaultSiteProperties = lErr

    Err.Clear

End Function



Function SelectDefaultEngine(aRUEngines)

'********************************************************

'*Purpose:  Returns the name of the Engine that should be selected as default

'*Inputs:   aRUEngines: list of recently used engines

'*Outputs:

'********************************************************

    On Error Resume Next

    Const PROCEDURE_NAME = "selectDefaultEngine"

    Dim lErr

    Dim sName



    lErr = NO_ERR



    If lErr = NO_ERR Then

        lErr = getSubscriptionEngine(sName)

        If lErr <> NO_ERR Then

            Call LogErrorXML(aConnectionInfo, lErr, "", "", "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getSubscriptionEngine", LogLevelTrace)

        End If

    End If



    If lERR = NO_ERR Then

        'If no engine, select the Last Recently Used:

        If Len(CStr(sName)) = 0 Then



            If (Len(CStr(aRUEngines(0))) > 0) Then

                sName = CStr(aRUEngines(0))

            Else

                'If no other option, select the default

                sName = "localhost"

            End If



        End If

    End If



    selectDefaultEngine = sName

    Err.Clear

End Function



'Old functions



Function RenderRUEngines(aRUEngines())

'********************************************************

'*Purpose:  Shows a list of the Most Rencently Used Subscription Engines

'*Inputs:   aRUEngines: The array with the name of the MRU

'*Outputs:

'********************************************************

Dim i

Dim nCount



    On Error Resume Next



    'Count the number of engines:

    nCount = 0

    For i = 0 to UBound(aRUEngines)

        If aRUEngines(i) = "" Then Exit For

        nCount = nCount + 1

    Next



    'Only show text, if there a RUEngines

    If nCount > 0 Then



        Response.Write("<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0>")

        Response.Write("  <TR>")

        Response.Write("    <TD COLSPAN=2>")

        Response.Write("      <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=" & aFontInfo(N_SMALL_FONT) & """>")

        Response.Write(asDescriptors(577)) 'Descriptor:Recently used Subscription Engines:

        Response.Write("      </FONT>")

        Response.Write("    </TD>")

        Response.Write("  </TR>")

        Response.Write("  <TR>")

        Response.Write("    <TD HEIGHT=5 COLSPAN=2><IMG SRC=""../images/1ptrans.gif"" HEIGHT=5 ALT=""""></TD>")

        Response.Write("  </TR>")

        Response.Write("  <TR>")

        Response.Write("    <TD WIDTH=10><IMG SRC=""../images/1ptrans.gif"" WIDTH=10 ALT=""""></TD>")

        Response.Write("    <TD>")

        Response.Write("      <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=" & aFontInfo(N_SMALL_FONT) & """>")



        For i = 0 to nCount - 1

            Response.Write("<LI>" & Server.HTMLEncode(aRUEngines(i)) & "</LI>")

        Next



        Response.Write("      </FONT>")

        Response.Write("    </TD>")

        Response.Write("  </TR>")

        Response.Write("</TABLE>")



    End If



End Function







Function RenderList_Devices(sFolderID, sDeviceID, sDeviceType, sFolderContentsXML)

'********************************************************

'*Purpose:

'*Inputs:

'*Outputs:

'********************************************************

Dim lErr



Dim oContentsDOM

Dim oFolders

Dim oDevices

Dim oItem

Dim oParent



    On Error Resume Next

    lErr = NO_ERR



    Set oContentsDOM = Server.CreateObject("Microsoft.XMLDOM")

    oContentsDOM.async = False



    If oContentsDOM.loadXML(sFolderContentsXML) = False Then

        lErr = ERR_XML_LOAD_FAILED

        Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", "RenderList_Services", "", "Error loading folderContents.xml file", LogLevelError)

        'add error message

    Else

        Set oFolders = oContentsDOM.selectNodes("/mi/fct/oi[@tp=""" & TYPE_FOLDER & """]")

        Set oDevices = oContentsDOM.selectNodes("/mi/fct/oi[@tp=""" & TYPE_DEVICE & """]")

        Set oParent  = oContentsDOM.selectSingleNode("//fd[../a/fd[@id='" & sFolderID & "']]")



		End If



    If lErr = NO_ERR Then

        Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=""100%"">"

        Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#99ccff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"

        Response.Write "<TR BGCOLOR=""#6699cc"">"

        Response.Write "<TD WIDTH=16><IMG SRC=""../images/1ptrans.gif"" WIDTH=16></TD>"

        Response.Write "<TD><IMG SRC=""../images/1ptrans.gif"" WIDTH=2></TD>"

        Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#ffffff"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(306) & "</b></font></TD>" 'Descriptor: Name

        Response.Write "<TD>&nbsp;&nbsp;</TD>"

        Response.Write "<TD>&nbsp;&nbsp;</TD>"

        Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#ffffff"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(34) & "</b></font></TD>" 'Descriptor: Modified

        Response.Write "<TD>&nbsp;&nbsp;</TD>"

        Response.Write "<TD>&nbsp;&nbsp;</TD>"

        Response.Write "<TD><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#ffffff"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & asDescriptors(22) & "</b></font></TD>" 'Descriptor: Description

        Response.Write "<TD>&nbsp;&nbsp;</TD>"

        Response.Write "</TR>"

        Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#003366""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"



        If (oFolders.length + oDevices.length) = 0 Then



            Response.Write "<TR>"

            Response.Write "  <TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""10"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD>"

            Response.Write "</TR>"



            Response.Write "<TR>"

            Response.Write " <TD ALIGN=""CENTER"" COLSPAN=13 BGCOLOR=""#ffffff"">"

            Response.Write "    <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"

            Response.Write "      <B>" & asDescriptors(837) & "</B>" 'There are no valid objects in this folder

            Response.Write "    </FONT></TD>"

            Response.Write "  </TD>"

            Response.Write "</TR>"



            Response.Write "<TR>"

            Response.Write "  <TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""10"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD>"

            Response.Write "</TR>"



            If Not oParent Is Nothing Then

                aWizardInfo(CHANNEL_FOLDER_ID) = ""

                aWizardInfo(CHANNEL_PARENT_ID) = oParent.getAttribute("id")



                Response.Write "<TR>"

                Response.Write "  <TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""5"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD>"

                Response.Write "</TR>"



                Response.Write "<TR>"

                Response.Write "<TD COLSPAN=13 BGCOLOR=""#ffffff"">"

                Response.Write "    <FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """>"

                Response.Write "     <B><A HREF=""select_devices.asp?fid=" & oParent.getAttribute("id") & "&id=" & sDeviceID & "&device=" & sDeviceType & """>" & asDescriptors(147) & "</A></B>" 'Back to Parent Folder

                Response.Write "</FONT></TD>"

                Response.Write "</TD>"

                Response.Write "</TR>"



            End If



        Else

            'Show folders at the top:

            For Each oItem in oFolders

                Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"

                Response.Write "<TR>"

                Response.Write "<TD WIDTH=16><A HREF=""select_devices.asp?fid=" & oItem.attributes.getNamedItem("id").text & "&id=" & sDeviceID & "&device=" & sDeviceType & """><IMG SRC=""../images/folder2.gif"" HEIGHT=""16"" WIDTH=""16"" BORDER=""0"" ALT="""" /></A></TD>"

                Response.Write "<TD><IMG SRC=""../images/1ptrans.gif"" WIDTH=2></TD>"

                Response.Write "<TD>"

                Response.Write "<A HREF=""select_devices.asp?fid=" & oItem.attributes.getNamedItem("id").text & "&id=" & sDeviceID & "&device=" & sDeviceType & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & oItem.attributes.getNamedItem("n").text & "</b></font></A>"

                Response.Write "</TD>"

                Response.Write "<TD></TD>"

                Response.Write "<TD></TD>"

                Response.Write "<TD>"

                Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & DisplayDateAndTime(CDate(oItem.attributes.getNamedItem("mdt").text), "") & "</font>"

                Response.Write "</TD>"

                Response.Write "<TD></TD>"

                Response.Write "<TD></TD>"

                Response.Write "<TD>"

                Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oItem.attributes.getNamedItem("des").text & "</font>"

                Response.Write "</TD>"

                Response.Write "<TD></TD>"

                Response.Write "</TR>"

                Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""2"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"

                Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#99ccff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"

            Next



            'The the services

            For Each oItem in oDevices

                Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"

                Response.Write "<TR>"

                Response.Write "<TD WIDTH=16>"

                If sDeviceID =  oItem.attributes.getNamedItem("id").text Then

                    Response.Write "<A HREF=""modify_devices.asp?id=" & oItem.attributes.getNamedItem("id").text & "&device=" & sDeviceType & "&fid=" & sFolderID & "&n=" &  Server.URLEncode(oItem.attributes.getNamedItem("n").text) &"""><IMG SRC=""../images/check.gif"" HEIGHT=""16"" WIDTH=""16"" BORDER=""0"" ALT="""" /></A>"

                End If

                Response.Write "</TD>"

                Response.Write "<TD><IMG SRC=""../images/1ptrans.gif"" WIDTH=2></TD>"

                Response.Write "<TD>"

                Response.Write "<A HREF=""modify_devices.asp?id=" & oItem.attributes.getNamedItem("id").text & "&device=" & sDeviceType & "&fid=" & sFolderID & "&n=" &  Server.URLEncode(oItem.attributes.getNamedItem("n").text) &"""><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & oItem.attributes.getNamedItem("n").text & "</b></font></A>"

                Response.Write "</TD>"

                Response.Write "<TD></TD>"

                Response.Write "<TD></TD>"

                Response.Write "<TD>"

                Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & DisplayDateAndTime(CDate(oItem.attributes.getNamedItem("mdt").text), "") & "</font>"

                Response.Write "</TD>"

                Response.Write "<TD></TD>"

                Response.Write "<TD></TD>"

                Response.Write "<TD>"

                Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oItem.attributes.getNamedItem("des").text & "</font>"

                Response.Write "</TD>"

                Response.Write "<TD></TD>"

                Response.Write "</TR>"

                Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#ffffff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""2"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"

                Response.Write "<TR><TD COLSPAN=13 BGCOLOR=""#99ccff""><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""1"" BORDER=""0"" ALT="""" /></TD></TR>"

            Next



        End If



        Response.Write "<TR><TD COLSPAN=13 ALIGN=""right""><FORM ACTION=""modify_devices.asp""> <BR /> <BR />"

        Response.Write "<INPUT name=cancel type=submit class=""buttonClass"" value=""" & asDescriptors(120) & """></INPUT>" 'Descriptor:Cancel

        Response.Write "</FORM></TD></TR>"



        Response.Write "</TABLE>"



    End If



    Set oContentsDOM = Nothing

    Set oFolders = Nothing

    Set oDevices = Nothing

    Set oItem = Nothing



    RenderList_Devices = lErr

    Err.Clear



End Function





Function RenderDevicesFolderPath(sFolderId, sDeviceId, sDeviceType, sFolderXML)

'********************************************************

'*Purpose:  Shows the path of the selected folder

'*Inputs:   sFolderXML: with the information of a document

'*Outputs:  none

'********************************************************

Dim lErr

Dim oContentsDOM

Dim oFolder

Dim iNumFolders

Dim i



    On Error Resume Next

    lErr = NO_ERR



    'Load XML into a DOM Object:

    Set oContentsDOM = Server.CreateObject("Microsoft.XMLDOM")

    oContentsDOM.async = False

    If oContentsDOM.loadXML(sFolderXML) = False Then

        lErr = ERR_XML_LOAD_FAILED

        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", "RenderFolderPath", "", "Error loading folder xml file", LogLevelError)

    Else

        iNumFolders = CInt(oContentsDOM.selectNodes("//a").length)

        If Err.number <> NO_ERR Then

            lErr = Err.number

            Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", "RenderFolderPath", "", "Error retrieving oi nodes", LogLevelError)

        End If

    End If



    'Start by saying: you are here:

    Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"

    Response.Write asDescriptors(26) & " " 'Descriptor: You are here:



    If lErr = NO_ERR Then



        'No folders? We must be on the root:

        If iNumFolders = 0 Then

            Response.Write "<b>" & "&gt;" & "</b>" 'Root



        Else

            'Get the path node:

            Set oFolder = oContentsDOM.selectSingleNode("/mi/as")



            'Go recursively:

            For i=1 To iNumFolders



                Set oFolder = oFolder.selectSingleNode("a")

                Response.Write " &gt; "



                If i = iNumFolders Then

                    Response.Write "<b>" & oFolder.selectSingleNode("fd").attributes.getNamedItem("n").text & "</b>"

                Else

                    Response.Write "<A HREF=""select_devices.asp?fid=" & oFolder.selectSingleNode("fd").attributes.getNamedItem("id").text & "&id=" & sDeviceID & "&device=" & sDeviceType & """><font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#000000"" size=""" & aFontInfo(N_SMALL_FONT) & """><b>" & oFolder.selectSingleNode("fd").attributes.getNamedItem("n").text & "</b></font></A>"

                End If



            Next



        End If



    Else

        'add handling

    End If



    Response.Write "</font>"



    Set oContentsDOM = Nothing

    Set oFolder = Nothing



    Err.Clear

End Function



Function GetNewSiteName(sDefaultName)

'********************************************************

'*Purpose:  Returns the name of a new site.

'*Inputs:   none

'*Outputs:  The new site name

'********************************************************

Const PROCEDURE_NAME = "GetNewSiteName"

Dim lErr

Dim aNames

Dim aSites

Dim sNewName



Dim lCount, i



    On Error Resume Next

    lErr = NO_ERR



    sNewName = sDefaultName



    lErr = getAllSites(aSites, lCount)

    If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error loading calling getAllSites", LogLevelError)



    If lErr = NO_ERR Then



        If lCount > 0 Then

            Redim aNames(lCount - 1)



            For i = 0 To lCount

                aNames(i) = aSites(i, 1)

            Next



            sNewName = GetNewName(aNames, sNewName)

        End If

    End If



    GetNewSiteName = sNewName

    Err.Clear



End Function



Function CreateEncoderDecoderObject(oObj)

Const PROCEDURE_NAME = "CreateEncoderDecoderObject"



    On Error Resume Next

    Dim lErrNumber

    lErrNumber = 0



    Set oObj = Server.CreateObject("MSTRCOMDataEncoders.PrintableEncoder")

    lErrNumber = Err.number



    If lErrNumber <> NO_ERR Then



		sErrDescription = Err.description

		Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error creating MSTRCOMDataEncoders.PrintableEncoder", LogLevelError)



	End If



    CreateEncoderDecoderObject = lErrNumber

    Err.Clear



End Function



Function EncodeDBAlias(oObj, sNormalString, sEncodedString)

Const PROCEDURE_NAME = "EncodeDBAlias"



    On Error Resume Next

    Dim lErrNumber

    lErrNumber = 0



	lErrNumber = oObj.EncodeVar(sNormalString, sEncodedString)



    If lErrNumber <> True Or Err.Number <> NO_ERR Then



		lErrNumber = Err.number

		sErrDescription = Err.description

		Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error while encoding the DBAlias Name", LogLevelError)



	Else

		lErrNumber = NO_ERR

	End If



    EncodeDBAlias = lErrNumber

    Err.Clear



End Function



Function DecodeDBAlias(oObj, sEncodedString, sDecodedString)

Const PROCEDURE_NAME = "DecodeDBAlias"



    On Error Resume Next

    Dim lErrNumber

    lErrNumber = 0



	lErrNumber = oObj.DecodeVar(sEncodedString, sDecodedString)



    If lErrNumber <> True Or Err.Number <> NO_ERR Then



		lErrNumber = Err.number

		sErrDescription = Err.description

		Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error while decoding the DBAlias Name", LogLevelError)



	Else

		lErrNumber = NO_ERR

	End If



    DecodeDBAlias = lErrNumber

    Err.Clear



End Function



Function GetPropertyValue(oDOM, sPropertyId)

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

    Set oNode = oDom.selectSingleNode(".//pr[@id=""" & sPropertyId & """]")



    'By default Return an Empty Value, if the node exist, return its value:

    sValue = ""

    If Not oNode Is Nothing Then

        sValue = oNode.getAttribute("v")

    End If



    Set oNode = Nothing



    GetPropertyValue = sValue

    Err.Clear



End Function

Function FindBackupPropertyFiles(sExists)

'********************************************************

'*Purpose: checks if there are any backup property files

'*Inputs: none

'*Outputs: boolean

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "FindBackupPropertyFiles"

	Dim lErrNumber

	Dim sReturnXML

	Dim oDOM



	lErrNumber = NO_ERR



	lErrNumber = co_FindBackupPropertyFiles(sReturnXML)

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_FindBackupPropertyFiles", LogLevelTrace)

	End If



	If lErrNumber = NO_ERR Then

		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sReturnXML, oDOM)

		If lErrNumber <> NO_ERR Then

			Call LogErrorXML(aConnectionInfo, lErrNumber, CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString", LogLevelTrace)

		End If

	End If



	If lErrNumber = NO_ERR Then

		sExists = UCase(oDOM.selectSingleNode(".//er").getAttribute("des"))

	End If


    Set oDOM = Nothing

	FindBackupPropertyFiles = lErrNumber

	Err.Clear

End Function


Function RestoreBackupPropertyFiles()

'********************************************************

'*Purpose: restore property files from backups

'*Inputs: none

'*Outputs: none

'********************************************************

	On Error Resume Next

	Const PROCEDURE_NAME = "RestoreBackupPropertyFiles"

	Dim lErrNumber

	Dim sReturnXML

	Dim oDOM


	lErrNumber = NO_ERR

	lErrNumber = co_RestoreBackupPropertyFiles()

	If lErrNumber <> NO_ERR Then

		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "SiteConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_RestoreBackupPropertyFiles", LogLevelTrace)

	End If

	RestoreBackupPropertyFiles = lErrNumber

	Err.Clear

End Function


%>