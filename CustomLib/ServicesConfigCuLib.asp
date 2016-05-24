<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!-- #include file="../CoreLib/ServicesConfigCoLib.asp" -->
<!-- #include file="PromptConstCuLib.asp" -->
<%
'Array with services info:
Dim aSvcConfigInfo()

'WebHelper PROGID
Const PROGID_HELPER_RESULTSET = "WebAPIHelper.DSSXMLResultSet"

'Slicing Question anser types:
Const ANSWER_DEFAULT = ""
Const ANSWER_USER_ID = "USER_ID"
Const ANSWER_SUBSCRIPTION_ID = "SUBSCRIPTION_ID"
Const ANSWER_ADDRESS_ID = "ADDRESS_ID"
Const ANSWER_SUBSCRIPTION_GUID = "SUBSCRIPTION_GUID"
Const ANSWER_SUBSCRIPTION_SET_ID = "SUBSCRIPTION_SET_ID"
Const ANSWER_ACCOUNT_ID = "ACCOUNT_ID"
Const ANSWER_TRANS_PROPS_ID = "TRANS_PROPS_ID"
Const ANSWER_ADD_TRANS_PROPS = "ADD_TRANS_PROPS"
Const ANSWER_STATUS = "STATUS"
Const ANSWER_EXPIRATION_DATE = "EXPIRATION_DATE"
Const ANSWER_CREATED_DATE = "CREATED_DATE"
Const ANSWER_CREATED_BY = "CREATED_BY"
Const ANSWER_LAST_MOD_DATE = "LAST_MOD_DATE"
Const ANSWER_LAST_MOD_BY = "LAST_MOD_BY"
Const ANSWER_PROMPT_ANSWER = "PROMPT_ANSWER"
Const ANSWER_PREFERENCE_ID = "PREFERENCE_ID"
Const ANSWER_CONSTANT = "CONSTANT"
Const ANSWER_OTHER_ID = "OTHER"

'aSvcConfig index:
Const SVCCFG_SVC_ID = 0
Const SVCCFG_SVC_NAME = 1
Const SVCCFG_SVC_PARENT_ID = 2
Const SVCCFG_SVC_CONFIG_ID = 3
Const SVCCFG_SS_ID = 4
Const SVCCFG_SS_NAME = 5
Const SVCCFG_SS_CONFIG_ID = 6
Const SVCCFG_SS_MAP_ID = 7
Const SVCCFG_QO_ID = 8
Const SVCCFG_QO_NAME = 9
Const SVCCFG_QO_PARENT_ID = 10
Const SVCCFG_AQ_ID = 11
Const SVCCFG_AQ_NAME = 12
Const SVCCFG_AQ_PARENT_ID = 13
Const SVCCFG_STEP = 14
Const MAX_SVCCFG_INFO = 14

'Prompt Info:
Const PROMPT_INDEX = 0
Const PROMPT_TITLE = 1
Const PROMPT_DESC = 2
Const PROMPT_TYPE = 3
Const PROMPT_MIN = 4
Const PROMPT_MAX = 5
Const PROMPT_ISID = 6
Const MAX_PROMPT_INFO = 6

'Prompt Constants:
Const DssXmlExecutionResolve = 65536
Const DssXmlPromptLong = 2
Const DssXmlPromptString = 3
Const DssXmlPromptDouble = 4
Const DssXmlPromptDate = 5
Const DssXmlPromptElements = 7


'QuestionObjects Info:
Const QO_ID = 0
Const QO_NAME = 1
Const QO_DESCRIPTION = 2
Const QO_VALUE = 3
Const QO_ALTERNATE_ID = 4
Const QO_MAP_ID = 5
Const QO_IS_ID = 6
Const QO_PROMPT_COUNT = 7
Const MAX_QO_INFO = 7

Const TABLE_ID = 0
Const TABLE_COLUMNS = 1
Const TABLE_COLUMN_VALUES = 2
Const TABLE_COLUMN_GUIDS = 3
Const TABLE_DBALIAS = 4
Const MAX_TABLE_INFO = 4

Const MAP_ID = 0
Const MAP_NAME = 1
Const MAP_DESC = 2
Const MAP_DBALIAS = 3
Const MAP_FILTER = 4
Const MAP_TABLES = 5
Const MAP_QO_IS = 6
Const MAP_QO_PROMPT_COUNT = 7
Const MAX_MAP_INFO = 7

Const NEW_OBJECT_ID = "new"
Const STATIC_SS     = "static"
Const DYNAMIC_SS    = "dynamic"
Const DEFAULT_STATIC_SS_ID = "DEFAULT_STATIC_SUBSCRIPTION_SET"
Const DEFAULT_DYNAMIC_SS_ID = "DEFAULT_DYNAMIC_SUBSCRIPTION_SET"

Const PROP_NAME = "NAME"
Const PROP_DESC = "DESCRIPTION"
Const PROP_PHYSICAL_ID = "PHYSICAL_ID"
Const PROP_DEFAULT_ANSWER = "DEFAULT_SLICING_ANSWER"
Const PROP_STORAGE_MAPPING_ID = "STORAGE_MAPPING"
Const PROP_STORAGE_MAPPING_NAME = "STORAGE_MAPPING_NAME"
Const PROP_VALUE = "VALUE"

Const PROP_QUESTION_TYPE = "QUESTION_TYPE"
Const PROP_SLICED_BY = "SLICED_BY"
Const PROP_ALTERNATE_QUESTION = "ALTERNATE_QUESTION"
Const PROP_ALTERNATE_NAME = "ALTERNATE_NAME"
Const PROP_ALTERNATE_PROMPT_COUNT = "ALTERNATE_COUNT"
Const PROP_ALTERNATE_IS_ID = "ALTERNATE_IS"
Const PROP_ALTERNATE_MAP = "ALTERNATE_MAP"
Const PROP_ALTERNATE_SBR = "ALTERNATE_SBR"
Const PROP_IS_SHOWN = "IS_SHOWN"
Const PROP_IS_ALTERNATE = "IS_ALTERNATE"
Const PROP_STORE_IN_SBR = "STORE_IN_SBR"
Const PROP_PROMPT_COUNT = "PROMPT_COUNT"
Const PROP_IS_ID = "IS_ID"

Const SVC_CONFIG_CACHE_FOLDER = "admin"

'Storage types
Const STORAGE_NONE = 0
Const STORAGE_SBR_ONLY = 1
Const STORAGE_MAP_ONLY = 2
Const STORAGE_ALL      = 3

Function ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)
'********************************************************
'*Purpose: Reads the request object and retrieve the values of the aSvcConfigInfo array.
'           If necessary, retrieve the values that are needed from the Repositories,
'           such as, names, parents, etc..
'*Inputs:  oRequest: The request object
'*Outputs: aSvcConfigInfo(): an array with the minimal information needed to config a service
'********************************************************
Const PROCEDURE_NAME = "ParseRequestForSvcConfig"
Dim lErr
Dim sErr

Dim aMissingObjectsIds()
Dim aMissingObjectsParents()
Dim sObjectsXML
Dim oObject
Dim oDOM
Dim sSiteId
Dim i

    On Error Resume Next
    lErr = NO_ERR

    Redim aSvcConfigInfo(MAX_SVCCFG_INFO)

    aSvcConfigInfo(SVCCFG_SVC_ID) = oRequest("id")
    aSvcConfigInfo(SVCCFG_SVC_NAME) = oRequest("n")
    aSvcConfigInfo(SVCCFG_SVC_PARENT_ID) = oRequest("sfid")
    aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID) = oRequest("cfgid")
    aSvcConfigInfo(SVCCFG_SS_ID) = oRequest("ssid")
    aSvcConfigInfo(SVCCFG_SS_NAME) = oRequest("ssn")
    aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) = oRequest("sscfgid")
    aSvcConfigInfo(SVCCFG_SS_MAP_ID) = oRequest("ssmid")
    aSvcConfigInfo(SVCCFG_QO_ID) = oRequest("qid")
    aSvcConfigInfo(SVCCFG_QO_NAME) = oRequest("qn")
    aSvcConfigInfo(SVCCFG_QO_PARENT_ID) = oRequest("qfid")
    aSvcConfigInfo(SVCCFG_AQ_ID) = oRequest("aid")
    aSvcConfigInfo(SVCCFG_AQ_NAME) = oRequest("an")
    aSvcConfigInfo(SVCCFG_AQ_PARENT_ID) = oRequest("afid")
    aSvcConfigInfo(SVCCFG_STEP) = oRequest("st")

    'Find missing required parameters:
    If (Len(aSvcConfigInfo(SVCCFG_SVC_ID)) > 3) And (Len(aSvcConfigInfo(SVCCFG_SVC_NAME)) = 0) Or _
       (Len(aSvcConfigInfo(SVCCFG_SS_ID)) > 3) And (Len(aSvcConfigInfo(SVCCFG_SS_NAME)) = 0) Or _
       (Len(aSvcConfigInfo(SVCCFG_QO_ID)) > 3) And (Len(aSvcConfigInfo(SVCCFG_QO_NAME)) = 0) Or _
       (Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 3) And (Len(aSvcConfigInfo(SVCCFG_AQ_NAME)) = 0) Then

        sSiteId = Application.Value("SITE_ID")

        Redim aMissingObjectsIds(3)
        Redim aMissingObjectsParents(3)
        aMissingObjectsIds(0) = aSvcConfigInfo(SVCCFG_SVC_ID)
        aMissingObjectsIds(1) = aSvcConfigInfo(SVCCFG_QO_ID)
        aMissingObjectsIds(2) = aSvcConfigInfo(SVCCFG_AQ_ID)
        aMissingObjectsIds(3) = aSvcConfigInfo(SVCCFG_SS_ID)

        lErr = co_getObjectsParentInfo(sSiteId, aMissingObjectsIds, sObjectsXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectsParentInfo", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sObjectsXML, oDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sObjectsXML", LogLevelTrace)
        End If


        If lErr = NO_ERR Then
            For i = 0 To 3
                If Len(aMissingObjectsIds(i)) > 0 Then
                    Set oObject = oDOM.selectSingleNode("//oi[@id='" & aMissingObjectsIds(i) & "']")
                    If Not oObject Is Nothing Then
                        aMissingObjectsIds(i) = oObject.getAttribute("n")
                        aMissingObjectsParents(i) = oObject.getAttribute("pid")
                    End If
                End If
            Next
        End If


        aSvcConfigInfo(SVCCFG_SVC_NAME) = aMissingObjectsIds(0)
        aSvcConfigInfo(SVCCFG_QO_NAME)  = aMissingObjectsIds(1)
        aSvcConfigInfo(SVCCFG_AQ_NAME)  = aMissingObjectsIds(2)
        aSvcConfigInfo(SVCCFG_SS_NAME)  = aMissingObjectsIds(3)

        aSvcConfigInfo(SVCCFG_SVC_PARENT_ID) = aMissingObjectsParents(0)
        aSvcConfigInfo(SVCCFG_QO_PARENT_ID)  = aMissingObjectsParents(1)
        aSvcConfigInfo(SVCCFG_AQ_PARENT_ID)  = aMissingObjectsParents(2)
        aSvcConfigInfo(SVCCFG_SS_NAME)  = aMissingObjectsIds(3)

    End If

    If aSvcConfigInfo(SVCCFG_QO_ID) = NEW_OBJECT_ID Then
        aSvcConfigInfo(SVCCFG_QO_NAME) = asDescriptors(797) 'Descriptor: "Custom question"
    End If

    If aSvcConfigInfo(SVCCFG_SS_ID) = STATIC_SS Then
        aSvcConfigInfo(SVCCFG_SS_NAME) = asDescriptors(743) 'Descriptor: Static subscription sets default settings
    End If

    If aSvcConfigInfo(SVCCFG_SS_ID) = DYNAMIC_SS Then
        aSvcConfigInfo(SVCCFG_SS_NAME) = asDescriptors(791) 'Descriptor: "Dynamic subscription sets default settings"
    End If


    Set oDOM = Nothing
    Set oObject = Nothing

    ParseRequestForSvcConfig = lErr
    Err.Clear

End Function

Function CreateRequestForSvcConfig(aSvcConfigInfo)
'********************************************************
'*Purpose: Based on the aSvcConfigInfo array, creates the string that can be used
'           as the parameters of the link to a page.
'*Inputs:  aSvcConfigInfo: an array with the information needed to config a service
'*Outputs: This functions returns the string directly, not an error
'********************************************************
Dim sRequest

    sRequest = ""

    If Len(aSvcConfigInfo(SVCCFG_SVC_ID))> 0 Then sRequest = sRequest & "&id=" & aSvcConfigInfo(SVCCFG_SVC_ID)
    If Len(aSvcConfigInfo(SVCCFG_SVC_NAME))> 0 Then sRequest = sRequest & "&n=" & Server.URLEncode(aSvcConfigInfo(SVCCFG_SVC_NAME))
    If Len(aSvcConfigInfo(SVCCFG_SVC_PARENT_ID))> 0 Then sRequest = sRequest & "&sfid=" & aSvcConfigInfo(SVCCFG_SVC_PARENT_ID)
    If Len(aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID))> 0 Then sRequest = sRequest & "&cfgid=" & aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID)
    If Len(aSvcConfigInfo(SVCCFG_SS_ID))> 0 Then sRequest = sRequest & "&ssid=" & aSvcConfigInfo(SVCCFG_SS_ID)
    If Len(aSvcConfigInfo(SVCCFG_SS_NAME))> 0 Then sRequest = sRequest & "&ssn=" & Server.URLEncode(aSvcConfigInfo(SVCCFG_SS_NAME))
    If Len(aSvcConfigInfo(SVCCFG_SS_CONFIG_ID))> 0 Then sRequest = sRequest & "&sscfgid=" & aSvcConfigInfo(SVCCFG_SS_CONFIG_ID)
    If Len(aSvcConfigInfo(SVCCFG_SS_MAP_ID))> 0 Then sRequest = sRequest & "&ssmid=" & aSvcConfigInfo(SVCCFG_SS_MAP_ID)
    If Len(aSvcConfigInfo(SVCCFG_QO_ID))> 0 Then sRequest = sRequest & "&qid=" & aSvcConfigInfo(SVCCFG_QO_ID)
    If Len(aSvcConfigInfo(SVCCFG_QO_NAME))> 0 Then sRequest = sRequest & "&qn=" & Server.URLEncode(aSvcConfigInfo(SVCCFG_QO_NAME))
    If Len(aSvcConfigInfo(SVCCFG_QO_PARENT_ID))> 0 Then sRequest = sRequest & "&qfid=" & aSvcConfigInfo(SVCCFG_QO_PARENT_ID)
    If Len(aSvcConfigInfo(SVCCFG_AQ_ID))> 0 Then sRequest = sRequest & "&aid=" & aSvcConfigInfo(SVCCFG_AQ_ID)
    If Len(aSvcConfigInfo(SVCCFG_AQ_NAME))> 0 Then sRequest = sRequest & "&an=" & Server.URLEncode(aSvcConfigInfo(SVCCFG_AQ_NAME))
    If Len(aSvcConfigInfo(SVCCFG_AQ_PARENT_ID))> 0 Then sRequest = sRequest & "&afid=" & aSvcConfigInfo(SVCCFG_AQ_PARENT_ID)
    If Len(aSvcConfigInfo(SVCCFG_STEP))> 0 Then sRequest = sRequest & "&st=" & aSvcConfigInfo(SVCCFG_STEP)

    If Len(sRequest) > 0 Then sRequest = Mid(sRequest, 2)

    CreateRequestForSvcConfig = sRequest

End Function

Function RenderSvcConfigInputs(aSvcConfigInfo)
'********************************************************
'*Purpose: Based on the aSvcConfigInfo array, renders the hidden inputs that will be used
'           to submit their values to the next page.
'*Inputs:  aSvcConfigInfo: an array with the information needed to config a service
'*Outputs: This functions returns nothing, since it is a rendering one
'********************************************************

    Call Response.Write(vbCrLf)

    If Len(aSvcConfigInfo(SVCCFG_SVC_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""id"" VALUE=""" & aSvcConfigInfo(SVCCFG_SVC_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_SVC_NAME))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""n"" VALUE="""  & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SVC_NAME)) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_SVC_PARENT_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""sfid"" VALUE=""" & aSvcConfigInfo(SVCCFG_SVC_PARENT_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""cfgid"" VALUE=""" & aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_SS_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""ssid"" VALUE=""" & aSvcConfigInfo(SVCCFG_SS_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_SS_NAME))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""ssn"" VALUE=""" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SS_NAME)) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_SS_CONFIG_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""sscfgid"" VALUE=""" & aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_SS_MAP_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""ssmid"" VALUE=""" & aSvcConfigInfo(SVCCFG_SS_MAP_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_QO_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""qid"" VALUE=""" & aSvcConfigInfo(SVCCFG_QO_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_QO_NAME))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""qn"" VALUE=""" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_QO_NAME)) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_QO_PARENT_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""qfid"" VALUE=""" & aSvcConfigInfo(SVCCFG_QO_PARENT_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_AQ_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""aid"" VALUE=""" & aSvcConfigInfo(SVCCFG_AQ_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_AQ_NAME))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""an"" VALUE=""" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_AQ_NAME)) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_AQ_PARENT_ID))> 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""afid"" VALUE=""" & aSvcConfigInfo(SVCCFG_AQ_PARENT_ID) & """ />" & vbCrLf)
    If Len(aSvcConfigInfo(SVCCFG_STEP)) > 0 Then Call Response.Write("<INPUT TYPE=""HIDDEN"" NAME=""st"" VALUE=""" & aSvcConfigInfo(SVCCFG_STEP) & """ />" & vbCrLf)

End Function

Function ParseRequestForSubsSet(oRequest, aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions, sAction)
'********************************************************
'*Purpose: Reads the request object and retrieve the values of the aSvcConfigInfo array
'           and all the questions associated with this service.
'           If necessary, retrieve the values that are needed from the Repositories,
'           such as, names, parents, etc..
'*Inputs:  oRequest: The request object
'*Outputs: aSvcConfigInfo: an array with the minimal information needed to config a service
'*         aNormalQuestions: An array with info for normal (non-slicing) questions
'*         aSlicingQuestions: An array with info for slicing questions.
'*         aExtraQuestions: An array with info for questions that are not defined on the Project Repository.
'*         sAction: The action requested.
'********************************************************
Const PROCEDURE_NAME = "ParseRequestForSubsSet"
Dim lErr
Dim sErr

Dim aQuestionsId
Dim sId
Dim lCount
Dim sIds
Dim i

    On Error Resume Next
    lErr = NO_ERR

    'First, get default request values:
    If lErr = NO_ERR Then
        lErr = ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectsForSite", LogLevelTrace)
    End If

    'Now get values from Normal questions objects
    If lErr = NO_ERR Then
        sIds = oRequest("normal")

        If Len(sIds) > 0 Then
            aQuestionsId = Split(sIds, ";")
            lCount = UBound(aQuestionsId) - 1

            Redim aNormalQuestions(lCount, MAX_QO_INFO)
            For i = 0 To lCount
                sId = aQuestionsId(i)

                aNormalQuestions(i, QO_ID) = sId
                aNormalQuestions(i, QO_NAME) = oRequest("n" & sId)
                aNormalQuestions(i, QO_DESCRIPTION) = oRequest("d" & sId)
                If Len(oRequest("v" & sId)) > 0 Then
                    aNormalQuestions(i, QO_VALUE) = "true"
                Else
                    aNormalQuestions(i, QO_VALUE) = "false"
                End If
            Next
        End If
    End If

    'Values from Slicing questions objects
    If lErr = NO_ERR Then
        sIds = oRequest("slicing")

        If Len(sIds) > 0 Then
            aQuestionsId = Split(sIds, ";")
            lCount = UBound(aQuestionsId) - 1

            Redim aSlicingQuestions(lCount, MAX_QO_INFO)
            For i = 0 To lCount
                sId = aQuestionsId(i)

                aSlicingQuestions(i, QO_ID) = sId
                aSlicingQuestions(i, QO_NAME) = oRequest("n" & sId)
                aSlicingQuestions(i, QO_DESCRIPTION) = oRequest("d" & sId)
                aSlicingQuestions(i, QO_VALUE) = oRequest("v" & sId)
                aSlicingQuestions(i, QO_ALTERNATE_ID) = oRequest("a" & sId)

                If Len(oRequest("b" & sId)) > 0 Then sAction = "b" & sId
            Next
        End If
    End If

    'Finally, get values for Extra questions objects
    If lErr = NO_ERR Then
        sIds = oRequest("extra")

        If Len(sIds) > 0 Then
            aQuestionsId = Split(sIds, ";")
            lCount = UBound(aQuestionsId) - 1

            Redim aExtraQuestions(lCount, MAX_QO_INFO)
            For i = 0 To lCount
                sId = aQuestionsId(i)

                aExtraQuestions(i, QO_ID) = NEW_OBJECT_ID
                aExtraQuestions(i, QO_ALTERNATE_ID) = sId
                aExtraQuestions(i, QO_NAME) = oRequest("n" & sId)
                aExtraQuestions(i, QO_DESCRIPTION) = oRequest("d" & sId)
                aExtraQuestions(i, QO_VALUE) = oRequest("v" & sId)
                aExtraQuestions(i, QO_MAP_ID) = oRequest("m" & sId)
                aExtraQuestions(i, QO_IS_ID) = oRequest("i" & sId)
                aExtraQuestions(i, QO_PROMPT_COUNT) = oRequest("p" & sId)

                If Len(oRequest("e" & sId)) > 0 Then sAction = "e" & sId
                If Len(oRequest("r" & sId)) > 0 Then sAction = "r" & sId
            Next
        End If

    End If


    'Check for other possible actions
    If lErr = NO_ERR Then

        If Len(oRequest("next")) > 0  Then sAction = "next"
        If Len(oRequest("back")) > 0  Then sAction = "back"
        If Len(oRequest("addqo")) > 0 Then sAction = "addqo"

    End If

    Erase aQuestionsId

    ParseRequestForSubsSet = lErr
    Err.Clear

End Function


Function ParseRequestForMapInfo(oRequest, aSvcConfigInfo, aMapInfo)
'********************************************************
'*Purpose: Reads the request object and retrieve the values of the aSvcConfigInfo array
'           and all the questions associated with the Map Tables
'*Inputs:  oRequest: The request object
'*Outputs: aSvcConfigInfo: an array with the minimal information needed to config a service
'*         aTablesInfo: The Tables that the mapping use and its values
'********************************************************
Const PROCEDURE_NAME = "ParseRequestForMapInfo"
Dim lErr
Dim sErr

Dim sDBAlias
Dim sTables

Dim aTables
Dim sColumns
Dim aColumns
Dim sGUIDs
Dim aGUIDs
Dim sValues

Dim lCount, lColumnCount
Dim i, j

    On Error Resume Next
    lErr = NO_ERR

    'First, get default request values:
    If lErr = NO_ERR Then
        lErr = ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectsForSite", LogLevelTrace)
    End If

    'Get the values about the Map Info:
    If lErr = NO_ERR Then
        Redim aMapInfo(MAX_MAP_INFO)

        aMapInfo(MAP_ID) = oRequest("mid")
        aMapInfo(MAP_NAME) = oRequest("mn")
        aMapInfo(MAP_DBALIAS) = oRequest("dba")
        aMapInfo(MAP_FILTER) = oRequest("mf")
        aMapInfo(MAP_TABLES) = oRequest("tbls")
        aMapInfo(MAP_QO_IS) = oRequest("isid")
        aMapInfo(MAP_QO_PROMPT_COUNT) = oRequest("pcnt")

        If aSvcConfigInfo(SVCCFG_STEP) = DYNAMIC_SS And Len(aSvcConfigInfo(SVCCFG_AQ_ID)) = 0 And Len(aMapInfo(MAP_ID)) = 0 Then
            aMapInfo(MAP_ID) = aSvcConfigInfo(SVCCFG_SS_MAP_ID)
        End If

    End If


    ParseRequestForMapInfo = lErr
    Err.Clear

End Function


Function ParseRequestForMap(oRequest, aSvcConfigInfo, aMapInfo, aTablesInfo)
'********************************************************
'*Purpose: Reads the request object and retrieve the values of the aSvcConfigInfo array
'           and all the questions associated with the Map Tables
'*Inputs:  oRequest: The request object
'*Outputs: aSvcConfigInfo: an array with the minimal information needed to config a service
'*         aTablesInfo: The Tables that the mapping use and its values
'********************************************************
Const PROCEDURE_NAME = "ParseRequestForMap"
Dim lErr
Dim sErr

Dim sDBAlias
Dim sTables

Dim aTables
Dim sColumns
Dim aColumns
Dim sGUIDs
Dim aGUIDs
Dim sValues

Dim lCount, lColumnCount
Dim i, j

    On Error Resume Next
    lErr = NO_ERR

    'First, get default request values:
    If lErr = NO_ERR Then
        lErr = ParseRequestForMapInfo(oRequest, aSvcConfigInfo, aMapInfo)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectsForSite", LogLevelTrace)
    End If

    'Now get the tables and columns if necessary:
    If lErr = NO_ERR Then

        If Len(aMapInfo(MAP_TABLES)) > 0 Then
            aTables = Split(aMapInfo(MAP_TABLES), ";")
            lCount = UBound(aTables) - 1

            Redim aTablesInfo(lCount, MAX_TABLE_INFO)

            For i = 0  To lCount
                sColumns = oRequest("c" & aTables(i))
                sGUIDs = oRequest("g" & aTables(i))
                sValues = ""

                aColumns = Split(sColumns, ";")
                aGUIDs = Split(sGUIDs, ";")

                lColumnCount = UBound(aColumns) - 1

                For j = 0 to lColumnCount
                    sValues = sValues & oRequest("z" & aGUIDs(j)) & ";"
                Next

                aTablesInfo(i, TABLE_ID) = aTables(i)
                aTablesInfo(i, TABLE_COLUMNS) = sColumns
                aTablesInfo(i, TABLE_COLUMN_VALUES) = sValues
                aTablesInfo(i, TABLE_COLUMN_GUIDS) = sGUIDs
                aTablesInfo(i, TABLE_DBALIAS) = aMapInfo(MAP_DBALIAS)
            Next

        End If

    End If

    ParseRequestForMap = lErr
    Err.Clear

End Function


Function RenderSvcConfigPath(aSvcConfigInfo)
'********************************************************
'*Purpose: Based on the aSvcConfigInfo array, renders what part of the Service Configuration
'           a user is at.
'*Inputs:  aSvcConfigInfo: an array with the information needed to config a service
'*Outputs: This functions returns nothing, since it is a rendering one
'********************************************************
Dim bAddLink
Dim sPath


    bAddLink = False

    If InStr(aPageInfo(S_NAME_PAGE), "services_map") > 0 Then
        'We have info about the storage mapping, but the mapping can either be the
        'storage fo the SubsSet, or of a QO:
        sPath =  "</B>" & sPath

        If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) = 0 Then
            sPath = asDescriptors(812) & sPath '"Subscription Set Storing Mapping"
        Else
            If StrComp(aMapInfo(MAP_ID), NEW_OBJECT_ID) = 0 Then
                sPath = Replace(asDescriptors(813), "#",  Server.HTMLEncode(aSvcConfigInfo(SVCCFG_AQ_NAME))) & sPath 'Descriptor:New Storing Mapping for #
            Else
                sPath = Server.HTMLEncode(aMapInfo(MAP_NAME))
            End If
        End If

        sPath = "&gt; <B>" & sPath

        bAddLink = True
    End If

    If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then

        sPath =  "</B>" & sPath

        If bAddLink Then
            sPath = "<A HREF=""services_select_qo.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo) & """>" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_AQ_NAME)) & "</A>" & sPath
        Else
            sPath =  Server.HTMLEncode(aSvcConfigInfo(SVCCFG_AQ_NAME)) & sPath
        End If

        sPath = "&gt; <B>" & sPath

        bAddLink = True

    End If


    If Len(aSvcConfigInfo(SVCCFG_QO_ID))> 0 Then

        sPath =  "</B>" & sPath

        If bAddLink Then
            sPath = "<A HREF=""services_select_qo.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo) & """>" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_QO_NAME)) & "</A>" & sPath
        Else
            sPath =  Server.HTMLEncode(aSvcConfigInfo(SVCCFG_QO_NAME)) & sPath
        End If

        sPath = "&gt; <B>" & sPath

        bAddLink = True

    End If


    If Len(aSvcConfigInfo(SVCCFG_SS_ID))> 0 Then

        sPath =  "</B>" & sPath

        If bAddLink Then
            sPath = "<A HREF=""services_subsset.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo) & """>" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SS_NAME)) & "</A>" & sPath
        Else
            sPath =  Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SS_NAME)) & sPath
        End If

        sPath = "&gt; <B>" & sPath

        bAddLink = True

    End If

    If Len(aSvcConfigInfo(SVCCFG_SVC_ID))> 0 Then

        sPath =  "</B>" & sPath

        If bAddLink Then
            If aSvcConfigInfo(SVCCFG_STEP) = STATIC_SS Then
                sPath = "<A HREF=""services_subsset_modify.asp?back=true&" & CreateRequestForSvcConfig(aSvcConfigInfo) & """>" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SVC_NAME)) & "</A>" & sPath
            ElseIf aSvcConfigInfo(SVCCFG_STEP) = DYNAMIC_SS Then
                sPath = "<A HREF=""services_subsset_modify.asp?back=true&" & CreateRequestForSvcConfig(aSvcConfigInfo) & """>" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SVC_NAME)) & "</A>" & sPath
            End If
        Else
            sPath =  Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SVC_NAME)) & sPath
        End If

        sPath = "&gt; <B>" & sPath

        bAddLink = True

    End If

    sPath = "<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_MEDIUM_FONT) & """>" & asDescriptors(765) & sPath & "</FONT>" 'Descriptor:Configuring:
    Response.Write sPath

End Function



Function GetConfiguredServices(aServicesInfo)
'********************************************************
'*Purpose: Return all services that has been configured in
'           one way or another
'*Inputs:  none
'*Outputs: An array of Configured Services with:
'           aSerivcesInfo(i, 0) = ServiceId
'           aSerivcesInfo(i, 1) = Service Name
'           aSerivcesInfo(i, 2) = Service Parent Id
'           aSerivcesInfo(i, 3) = Service Parent Name
'           aSerivcesInfo(i, 4) = The configuration object Id
'********************************************************
Const PROCEDURE_NAME = "getConfiguredServices"
Dim lErr
Dim sServiceMapsXML
Dim sObjectsParentInfoXML
Dim sSiteId

Dim oMDObjectsDOM
Dim oRepObjectsDOM
Dim oServiceMaps
Dim oServices
Dim oService
Dim oParent

Dim aServicesIDs()
Dim i
Dim lCount

    On Error Resume Next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    'Get all services maps:
    If lErr = NO_ERR Then
        lErr = co_getObjectsForSite(sSiteId, TYPE_SERVICE_CONFIG, sServiceMapsXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectsForSite", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sServiceMapsXML, oMDObjectsDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sServiceMapsXML", LogLevelTrace)
        End If

    End If

    'From the services maps, get all the corresponding services:
    If lErr = NO_ERR Then

        'Get service maps nodes:
        Set oServiceMaps = oMDObjectsDOM.selectNodes("//oi[@tp='" & TYPE_SERVICE_CONFIG & "']")
        lCount = oServiceMaps.length

        'If there are any services configured:
        If lCount > 0 Then

            ReDim aServicesIDs(lCount - 1)
            For i = 0 To lCount - 1
                aServicesIDs(i) = CStr(oServiceMaps(i).getAttribute("phid"))
            Next

            'lErr = getIDsFromConfigObjects(oServiceMaps, aServicesIDs)
            'If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getIDsFromConfigObjects", LogLevelTrace)
            'Get the Project ID of all the services:

            'Now get the ObjectInfo for all the services in the list:
            lErr = co_getObjectsParentInfo(sSiteId, aServicesIDs, sObjectsParentInfoXML)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectParentInfo", LogLevelTrace)
            Else
                lErr = LoadXMLDOMFromString(aConnectionInfo, sObjectsParentInfoXML, oRepObjectsDOM)
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sObjectsParentInfoXML", LogLevelTrace)
            End If


            If lErr = NO_ERR Then
                Set oServices = oRepObjectsDOM.selectNodes("//oi[@tp='" & TYPE_SERVICE & "']")
                lCount = oServices.length

                If lCount > 0 Then

                    ReDim aServicesInfo(lCount - 1, 4)

                    For i = 0 To lCount - 1

                        Set oService = oServices(i)
                        Set oParent = oRepObjectsDOM.selectSingleNode("//a[@id='" & oServices(i).getAttribute("pid") & "']")

                        aServicesInfo(i, 0) = oService.getAttribute("id")
                        aServicesInfo(i, 1) = oService.getAttribute("n")

                        aServicesInfo(i, 2) = oParent.getAttribute("id")
                        aServicesInfo(i, 3) = oParent.getAttribute("n")

                        aServicesInfo(i, 4) = oMDObjectsDOM.selectSingleNode("//oi[prs/pr[@v='" & aServicesInfo(i, 0) & "']]").getAttribute("id")
                    Next

                End If
            End If
        End If
    End If

    Set oMDObjectsDOM = Nothing
    Set oParent = Nothing
    Set oRepObjectsDOM = Nothing
    Set oService = Nothing
    Set oServiceMaps = Nothing
    Set oServices = Nothing
    Erase aServices

    getConfiguredServices = lErr
    Err.Clear

End Function

Function GetIDsFromConfigObjects(oConfigObjects, aIDs())
'********************************************************
'*Purpose: The IDs of the corresponding config objects
'*Inputs:  aConfigObjects: a Node collection of the Config objects we want to get the IDs
'*Outputs: aIDs: An array of IDs
'********************************************************
Const PROCEDURE_NAME = "getConfigObject"
Dim lCount

    lCount = oConfigObjects.length

    'Get the Project ID of all the services:
    ReDim aIDs(lCount - 1)
    For i = 0 To lCount - 1
        aIDs(i) = oConfigObjects(i).getAttribute("phid")
    Next

End Function


Function GetConfigObjectID(sObjectId, sType, sConfigObjectID)
'********************************************************
'*Purpose: Retrieves the configuration objects id of a given object id.
'           If there are no configuration objects for this object, it returns NEW_OBJECT_ID
'*Inputs:  sObjectId
'*Outputs: aIDs: An array of IDs
'********************************************************
Const PROCEDURE_NAME = "getConfigObjectID"
Dim lErr

Dim sSiteObjectsXML
Dim sSiteId
Dim sSearchType
Dim oDOM
Dim oSvcConfig

    On Error Resume next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    'Transform the type, if necessary to a MD type:
    Select Case sType
    Case TYPE_QUESTION
        sSearchType = TYPE_QUESTION_CONFIG
    Case  TYPE_SUBSET
        sSearchType = TYPE_SUBSSET_CONFIG
    Case TYPE_SERVICE
        sSearchType = TYPE_SERVICE_CONFIG
    Case Else
        sSearchType = sType
    End Select

    'Make sure we got the right object Id:
    If sSearchType = TYPE_SUBSSET_CONFIG Then
        If sObjectId = DYNAMIC_SS Then
            sObjectId = DEFAULT_DYNAMIC_SS_ID
        ElseIf sObjectId = STATIC_SS Then
            sObjectId = DEFAULT_STATIC_SS_ID
        End If
    End If


    If lErr = NO_ERR Then
        lErr = co_getObjectsForSite(sSiteId, sSearchType, sSiteObjectsXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectsForSite", LogLevelTrace)
    End If

    If lErr = NO_ERR Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sSiteObjectsXML, oDOM)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sSiteObjectsXML", LogLevelTrace)
    End If

    If lErr = NO_ERR Then
        Set oSvcConfig = oDOM.selectSingleNode("//oi[@phid = '" & sObjectId & "']")
        If oSvcConfig Is Nothing Then
            sConfigObjectID = NEW_OBJECT_ID
        Else
            sConfigObjectID = oSvcConfig.getAttribute("id")
        End If
    End If

    set oSiteObjectsDOM = Nothing
    Set oSvcConfig = Nothing

    getConfigObjectID = lErr
    Err.Clear

End Function

Function GetSvcConfigDefaultAnswer(aSvcConfigInfo, sAnswer)
'********************************************************
'*Purpose: The IDs of the corresponding config objects
'*Inputs:  aConfigObjects: a Node collection of the Config objects we want to get the IDs
'*Outputs: aIDs: An array of IDs
'********************************************************
Const PROCEDURE_NAME = "getSvcConfigDefaultAnswer"
Dim lErr

Dim sSiteId
Dim sObjectPropsXML
Dim oDOM
Dim oConfig

    On Error Resume next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    'Make sure we have a config object id:
    If Len(aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID)) = 0 Then
        lErr = getConfigObjectID(aSvcConfigInfo(SVCCFG_SVC_ID), TYPE_SERVICE, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID))
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getConfigObjectID", LogLevelTrace)
    End If


    If lErr = NO_ERR Then

        If aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID) = NEW_OBJECT_ID Then
            sAnswer = ANSWER_DEFAULT

        Else
            lErr = co_getObjectProperties(sSiteId, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID), sObjectPropsXML)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectProperties", LogLevelTrace)
            Else
                lErr = LoadXMLDOMFromString(aConnectionInfo, sObjectPropsXML, oDOM)
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sObjectPropsXML", LogLevelTrace)
            End If

            If lErr = NO_ERR Then
                sAnswer = GetPropertyValue(oDOM.selectSingleNode("/mi/in/oi"), PROP_DEFAULT_ANSWER)
            End If
        End If
    End If

    Set oDOM = Nothing
    Set oConfig = Nothing

    getSvcConfigDefaultAnswer = lErr
    Err.Clear

End Function


Function SetSvcConfigDefaultAnswer(aSvcConfigInfo, sAnswer)
'********************************************************
'*Purpose: The IDs of the corresponding config objects
'*Inputs:  aConfigObjects: a Node collection of the Config objects we want to get the IDs
'*Outputs: aIDs: An array of IDs
'********************************************************
Const PROCEDURE_NAME = "setSvcConfigDefaultAnswer"
Dim lErr

Dim sSiteId
Dim sObjectPropsXML

    On Error Resume next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    'Make sure we have a config object id:
    If Len(aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID)) = 0 Then
        lErr = getConfigObjectID(aSvcConfigInfo(SVCCFG_SVC_ID), TYPE_SERVICE, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID))
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getConfigObjectID", LogLevelTrace)
    End If


    If lErr = NO_ERR Then

        sObjectPropsXML = GenerateServiceConfigXML(aSvcConfigInfo, sAnswer)

        'If a new object
        If aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID) = NEW_OBJECT_ID Then

            'With a default answer, don't create a new object:
            If sAnswer <> ANSWER_DEFAULT Then
                aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID) = GetGUID()

                lErr = co_createObject(sSiteId, sSiteId, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID), sObjectPropsXML)
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_createObject", LogLevelTrace)

            End If

        Else
            lErr = co_updateObjectProperties(sSiteId, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID), sObjectPropsXML)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_updateObjectProperties", LogLevelTrace)
        End If
    End If

    setSvcConfigDefaultAnswer = lErr
    Err.Clear

End Function


Function GetSubscriptionSets(aSvcConfigInfo, nType, aSubsSets)
'********************************************************
'*Purpose: Return the static or Dynamic subscription sets for a given service
'*Inputs:  aSvcConfigInfo: the Info array for Service Conffig
'*         nType: If the function should return dynamic or static subscription sets.
'*Outputs: aSubsSets: The array of subscription sets.
'********************************************************
Const PROCEDURE_NAME = "getSubscriptionSets"
Dim lErr

Dim sSiteId

Dim sSubsSetsXML
Dim sDefaultId
Dim oSubsSetsDOM
Dim oSubsSets

Dim sPropsXML
Dim oPropsDOM
Dim oConfig

Dim lCount

    On Error Resume next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    'Get all subscription sets:
    If lErr = NO_ERR Then
        lErr = co_getSubscriptionSetsForService(sSiteId, aSvcConfigInfo(SVCCFG_SVC_ID), sSubsSetsXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getSubscriptionSetsForService", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sSubsSetsXML, oSubsSetsDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sSubsSetsXML", LogLevelTrace)
        End If
    End If

    'Get configured subscription sets:
    If lErr = NO_ERR Then

        If Len(aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID)) = 0 Then
            lErr = getConfigObjectID(aSvcConfigInfo(SVCCFG_SVC_ID), TYPE_SERVICE, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID))
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getConfigObjectID", LogLevelTrace)
        End If

        If lErr = NO_ERR Then

            If aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID) = NEW_OBJECT_ID Then
                Call LoadXMLDOMFromString(aConnectionInfo, "<mi><in></in></mi>", oPropsDOM)
            Else
                lErr = co_getObjectProperties(sSiteId, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID), sPropsXML)
                If lErr <> NO_ERR Then
                    Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectProperties", LogLevelTrace)
                Else
                    lErr = LoadXMLDOMFromString(aConnectionInfo, sPropsXML, oPropsDOM)
                    If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sPropsXML", LogLevelTrace)
                End If
            End If
        End If

    End If

    'Check for the default subscription set:
    If lErr = NO_ERR Then
        If nType = STATIC_SS Then
            sDefaultId = DEFAULT_STATIC_SS_ID
        Else
            sDefaultId = DEFAULT_DYNAMIC_SS_ID
        End If

        Set oConfig = oPropsDOM.selectSingleNode("//oi[@tp='" & TYPE_SUBSSET_CONFIG & "' and prs/pr[@v='" & sDefaultId & "']]")
        If Not oConfig Is Nothing Then
            aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) = oConfig.getAttribute("id")
        Else
            aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) = ""
        End If
    End If

    'Return the information of the subscription set of this service
    If lErr = NO_ERR Then
        If nType = STATIC_SS Then
            Set oSubsSets = oSubsSetsDOM.selectNodes("//oi[@tp='" & TYPE_SUBSET & "' and @static='yes']")
        Else
            'Set oSubsSets = oSubsSetsDOM.selectNodes("//oi[@tp='" & TYPE_SUBSET & "']")
            Set oSubsSets = oSubsSetsDOM.selectNodes("//oi[@tp='" & TYPE_SUBSET & "' and @static='no']")
        End If
        lCount = oSubsSets.length

        If lCount > 0 Then

            Redim aSubsSets(lCount - 1, 2)

            For i = 0 To lCount - 1
                aSubsSets(i, 0) = oSubsSets(i).getAttribute("id")
                aSubsSets(i, 1) = oSubsSets(i).getAttribute("n")
                'aSubsSets(i, 1) = getSubscriptionSetName(aSubsSets(i, 0), oSubsSets(i))

                Set oConfig = oPropsDOM.selectSingleNode("//oi[@tp='" & TYPE_SUBSSET_CONFIG & "' and prs/pr[@v='" & aSubsSets(i, 0) & "']]")
                If Not oConfig Is Nothing Then aSubsSets(i, 2) = oConfig.getAttribute("id")
            Next

        End If
    End If


    Set oConfig = Nothing
    Set oPropsDOM = Nothing
    Set oSubsSets = Nothing
    Set oSubsSetsDOM = Nothing

    getSubscriptionSets = lErr
    Err.Clear

End Function

Function GetSubscriptionSetName(sSubsSetId, oSubsSetDOM)
'********************************************************
'*Purpose: Return the name associated with a subscription set.
'           A subscription set itself, has no name, so we associate
'           the name of the schedules to it.
'*Inputs:  sSubsSetId: The ID of the subscription set
'*         oSubsSetsDOM: The DOM object of the subsSet
'*Outputs: returns the name, not an error number (no API calls to return an error)
'********************************************************
Const PROCEDURE_NAME = "getSubscriptionSetName"
Dim sName
Dim oSchedules
Dim oSchedule

    On Error Resume next

    'Use the ID as default
    sName = sSubsSetId

    Set oSchedules = oSubsSetDOM.selectNodes("mi/oi[@tp='" & TYPE_SCHEDULE & "']")
    lCount = oSchedules.length

    If lCount > 0 Then

        sName = ""
        For Each oSchedule in oSchedules
            sName = sName & oSchedule.getAttribute("n") & ", "
        Next

        sName = Left(sName, Len(sName) - 2)

    End If

    Set oSchedule = Nothing
    Set oSchedules = Nothing

    getSubscriptionSetName = sName
    Err.Clear

End Function


Function GetSubscriptionSetConfig(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
'********************************************************
'*Purpose: Return the configuration of the subscription sets based on the information on aSvcConfigInfo
'*Inputs:  aSvcConfigInfo: the Info array for Service Conffig
'*Outputs: aNormalQuestions: An array with info for normal (non-slicing) questions
'*         aSlicingQuestions: An array with info for slicing questions.
'*         aExtraQuestions: An array with info for questions that are not defined on the Project Repository.
'********************************************************
Const PROCEDURE_NAME = "getSubscriptionSetConfig"
Dim lErr

Dim sCacheName
Dim sCacheXML

    On Error Resume Next
    lErr = NO_ERR

    If lErr = NO_ERR Then
        sCacheName = GetSvcConfigCacheName(aSvcConfigInfo)
        lErr = ReadCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, sCacheXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ReadCache", LogLevelTrace)
    End If

    If lErr = NO_ERR Then
        If Len(sCacheXML) = 0 Then
            lErr = getQuestionsConfigFromMD(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getQuestionsForServiceFromMD", LogLevelTrace)
        Else
            lErr = getQuestionsConfigFromCache(aSvcConfigInfo, sCacheXML, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getQuestionsForServiceFromCache", LogLevelTrace)
        End If
    End If

    getSubscriptionSetConfig = lErr
    Err.Clear

End Function

Function GetQuestionsConfigFromMD(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
'********************************************************
'*Purpose: Return the configuration of the question objects of a subscription set from the MD
'*Inputs:  aSvcConfigInfo: the Info array for Service Conffig
'*Outputs: aNormalQuestions: An array with info for normal (non-slicing) questions
'*         aSlicingQuestions: An array with info for slicing questions.
'*         aExtraQuestions: An array with info for questions that are not defined on the Project Repository.
'********************************************************
Const PROCEDURE_NAME = "getQuestionsConfigFromMD"
Dim lErr

Dim sSiteId

Dim sQuestionsXML
Dim oQuestionsDOM
Dim oQuestions

Dim sSubsSetXML
Dim oSubsSetDOM
Dim oSubsSet
Dim oConfigQuestion
Dim sQuestionId

Dim aMissingObjectsIds()
Dim aCustom()
Dim lMissingCount
Dim sMissingObjectsXML
Dim oMissingObjectsDOM
Dim oObject

Dim lCount
Dim i

    On Error Resume next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")
    lMissingCount = 0

    'Get all subscription sets:
    If lErr = NO_ERR Then
        lErr = co_getQuestionsForService(sSiteId, aSvcConfigInfo(SVCCFG_SVC_ID), sQuestionsXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getQuestionsForServiceForService", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sQuestionsXML, oQuestionsDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sQuestionsXML", LogLevelTrace)
        End If
    End If


    'Get configured subscription set:
    If lErr = NO_ERR Then

        If Len(aSvcConfigInfo(SVCCFG_SS_CONFIG_ID)) = 0 Then
            lErr = getConfigObjectID(aSvcConfigInfo(SVCCFG_SS_ID), TYPE_SUBSSET, aSvcConfigInfo(SVCCFG_SS_CONFIG_ID))
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getConfigObjectID", LogLevelTrace)
        End If

        If lErr = NO_ERR Then

            If aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) = NEW_OBJECT_ID Then
                Call LoadXMLDOMFromString(aConnectionInfo, "<mi><in></in></mi>", oSubsSetDOM)
            Else
                lErr = co_getObjectProperties(sSiteId, aSvcConfigInfo(SVCCFG_SS_CONFIG_ID), sSubsSetXML)
                If lErr <> NO_ERR Then
                    Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectProperties", LogLevelTrace)
                Else
                    lErr = LoadXMLDOMFromString(aConnectionInfo, sSubsSetXML, oSubsSetDOM)
                    If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sSubsSetXML", LogLevelTrace)
                End If
            End If
        End If

    End If

    'Get subscription set properties:
    If lErr = NO_ERR Then
        Set oSubsSet = oSubsSetDOM.selectSingleNode("mi/in/oi")

        If Not oSubsSet Is Nothing Then
            aSvcConfigInfo(SVCCFG_SS_MAP_ID) = GetPropertyValue(oSubsSet, PROP_STORAGE_MAPPING_ID)
        End If

        If Len(aSvcConfigInfo(SVCCFG_SS_MAP_ID)) = 0 Then
            aSvcConfigInfo(SVCCFG_SS_MAP_ID) = NEW_OBJECT_ID
        End If
    End If

    'Get Normal questions:
    If lErr = NO_ERR Then
        Set oQuestions = Nothing
        Set oQuestions = oQuestionsDOM.selectNodes("//oi[@tp='" & TYPE_QUESTION & "' and @slicing='no']")
        lCount = oQuestions.length

        If lCount > 0 Then
            Redim aNormalQuestions(lCount - 1, MAX_QO_INFO)

            For i = 0 To lCount - 1
                aNormalQuestions(i, QO_ID) = oQuestions(i).getAttribute("id")
                aNormalQuestions(i, QO_NAME) = oQuestions(i).getAttribute("n")
                aNormalQuestions(i, QO_DESCRIPTION) = getQuestionDescription(oQuestions(i))

                'Check if this question is configured, by default, always show:
                Set oConfigQuestion = Nothing
                Set oConfigQuestion = oSubsSetDOM.selectSingleNode("//oi[@tp='" & TYPE_QUESTION_CONFIG & "' and prs/pr[@v='" & aNormalQuestions(i, QO_ID) & "']]")
                If oConfigQuestion Is Nothing Then
                    aNormalQuestions(i, QO_VALUE) = "true"
                Else
                    aNormalQuestions(i, QO_VALUE) = GetPropertyValue(oConfigQuestion, PROP_IS_SHOWN)
                End If
            Next
        End If
    End If


    'Get slicing questions:
    If lErr = NO_ERR Then
        Set oQuestions = Nothing
        Set oQuestions = oQuestionsDOM.selectNodes("//oi[@tp='" & TYPE_QUESTION & "' and @slicing='yes']")
        lCount = oQuestions.length

        If lCount > 0 Then
            Redim aSlicingQuestions(lCount - 1, MAX_QO_INFO)

            For i = 0 To lCount - 1
                aSlicingQuestions(i, QO_ID) = oQuestions(i).getAttribute("id")
                aSlicingQuestions(i, QO_NAME) = oQuestions(i).getAttribute("n")
                aSlicingQuestions(i, QO_DESCRIPTION) = getQuestionDescription(oQuestions(i))

                'Check if this questions is configured:
                Set oConfigQuestion = Nothing
                Set oConfigQuestion = oSubsSetDOM.selectSingleNode("//oi[@tp='" & TYPE_QUESTION_CONFIG & "' and prs/pr[@v='" & aSlicingQuestions(i, QO_ID) & "']]")
                If Not oConfigQuestion Is Nothing Then
                    aSlicingQuestions(i, QO_ALTERNATE_ID) = GetPropertyValue(oConfigQuestion, PROP_ALTERNATE_QUESTION)
                    If Len(aSlicingQuestions(i, QO_ALTERNATE_ID)) > 0 Then
                        aSlicingQuestions(i, QO_VALUE) = ANSWER_OTHER_ID
                    Else
                        aSlicingQuestions(i, QO_VALUE) = GetPropertyValue(oConfigQuestion, PROP_SLICED_BY)
                    End If

                End If
            Next
        End If
    End If

    'Get extra questions:
    If lErr = NO_ERR Then
        Set oQuestions = Nothing
        Set oQuestions = oSubsSetDOM.selectNodes("//oi[@tp='" & TYPE_QUESTION_CONFIG & "' and prs/pr[@id='" & PROP_QUESTION_TYPE & "' and @v='1']]")
        lCount = oQuestions.length

        If lCount > 0 Then
            Redim aExtraQuestions(lCount - 1, MAX_QO_INFO)
            Redim Preserve aMissingObjectsIds(lCount - 1)
            lMissingCount = 0

            For i = 0 To lCount - 1
                Set oConfigQuestion = oQuestions(i)
                aExtraQuestions(i, QO_ID) = NEW_OBJECT_ID
                aExtraQuestions(i, QO_ALTERNATE_ID) = GetPropertyValue(oConfigQuestion, PROP_PHYSICAL_ID)
                aExtraQuestions(i, QO_VALUE) = GetPropertyValue(oConfigQuestion, PROP_IS_ALTERNATE)
                aExtraQuestions(i, QO_MAP_ID) = GetPropertyValue(oConfigQuestion, PROP_STORAGE_MAPPING_ID)
                If Len(aExtraQuestions(i, QO_MAP_ID)) = 0 Then aExtraQuestions(i, QO_MAP_ID) = "sbr"
                aExtraQuestions(i, QO_IS_ID) = GetPropertyValue(oConfigQuestion, PROP_IS_ID)
                aExtraQuestions(i, QO_PROMPT_COUNT) = GetPropertyValue(oConfigQuestion, PROP_PROMPT_COUNT)
                aExtraQuestions(i, QO_DESCRIPTION) = GetStorageMapName(aExtraQuestions(i, QO_MAP_ID))

                aMissingObjectsIds(i) = aExtraQuestions(i, QO_ALTERNATE_ID)

            Next

            'Get the name of the questions from the Aurora repository:
            lErr = co_getObjectsParentInfo(sSiteId, aMissingObjectsIds, sMissingObjectsXML)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getObjectsParentInfo", LogLevelTrace)
            Else
                lErr = LoadXMLDOMFromString(aConnectionInfo, sMissingObjectsXML, oMissingObjectsDOM)
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sMissingObjectsXML", LogLevelTrace)
            End If

            If lErr = NO_ERR Then

                For i = 0 To lCount - 1
                    Set oObject = oMissingObjectsDOM.selectSingleNode("//oi[@id='" & aExtraQuestions(i, QO_ALTERNATE_ID) & "']")
                    If Not (oObject Is Nothing) Then
                        aExtraQuestions(i, QO_NAME) = oObject.getAttribute("n")
                    Else
                        aExtraQuestions(i, QO_NAME) = aExtraQuestions(i, QO_ALTERNATE_ID)
                    End If
                Next

            End If

        End If

    End If

    Set oConfigQuestion = Nothing
    Set oMissingObjectsDOM = Nothing
    Set oObject = Nothing
    Set oQuestions = Nothing
    Set oQuestionsDOM = Nothing
    Set oSubsSetDOM = Nothing
    Erase aMissingObjectsIds

    getQuestionsConfigFromMD = lErr
    Err.Clear

End Function

Function GetQuestionsConfigFromCache(aSvcConfigInfo, sCacheXML, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
'********************************************************
'*Purpose: Return the configuration of the questions of a subscription set from a Cache system
'*Inputs:  aSvcConfigInfo: the Info array for Service Conffig
'*         sCacheXML: The XML Cache.
'*Outputs: aNormalQuestions: An array with info for normal (non-slicing) questions
'*         aSlicingQuestions: An array with info for slicing questions.
'*         aExtraQuestions: An array with info for questions that are not defined on the Project Repository.
'********************************************************
Const PROCEDURE_NAME = "GetQuestionsConfigFromCache"
Dim lErr

Dim oDOM
Dim oSubsSet
Dim oQuestions
Dim oQuestion

Dim lCount
Dim i

    On Error Resume next
    lErr = NO_ERR


    'Get DOM object:
    If lErr = NO_ERR Then
        lErr = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oDOM)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sCacheXML", LogLevelTrace)
    End If

    'First, get the subscription set properties:
    If lErr = NO_ERR Then
        Set oSubsSet = oDOM.selectSingleNode("/mi/in/oi")

        aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) = oSubsSet.getAttribute("id")
        aSvcConfigInfo(SVCCFG_SS_NAME) = GetPropertyValue(oSubsSet, PROP_NAME)
        aSvcConfigInfo(SVCCFG_SS_MAP_ID) = GetPropertyValue(oSubsSet, PROP_STORAGE_MAPPING_ID)

    End If

    'Get Normal questions:
    If lErr = NO_ERR Then
        Set oQuestions = oDOM.selectNodes("//oi[@tp='normal']")
        lCount = oQuestions.length

        If lCount > 0 Then
            Redim aNormalQuestions(lCount - 1, MAX_QO_INFO)

            For i = 0 To lCount - 1
                Set oQuestion = oQuestions(i)
                aNormalQuestions(i, QO_ID) = oQuestion.getAttribute("id")
                aNormalQuestions(i, QO_NAME) = GetPropertyValue(oQuestion, PROP_NAME)
                aNormalQuestions(i, QO_DESCRIPTION) = GetPropertyValue(oQuestion, PROP_DESC)
                aNormalQuestions(i, QO_VALUE) = GetPropertyValue(oQuestion, PROP_VALUE)
                aNormalQuestions(i, QO_ALTERNATE_ID) = GetPropertyValue(oQuestion, PROP_ALTERNATE_QUESTION)
                aNormalQuestions(i, QO_PROMPT_COUNT) = GetPropertyValue(oQuestion, PROP_PROMPT_COUNT)
                aNormalQuestions(i, QO_IS_ID) = GetPropertyValue(oQuestion, PROP_IS_ID)
            Next
        End If
    End If


    'Get slicing questions:
    'Get Normal questions:
    If lErr = NO_ERR Then
        Set oQuestions = oDOM.selectNodes("//oi[@tp='slicing']")
        lCount = oQuestions.length

        If lCount > 0 Then
            Redim aSlicingQuestions(lCount - 1, MAX_QO_INFO)

            For i = 0 To lCount - 1
                Set oQuestion = oQuestions(i)

                aSlicingQuestions(i, QO_ID) = oQuestion.getAttribute("id")
                aSlicingQuestions(i, QO_NAME) = GetPropertyValue(oQuestion, PROP_NAME)
                aSlicingQuestions(i, QO_DESCRIPTION) = GetPropertyValue(oQuestion, PROP_DESC)
                aSlicingQuestions(i, QO_VALUE) = GetPropertyValue(oQuestion, PROP_VALUE)
                aSlicingQuestions(i, QO_ALTERNATE_ID) = GetPropertyValue(oQuestion, PROP_ALTERNATE_QUESTION)
                aSlicingQuestions(i, QO_MAP_ID) = GetPropertyValue(oQuestion, PROP_STORAGE_MAPPING_ID)
                aSlicingQuestions(i, QO_PROMPT_COUNT) = GetPropertyValue(oQuestion, PROP_PROMPT_COUNT)
                aSlicingQuestions(i, QO_IS_ID) = GetPropertyValue(oQuestion, PROP_IS_ID)
            Next
        End If
    End If

    'Get extra questions:
    If lErr = NO_ERR Then

        Set oQuestions = oDOM.selectNodes("//oi[@tp='custom']")
        lCount = oQuestions.length

        If lCount > 0 Then
            Redim aExtraQuestions(lCount - 1, MAX_QO_INFO)

            For i = 0 To lCount - 1
                Set oQuestion = oQuestions(i)

                aExtraQuestions(i, QO_ID) = NEW_OBJECT_ID
                aExtraQuestions(i, QO_NAME) = GetPropertyValue(oQuestion, PROP_NAME)
                aExtraQuestions(i, QO_DESCRIPTION) = GetPropertyValue(oQuestion, PROP_DESC)
                aExtraQuestions(i, QO_VALUE) = GetPropertyValue(oQuestion, PROP_VALUE)
                aExtraQuestions(i, QO_ALTERNATE_ID) = GetPropertyValue(oQuestion, PROP_ALTERNATE_QUESTION)
                aExtraQuestions(i, QO_ALTERNATE_NAME) = GetPropertyValue(oQuestion, PROP_ALTERNATE_NAME)
                aExtraQuestions(i, QO_MAP_ID) = GetPropertyValue(oQuestion, PROP_STORAGE_MAPPING_ID)
                aExtraQuestions(i, QO_PROMPT_COUNT) = GetPropertyValue(oQuestion, PROP_PROMPT_COUNT)
                aExtraQuestions(i, QO_IS_ID) = GetPropertyValue(oQuestion, PROP_IS_ID)
            Next
        End If
    End If

    Set oQuestion = Nothing
    Set oQuestions = Nothing
    Set oDOM = Nothing
    Set oSubsSet = Nothing

    GetQuestionsConfigFromCache = lErr
    Err.Clear

End Function

Function GetQuestionDescription(oQuestionDOM)
'********************************************************
'*Purpose: Return the description associated with a question object.
'           Instead of showing the description field, we'll show the list of
'           publication of the given Question
'*Inputs:  oQuestionDOM: The DOM object of the Question
'*Outputs: returns the name, not an error number (no API calls to return an error)
'********************************************************
Const PROCEDURE_NAME = "getQuestionDescription"
Dim sDesc

Dim oPublications
Dim oPublication

    On Error Resume next

    'Use the ID as default
    sDesc = ""

    Set oPublications = oQuestionDOM.selectNodes("mi/in/oi[@tp='" & TYPE_PUBLICATION & "']")
    lCount = oPublications.length

    If lCount > 0 Then

        sDesc = ""
        For Each oPublication in oPublications
            sDesc = sDesc & oPublication.getAttribute("n") & ", "
        Next

        sDesc = Left(sDesc, Len(sDesc) - 2)

    End If

    Set oPublication = Nothing
    Set oPublications = Nothing

    getQuestionDescription = sDesc
    Err.Clear

End Function

Function GetSvcConfigCacheName(aSvcConfigInfo)
'********************************************************
'*Purpose: Return the name of the cache file associated with a given SvcConfig Info
'*Inputs:  aSvcConfigInfo: The array with the necessary information
'*Outputs: returns the name of the file, not an error number (no API calls to return an error)
'********************************************************

    GetSvcConfigCacheName = "cfg" & aSvcConfigInfo(SVCCFG_SVC_ID) & "_" & aSvcConfigInfo(SVCCFG_SS_ID)

End Function


Function SaveSubscriptionSetConfig(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
'********************************************************
'*Purpose: Saves the configuration of the subscription sets into the MD
'*Inputs:  aSvcConfigInfo: the Info array for Service Conffig
'*         aNormalQuestions: An array with info for normal (non-slicing) questions
'*         aSlicingQuestions: An array with info for slicing questions.
'*         aExtraQuestions: An array with info for questions that are not defined on the Project Repository.
'********************************************************
Const PROCEDURE_NAME = "SaveSubscriptionSetConfig"
Dim lErr

Dim sPropsXML
Dim i
Dim lCount
Dim lExtraIndex
Dim sSSId

Dim sSiteId

    On Error Resume Next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")
    If aSvcConfigInfo(SVCCFG_SS_ID) = STATIC_SS Then
        sSSId = DEFAULT_STATIC_SS_ID
    Elseif aSvcConfigInfo(SVCCFG_SS_ID) = DYNAMIC_SS Then
        sSSId = DEFAULT_DYNAMIC_SS_ID
    Else
        sSSId = aSvcConfigInfo(SVCCFG_SS_ID)
    End If

    If lErr = NO_ERR Then
        sPropsXML = ""
        sPropsXML = sPropsXML & "<mi><in><oi tp=""" & TYPE_SUBSSET_CONFIG & """  >"
        sPropsXML = sPropsXML & " <prs>"
        sPropsXML = sPropsXML & "  <pr id=""" & PROP_NAME & """  v=""" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SS_NAME)) & """  />"
        sPropsXML = sPropsXML & "  <pr id=""" & PROP_PHYSICAL_ID & """  v=""" & sSSId & """  />"

        'Save the correct Physical ID:
        'FOR DYNAMIC subscriptin sets, save also the Mapping Id:
        If aSvcConfigInfo(SVCCFG_STEP) = DYNAMIC_SS Then sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID & """  v=""" & aSvcConfigInfo(SVCCFG_SS_MAP_ID) & """  />"

        sPropsXML = sPropsXML & " </prs>"
        sPropsXML = sPropsXML & " <mi><in>"


        'Add normal questions:
        If Not IsEmpty(aNormalQuestions) Then
            lCount = UBound(aNormalQuestions)

            For i = 0 To lCount
                sPropsXML = sPropsXML & "<oi tp=""" & TYPE_QUESTION_CONFIG & """  id=""" & GetGUID() & """ >"
                sPropsXML = sPropsXML & " <prs>"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_NAME & """  v=""" & Server.HTMLEncode(aNormalQuestions(i, QO_NAME)) & """  />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_PHYSICAL_ID & """  v=""" & aNormalQuestions(i, QO_ID) & """  />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID & """  v=""" & aNormalQuestions(i, QO_MAP_ID) & """  />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_QUESTION_TYPE & """  v=""0"" />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_ALTERNATE_QUESTION & """  v=""" & aNormalQuestions(i, QO_ALTERNATE_ID) & """  />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_SLICED_BY & """  v="""" />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORE_IN_SBR & """  v="""" />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_PROMPT_COUNT & """  v="""" />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_ID & """  v="""" />"

                If aNormalQuestions(i, QO_VALUE) = "false" Then
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_SHOWN & """  v=""false"" />"
                Else
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_SHOWN & """  v=""true"" />"
                End If

                sPropsXML = sPropsXML & " </prs>"
                sPropsXML = sPropsXML & "</oi>"
            Next
        End If

        'Add slicing questions:
        If Not IsEmpty(aSlicingQuestions) Then
            lCount = UBound(aSlicingQuestions)

            For i = 0 To lCount
                sPropsXML = sPropsXML & "<oi tp=""" & TYPE_QUESTION_CONFIG & """  id=""" & GetGUID() & """ >"
                sPropsXML = sPropsXML & " <prs>"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_NAME & """  v=""" & Server.HTMLEncode(aSlicingQuestions(i, QO_NAME)) & """  />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_PHYSICAL_ID & """  v=""" & aSlicingQuestions(i, QO_ID) & """  />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_QUESTION_TYPE & """  v=""2"" />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_PROMPT_COUNT & """  v="""" />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_ID & """  v="""" />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_SHOWN & """  v="""" />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID & """  v="""" />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORE_IN_SBR & """  v="""" />"

                If aSlicingQuestions(i, QO_VALUE) = ANSWER_OTHER_ID Then
                    lExtraIndex = GetExtraQuestionIndex(aExtraQuestions, aSlicingQuestions(i, QO_ALTERNATE_ID))
                    If aExtraQuestions(lExtraIndex, QO_MAP_ID) = "sbr" Then
                        sPropsXML = sPropsXML & "  <pr id=""" & PROP_SLICED_BY & """  v=""qo." & aSlicingQuestions(i, QO_ALTERNATE_ID) & "." & ANSWER_PROMPT_ANSWER & ".1"" />"
                    Else
                        sPropsXML = sPropsXML & "  <pr id=""" & PROP_SLICED_BY & """  v=""qo." & aSlicingQuestions(i, QO_ALTERNATE_ID) & "." & ANSWER_PREFERENCE_ID & """  />"
                    End If

                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_ALTERNATE_QUESTION & """  v=""" & aSlicingQuestions(i, QO_ALTERNATE_ID) & """ />"

                ElseIf aSlicingQuestions(i, QO_VALUE) = ANSWER_DEFAULT Then
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_SLICED_BY & """  v="""" />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_ALTERNATE_QUESTION & """  v="""" />"
                Else
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_SLICED_BY & """  v=""" & Server.HTMLEncode(aSlicingQuestions(i, QO_VALUE)) & """ />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_ALTERNATE_QUESTION & """  v="""" />"
                End If

                sPropsXML = sPropsXML & " </prs>"
                sPropsXML = sPropsXML & "</oi>"

                If aSlicingQuestions(i, QO_VALUE) = ANSWER_OTHER_ID Then
                    lExtraIndex = GetExtraQuestionIndex(aExtraQuestions, aSlicingQuestions(i, QO_ALTERNATE_ID))

                    sPropsXML = sPropsXML & "<oi tp=""" & TYPE_QUESTION_CONFIG & """ id=""" & GetGUID() & """ >"
                    sPropsXML = sPropsXML & " <prs>"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_NAME & """ v=""" & Server.HTMLEncode(aExtraQuestions(lExtraIndex, QO_NAME)) & """ />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_PHYSICAL_ID & """ v=""" & aExtraQuestions(lExtraIndex, QO_ALTERNATE_ID) & """ />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_QUESTION_TYPE & """ v=""1"" />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_PROMPT_COUNT & """ v=""" & aExtraQuestions(lExtraIndex, QO_PROMPT_COUNT) & """ />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_ID & """ v=""" & aExtraQuestions(lExtraIndex, QO_IS_ID) & """ />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_ALTERNATE & """ v=""true"" />"

                    If aExtraQuestions(lExtraIndex, QO_MAP_ID) = "sbr" Then
                        sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID & """ v="""" />"
                    Else
                        sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID & """ v=""" & aExtraQuestions(lExtraIndex, QO_MAP_ID) & """ />"
                    End If

                    sPropsXML = sPropsXML & " </prs>"
                    sPropsXML = sPropsXML & "</oi>"
                End If

             Next
        End If

        'Add Extra questions:
        If Not IsEmpty(aExtraQuestions) Then
            lCount = UBound(aExtraQuestions)

            For i = 0 To lCount
                If aExtraQuestions(i, QO_VALUE) = "false" Then
                    sPropsXML = sPropsXML & "<oi tp=""" & TYPE_QUESTION_CONFIG & """ id=""" & GetGUID() & """ >"
                    sPropsXML = sPropsXML & " <prs>"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_NAME & """ v=""" & Server.HTMLEncode(aExtraQuestions(i, QO_NAME)) & """ />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_PHYSICAL_ID & """ v=""" & aExtraQuestions(i, QO_ALTERNATE_ID) & """ />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_QUESTION_TYPE & """ v=""1"" />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_PROMPT_COUNT & """ v=""" & aExtraQuestions(i, QO_PROMPT_COUNT) & """ />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_ID & """ v=""" & aExtraQuestions(i, QO_IS_ID) & """ />"
                    sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_ALTERNATE & """ v=""false"" />"

                    If aExtraQuestions(i, QO_MAP_ID) = "sbr" Then
                        sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID & """ v="""" />"
                    Else
                        sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID & """ v=""" & aExtraQuestions(i, QO_MAP_ID) & """ />"
                    End If

                    sPropsXML = sPropsXML & " </prs>"
                    sPropsXML = sPropsXML & "</oi>"
                End If
            Next
        End If

        sPropsXML = sPropsXML & "  </in></mi>"
        sPropsXML = sPropsXML & "</oi></in></mi>"

    End If

    'Get current service config:
    If lErr = NO_ERR Then
        If Len(aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID)) = 0 Then
            lErr = getConfigObjectID(aSvcConfigInfo(SVCCFG_SVC_ID), TYPE_SERVICE, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID))
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getConfigObjectID", LogLevelTrace)
        End If

        'If no configuration object for this service, create one by assign it the default answer:
        If lErr = NO_ERR Then
            If aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID) = NEW_OBJECT_ID Then
                aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID) = GetGUID()
                lErr = co_createObject(sSiteId, sSiteId, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID), GenerateServiceConfigXML(aSvcConfigInfo, ANSWER_DEFAULT))
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_createObject", LogLevelTrace)
            End If
        End If
    End If

    'Get current subscription set config:
    If lErr = NO_ERR Then
        If Len(aSvcConfigInfo(SVCCFG_SS_CONFIG_ID)) = 0 Then
            lErr = getConfigObjectID(sSSId, TYPE_SUBSSET, aSvcConfigInfo(SVCCFG_SS_CONFIG_ID))
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getConfigObjectID", LogLevelTrace)
        End If
    End If

    'If editing, simply delete the old one, since we'll regenerate the whole Object XML
    If lErr = NO_ERR Then
        If aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) <> NEW_OBJECT_ID Then
            lErr = co_deleteObject(sSiteId, aSvcConfigInfo(SVCCFG_SS_CONFIG_ID))
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_deleteObject", LogLevelTrace)
        End If
    End If

    'Save into MD
    If lErr = NO_ERR Then
        aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) = GetGUID()
        lErr = co_createObject(sSiteId, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID), aSvcConfigInfo(SVCCFG_SS_CONFIG_ID), sPropsXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_deleteObject", LogLevelTrace)
    End If

	'Generate the SQL in the webserver and reset the engine.
    If lErr = NO_ERR Then
        lErr = co_generateSubscriptionSetSQL(sSiteId, aSvcConfigInfo(SVCCFG_SVC_ID), sSSId)
    End If

	If lErr = NO_ERR Then
	    lErr = ResetSubscriptionEngine()
	End If

    SaveSubscriptionSetConfig = lErr
    Err.Clear

End Function


Function SaveQuestionConfig(aSvcConfigInfo, sISID, lPromptCount)
'********************************************************
'*Purpose: Saves the configuration of the question object into the cache file.
'*          The configuration is stored in MD until the whole subscription set is saved.
'*Inputs:  aSvcConfigInfo: the Info array for Service Conffig
'********************************************************
Const PROCEDURE_NAME = "SaveQuestionConfig"
Dim lErr

Dim sCacheXML
Dim sCacheName

Dim oDOM
Dim oQuestion
Dim oProps
Dim oProp
Dim oExtras

    On Error Resume Next
    lErr = NO_ERR

    sCacheName = GetSvcConfigCacheName(aSvcConfigInfo)

    If lErr = NO_ERR Then
        lErr = ReadCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, sCacheXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ReadCache", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sCacheXML", LogLevelTrace)
        End If
    End If

    If lErr = NO_ERR Then
        Set oExtras = oDOM.selectSingleNode("//extras")
        Set oQuestion = oExtras.selectSingleNode("oi[prs/pr[@v='" & aSvcConfigInfo(SVCCFG_AQ_ID) &"']]")

        'If there was no extra question, create a new one:
        If oQuestion Is Nothing Then
            Set oQuestion = oExtras.appendChild(oDOM.createElement("oi"))
            oQuestion.setAttribute "id", NEW_OBJECT_ID
            oQuestion.setAttribute "tp", "custom"

            Set oProps = oQuestion.appendChild(oDOM.createElement("prs"))

            Set oProp = oProps.appendChild(oDOM.createElement("pr"))
            oProp.setAttribute "id", PROP_NAME
            oProp.setAttribute "v", aSvcConfigInfo(SVCCFG_AQ_NAME)

            Set oProp = oProps.appendChild(oDOM.createElement("pr"))
            oProp.setAttribute "id", PROP_DESC
            oProp.setAttribute "v", ""

            Set oProp = oProps.appendChild(oDOM.createElement("pr"))
            oProp.setAttribute "id", PROP_ALTERNATE_QUESTION
            oProp.setAttribute "v", aSvcConfigInfo(SVCCFG_AQ_ID)

            Set oProp = oProps.appendChild(oDOM.createElement("pr"))
            oProp.setAttribute "id", PROP_STORAGE_MAPPING_ID
            oProp.setAttribute "v", ""

            Set oProp = oProps.appendChild(oDOM.createElement("pr"))
            oProp.setAttribute "id", PROP_PROMPT_COUNT
            oProp.setAttribute "v", lPromptCount

            Set oProp = oProps.appendChild(oDOM.createElement("pr"))
            oProp.setAttribute "id", PROP_IS_ID
            oProp.setAttribute "v", sISID

            Set oProp = oProps.appendChild(oDOM.createElement("pr"))
            oProp.setAttribute "id", PROP_VALUE
            If aSvcConfigInfo(SVCCFG_QO_ID) = NEW_OBJECT_ID Then
                oProp.setAttribute "v", "false"
            Else
                oProp.setAttribute "v", "true"
            End If

        End If

        'If searching for an alternate question, set the new value to the original question:
        If aSvcConfigInfo(SVCCFG_QO_ID) <> NEW_OBJECT_ID Then
            Set oQuestion = oDOM.selectSingleNode("//oi[@id='" & aSvcConfigInfo(SVCCFG_QO_ID) & "']")
            Set oProp = oQuestion.selectSingleNode("prs/pr[@id='" & PROP_ALTERNATE_QUESTION & "']")
            oProp.Attributes.getNamedItem("v").Text = aSvcConfigInfo(SVCCFG_AQ_ID)
        End If

    End If

    If lErr = NO_ERR Then
        lErr = WriteCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, oDOM.xml)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling WriteCache", LogLevelTrace)
    End If

    Set oDOM = Nothing
    Set oExtras = Nothing
    Set oProp = Nothing
    Set oProps = Nothing
    Set oQuestion = Nothing

    SaveQuestionConfig = lErr
    Err.Clear

End Function

Function SaveMapConfig(aSvcConfigInfo, aMapInfo)
'********************************************************
'*Purpose: Saves the configuration of the Mapping Storage into the cache file.
'*          The configuration is stored in MD until the whole subscription set is saved.
'*Inputs:  aSvcConfigInfo: the Info array for Service Conffig
'********************************************************
Const PROCEDURE_NAME = "SaveMapConfig"
Dim lErr

Dim sCacheXML
Dim sCacheName

Dim oDOM
Dim oQuestion
Dim oSubsSet
Dim oProp

    On Error Resume Next
    lErr = NO_ERR

    sCacheName = GetSvcConfigCacheName(aSvcConfigInfo)

    If lErr = NO_ERR Then
        lErr = ReadCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, sCacheXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ReadCache", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sCacheXML", LogLevelTrace)
        End If
    End If

    If lErr = NO_ERR Then
        If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then
            If Len(aMapInfo(MAP_NAME)) = 0 Then aMapInfo(MAP_NAME) = GetStorageMapName(aMapInfo(MAP_ID))

            Set oQuestion = oDOM.selectSingleNode("//extras/oi[prs/pr[@v='" & aSvcConfigInfo(SVCCFG_AQ_ID) & "']]")

            Set oProp = oQuestion.selectSingleNode("prs/pr[@id='" & PROP_STORAGE_MAPPING_ID & "']")
            oProp.Attributes.getNamedItem("v").Text = aMapInfo(MAP_ID)

            Set oProp = oQuestion.selectSingleNode("prs/pr[@id='" & PROP_DESC & "']")
            oProp.Attributes.getNamedItem("v").Text = aMapInfo(MAP_NAME)
        Else
            Set oSubsSet = oDOM.selectSingleNode("mi/in/oi")

            Set oProp = oSubsSet.selectSingleNode("prs/pr[@id='" & PROP_STORAGE_MAPPING_ID & "']")
            oProp.Attributes.getNamedItem("v").Text = aMapInfo(MAP_ID)
        End If
    End If

    If lErr = NO_ERR Then
        lErr = WriteCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, oDOM.xml)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling WriteCache", LogLevelTrace)
    End If

    Set oDOM = Nothing
    Set oQuestion = Nothing
    Set oSubsSet = Nothing
    Set oProp = Nothing

    SaveMapConfig = lErr
    Err.Clear

End Function


Function CreateSubscriptionsSetCache(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
'********************************************************
'*Purpose: Saves the configuration of the subscription sets into a Cache File to retrieve it later
'*Inputs:  aSvcConfigInfo: the Info array for Service Conffig
'*         aNormalQuestions: An array with info for normal (non-slicing) questions
'*         aSlicingQuestions: An array with info for slicing questions.
'*         aExtraQuestions: An array with info for questions that are not defined on the Project Repository.
'********************************************************
Const PROCEDURE_NAME = "CreateSubscriptionsSetCache"
Dim lErr

Dim sCacheName
Dim sCacheXML
Dim oCacheDOM
Dim oMaps

Dim sPropsXML
Dim i
Dim lCount

    On Error Resume Next
    lErr = NO_ERR

    If lErr = NO_ERR Then
        sPropsXML = ""
        sPropsXML = sPropsXML & "<mi><in><oi tp=""" & TYPE_SUBSSET_CONFIG & """ id=""" & aSvcConfigInfo(SVCCFG_SS_CONFIG_ID) & """>"
        sPropsXML = sPropsXML & " <prs>"
        sPropsXML = sPropsXML & "  <pr id=""" & PROP_PHYSICAL_ID &        """ v=""" & aSvcConfigInfo(SVCCFG_SS_ID) & """ />"
        sPropsXML = sPropsXML & "  <pr id=""" & PROP_NAME &               """ v=""" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SS_NAME)) & """ />"
        sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID & """ v=""" & aSvcConfigInfo(SVCCFG_SS_MAP_ID) & """ />"
        sPropsXML = sPropsXML & " </prs>"

        'Add normal questions:
        sPropsXML = sPropsXML & " <normal>"
        If Not IsEmpty(aNormalQuestions) Then
            lCount = UBound(aNormalQuestions)
            For i = 0 To lCount
                sPropsXML = sPropsXML & "<oi tp=""normal"" id=""" & aNormalQuestions(i, QO_ID) & """>"
                sPropsXML = sPropsXML & " <prs>"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_NAME &                  """ v=""" & Server.HTMLEncode(aNormalQuestions(i, QO_NAME)) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_DESC &                  """ v=""" & Server.HTMLEncode(aNormalQuestions(i, QO_DESCRIPTION)) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_VALUE &                 """ v=""" & Server.HTMLEncode(aNormalQuestions(i, QO_VALUE)) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_ALTERNATE_QUESTION &    """ v=""" & aNormalQuestions(i, QO_ALTERNATE_ID) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID &    """ v=""" & aNormalQuestions(i, QO_MAP_ID) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_ID &                 """ v=""" & aNormalQuestions(i, QO_IS_ID) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_PROMPT_COUNT &          """ v=""" & aNormalQuestions(i, QO_PROMPT_COUNT) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_QUESTION_TYPE &         """ v=""0"" />"
                sPropsXML = sPropsXML & " </prs>"
                sPropsXML = sPropsXML & "</oi>"
            Next
        End If
        sPropsXML = sPropsXML & " </normal>"

        'Add slicing questions:
        sPropsXML = sPropsXML & " <slicing>"
        If Not IsEmpty(aSlicingQuestions) Then
            lCount = UBound(aSlicingQuestions)
            For i = 0 To lCount
                sPropsXML = sPropsXML & "<oi tp=""slicing"" id='" & aSlicingQuestions(i, QO_ID) & "'>"
                sPropsXML = sPropsXML & " <prs>"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_NAME &                  """ v=""" & Server.HTMLEncode(aSlicingQuestions(i, QO_NAME)) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_DESC &                  """ v=""" & Server.HTMLEncode(aSlicingQuestions(i, QO_DESCRIPTION)) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_VALUE &                 """ v=""" & Server.HTMLEncode(aSlicingQuestions(i, QO_VALUE)) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_ALTERNATE_QUESTION &    """ v=""" & aSlicingQuestions(i, QO_ALTERNATE_ID) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID &    """ v=""" & aSlicingQuestions(i, QO_MAP_ID) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_ID &                 """ v=""" & aSlicingQuestions(i, QO_IS_ID) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_PROMPT_COUNT &          """ v=""" & aSlicingQuestions(i, QO_PROMPT_COUNT) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_QUESTION_TYPE &         """ v=""1"" />"
                sPropsXML = sPropsXML & " </prs>"
                sPropsXML = sPropsXML & "</oi>"
            Next
        End If
        sPropsXML = sPropsXML & " </slicing>"

        'Add Extra questions:
        sPropsXML = sPropsXML & " <extras>"
        If Not IsEmpty(aExtraQuestions) Then
            lCount = UBound(aExtraQuestions)
            For i = 0 To lCount
                sPropsXML = sPropsXML & "<oi tp=""custom"" id=""" & aExtraQuestions(i, QO_ALTERNATE_ID) & """ >"
                sPropsXML = sPropsXML & " <prs>"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_NAME &                  """ v=""" & Server.HTMLEncode(aExtraQuestions(i, QO_NAME)) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_DESC &                  """ v=""" & Server.HTMLEncode(aExtraQuestions(i, QO_DESCRIPTION)) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_VALUE &                 """ v=""" & Server.HTMLEncode(aExtraQuestions(i, QO_VALUE)) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_ALTERNATE_QUESTION &    """ v=""" & aExtraQuestions(i, QO_ALTERNATE_ID) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_STORAGE_MAPPING_ID &    """ v=""" & aExtraQuestions(i, QO_MAP_ID) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_IS_ID &                 """ v=""" & aExtraQuestions(i, QO_IS_ID) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_PROMPT_COUNT &          """ v=""" & aExtraQuestions(i, QO_PROMPT_COUNT) & """ />"
                sPropsXML = sPropsXML & "  <pr id=""" & PROP_QUESTION_TYPE &         """ v=""2"" />"
                sPropsXML = sPropsXML & " </prs>"
                sPropsXML = sPropsXML & "</oi>"
            Next
        End If
        sPropsXML = sPropsXML & " </extras></oi>"

        'Add the maps:
        sPropsXML = sPropsXML & "<maps>"

        sCacheName = GetSvcConfigCacheName(aSvcConfigInfo)
        lErr = ReadCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, sCacheXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ReadCache", LogLevelTrace)

        If lErr = NO_ERR Then
            If Len(sCacheXML) > 0 Then
                lErr = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
                If lErr <> NO_ERR Then
                    Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sCacheXML", LogLevelTrace)
                Else
                    Set oMaps = oCacheDOM.selectNodes("//oi[@tp='" & TYPE_STORAGE_MAPPING & "']")
                    lCount = oMaps.length - 1

                    For i = 0 To lCount
                        sPropsXML = sPropsXML & oMaps(i).xml
                    Next
                End If
            End If
        End If

        sPropsXML = sPropsXML & "</maps>"

        sPropsXML = sPropsXML & "</in></mi>"

    End If

    'Save the XML into the cache:
    If lErr = NO_ERR Then
        lErr = WriteCache(GetSvcConfigCacheName(aSvcConfigInfo), SVC_CONFIG_CACHE_FOLDER, sPropsXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling WriteCache", LogLevelTrace)
    End If

    CreateSubscriptionsSetCache = lErr
    Err.Clear

End Function

Function GenerateServiceConfigXML(aSvcConfigInfo, sAnswer)
'********************************************************
'*Purpose: Return the XML of for a ServiceConfig object
'*Inputs:  aSvcConfigInfo: The array with the necessary information
'*Outputs: returns the XML, not an error number (no API calls to return an error)
'********************************************************
Dim sObjectPropsXML

    On Error Resume Next

    sObjectPropsXML = ""
    sObjectPropsXML = sObjectPropsXML & "<mi><in><oi tp=""" & TYPE_SERVICE_CONFIG & """><prs>"
    sObjectPropsXML = sObjectPropsXML & " <pr id=""NAME"" v=""" & Server.HTMLEncode(aSvcConfigInfo(SVCCFG_SVC_NAME)) & """ />"
    sObjectPropsXML = sObjectPropsXML & " <pr id=""" & PROP_PHYSICAL_ID &    """ v=""" & aSvcConfigInfo(SVCCFG_SVC_ID) & """ />"
    sObjectPropsXML = sObjectPropsXML & " <pr id=""" & PROP_DEFAULT_ANSWER & """ v=""" & sAnswer & """ />"
    sObjectPropsXML = sObjectPropsXML & "</prs></oi></in></mi>"

    GenerateServiceConfigXML = sObjectPropsXML
    Err.Clear

End Function

Function GetMapsForQuestion(aSvcConfigInfo, aMapList)
'********************************************************
'*Purpose: Return the maps for the custom question of the subsset.
'           The Maps may come from MD or from the cache file.
'*Inputs:  aSvcConfig: The array with information of the service.
'*Outputs: aMapList:  Returns the information of the maps found
'********************************************************
Const PROCEDURE_NAME = "GetMapsForQuestion"
Dim lErr
Dim lCount

Dim sSiteId
Dim sMapListXML
Dim oMapListDOM
Dim oMap
Dim oMaps
Dim aMapInfo
Dim aList()

Dim sCacheName
Dim sCacheXML
Dim oCacheDOM

Dim i, j

    On Error Resume Next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    If lErr = NO_ERR Then

        lErr = co_getMappingObjects(sSiteId, aSvcConfigInfo(SVCCFG_AQ_ID), sMapListXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getMappingObjects", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sMapListXML, oMapListDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sMapListXML", LogLevelTrace)
        End If

    End If

    If lErr = NO_ERR Then
        Set oMaps = oMapListDOM.selectNodes("//oi[@tp='" & TYPE_STORAGE_MAPPING & "']")

        If oMaps.length > 0 Then
            lCount = lCount + oMaps.length

            Redim Preserve aList(lCount - 1, MAX_MAP_INFO)

            For j = 0 To oMaps.length - 1
                Set oMap = oMaps(j)

                lErr = GetMapInfo(oMap.getAttribute("id"), aMapInfo)
                If lErr <> NO_ERR Then
                    Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sMapListXML", LogLevelTrace)
                    Exit For
                End If

                aList(i, MAP_ID) = oMap.getAttribute("id")
                aList(i, MAP_DESC) = aMapInfo(MAP_DESC)
                aList(i, MAP_NAME) = aMapInfo(MAP_NAME)
                aList(i, MAP_DBALIAS) = aMapInfo(MAP_DBALIAS)
                aList(i, MAP_FILTER) = aMapInfo(MAP_FILTER)
                i = i + 1

            Next
        End If
    End If

    'Get Maps from cache:
    If lErr = NO_ERR Then
        sCacheName = GetSvcConfigCacheName(aSvcConfigInfo)
        lErr = ReadCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, sCacheXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ReadCache", LogLevelTrace)
    End If

    If lErr = NO_ERR Then
        If Len(sCacheXML) > 0 Then
            lErr = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sMapListXML", LogLevelTrace)
            Else
                Set oMaps = oCacheDOM.selectNodes("//oi[@tp='" & TYPE_STORAGE_MAPPING & "']")

                If oMaps.length > 0 Then
                    lCount = lCount + oMaps.length

                    Redim Preserve aList(lCount - 1, MAX_MAP_INFO)

                    For j = 0 To oMaps.length - 1
                        Set oMap = oMaps(j)

                        lErr = GetMapInfo(oMap.getAttribute("id"), aMapInfo)
                        If lErr <> NO_ERR Then
                            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sMapListXML", LogLevelTrace)
                            Exit For
                        End If

                        aList(i, MAP_ID) = oMap.getAttribute("id")
                        aList(i, MAP_DESC) = aMapInfo(MAP_DESC)
                        aList(i, MAP_NAME) = aMapInfo(MAP_NAME)
                        aList(i, MAP_DBALIAS) = aMapInfo(MAP_DBALIAS)
                        aList(i, MAP_FILTER) = aMapInfo(MAP_FILTER)
                        i = i + 1

                    Next
                End If
            End If
        End If
    End If

    If lErr = NO_ERR Then
        If lCount > 0 Then
            Redim aMapList(lCount - 1, MAX_MAP_INFO)

            For i = 0 To lCount - 1
                aMapList(i, MAP_ID) = aList(i, MAP_ID)
                aMapList(i, MAP_DBALIAS) = aList(i, MAP_DBALIAS)
                aMapList(i, MAP_NAME) = aList(i, MAP_NAME)
                aMapList(i, MAP_DESC) = aList(i, MAP_DESC)
                aMapList(i, MAP_FILTER) = aList(i, MAP_FILTER)
            Next
        End If
    End If

    Set oCacheDOM = Nothing
    Set oMap = Nothing
    Set oMapListDOM = Nothing
    Set oMaps = Nothing

    GetMapsForQuestion = lErr
    Err.Clear

End Function

Function GetPromptsInfo(aConnectionInfo, sQOID, sQODetailsXML, aPromptInfo)
'***********************************************************************************************
'Purpose: Load Prompts info for a QO
'Inputs:  aConnectionInfo, sQOID, sQODetailsXML
'Outputs: aPromptInfo
'***********************************************************************************************
Dim oDetailDOM
Dim sInfoSourceID
Dim oCurrQO
Dim oISOI
Dim sReportID
Dim oSession
Dim oReport
Dim oAllPrompts
Dim sProjectID
Dim oInfoSource
Dim i
Dim j
Dim oSinglePrompt
Dim lErrNumber

    On Error Resume Next
    lErrNumber = NO_ERR

	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, Replace(Replace(Replace(sQODetailsXML, "&gt;", ">"), "&lt;", "<"), "&amp;", "&"), oDetailDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PromptCuLib.asp", "ReadPromptQuestionFromCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		set oCurrQO = oDetailDOM.selectSingleNode("/mi/qos/oi[@tp='" & TYPE_QUESTION & "' $and$ @id='" & sQOID & "']")
		sInfoSourceID = oCurrQO.getAttribute("isid")
		set oISOI = oDetailDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_INFORMATION_SOURCE & "' $and$ @id='" & sInfoSourceID & "']")
		set oInfoSource = oISOI.selectSingleNode("prs/pr[@id='ISM_connInfo']/info_source_props")
		lErrNumber = Err.number
	End If

	If lErrNumber = NO_ERR Then
		aConnectionInfo(S_SERVER_NAME_CONNECTION) = oInfoSource.SelectSingleNode("server/primary").getAttribute("name")
		aConnectionInfo(N_PORT_CONNECTION) = CLng(oInfoSource.SelectSingleNode("server/primary").getAttribute("port"))
		Call GetDSSSession(aConnectionInfo, oSession, sErrDescription)

		sProjectID = oInfoSource.SelectSingleNode("project").getAttribute("id")
		lErrNumber = MapProjectIDToName(oSession, sProjectID, aConnectionInfo(S_PROJECT_CONNECTION))
		If lErrNumber <> NO_ERR then
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErrDescription), "PromptCuLib.asp", "GetHydraPrompt", "", "Error calling MapProjectIDToName", LogLevelTrace)
		Else
			sReportID = oCurrQO.selectSingleNode("./prs/pr[@n='definition']/Question_Object_Definition").getAttribute("id")
			aConnectionInfo(S_UID_CONNECTION) = oInfoSource.SelectSingleNode("server/login").getAttribute("name")
			aConnectionInfo(S_PWD_CONNECTION) = Decrypt(oInfoSource.SelectSingleNode("server/login").getAttribute("pwd"))
			aConnectionInfo(S_TOKEN_CONNECTION) = oSession.CreateSession(aConnectionInfo(S_UID_CONNECTION), aConnectionInfo(S_PWD_CONNECTION), , aConnectionInfo(S_PROJECT_CONNECTION), GetLng())
			If Err.Number <> 0 Then
				Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(sErrDescription), "PromptCuLib.asp", "GetHydraPrompt", "", "Error creating castor session", LogLevelError)
			End If
			lErrNumber = Err.number
		End If
	End If

	If lErrNumber = NO_ERR Then
		Set oReport = Server.CreateObject(PROGID_HELPER_RESULTSET)
		oReport.SessionID = aConnectionInfo(S_TOKEN_CONNECTION)
		oReport.ExecutionFlags = DssXmlExecutionResolve


		oReport.Submit(sReportID)
		Call oReport.GetResults
		set oAllPrompts = oReport.PromptsObject
		lErrNumber = Err.number
	End If

	If lErrNumber = NO_ERR Then
		Redim aPromptInfo(oAllPrompts.OpenCount-1, MAX_PROMPT_INFO)
		j = 0
		For i = 1 to oAllPrompts.Count
			Set oSinglePrompt = oAllPrompts.Item(i)
			If oSinglePrompt.Used And not oSinglePrompt.Closed Then
				aPromptInfo(j, PROMPT_INDEX) = oSinglePrompt.Index
				aPromptInfo(j, PROMPT_TITLE) = oSinglePrompt.Title
				aPromptInfo(j, PROMPT_DESC) = oSinglePrompt.Meaning
				aPromptInfo(j, PROMPT_TYPE) = oSinglePrompt.PromptType
				aPromptInfo(j, PROMPT_MIN) = oSinglePrompt.Min
				aPromptInfo(j, PROMPT_MAX) = oSinglePrompt.Max
				aPromptInfo(j, PROMPT_ISID) =sInfoSourceID
				j = j + 1
			End If
		Next
		Call CloseCastorSession(aConnectionInfo)
		lErrNumber = Err.number
	End If

	set oDetailDOM = nothing
	set oCurrQO = nothing
	set oISOI = nothing
	set oSession = nothing
	set oReport = nothing
	set oAllPrompts = nothing
	set oSinglePrompt = nothing
	set oInfoSource = nothing

	GetPromptsInfo = lErrNumber
	Err.Clear
End Function



Function GetPromptsForQuestion(aSvcConfigInfo, aPromptList)
'********************************************************
'*Purpose: Return the prompts for the custom question of the subsset.
'           The prompt info is taken from castor
'*Inputs:  aSvcConfig: The array with information of the service.
'*Outputs: aMapList:  Returns the information of the maps found
'********************************************************
Const PROCEDURE_NAME = "GetPromptsForQuestion"
Dim lErr
Dim lCount

Dim sSiteId
Dim sQODetailsXML
Dim oDecoder
Dim oDetailDOM
Dim oProperty
Dim sEncodedData
Dim asQuestionObjectID()

    On Error Resume Next
    lErr = NO_ERR

    Redim asQuestionObjectID(0)
    asQuestionObjectID(0) = aSvcConfigInfo(SVCCFG_AQ_ID)
	sSiteId = Application.Value("SITE_ID")

    Set oDecoder = Server.CreateObject(PROGID_BASE64)
	lErr = Err.number
	If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetDetailsForQuestions", LogLevelTrace)

    If lErr = NO_ERR Then
	    lErr = co_GetDetailsForQuestions(sSiteId, asQuestionObjectID, sQODetailsXML)
	    If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetDetailsForQuestions", LogLevelTrace)
    End If


	'decode
	If lErr = NO_ERR Then
		lErr = LoadXMLDOMFromString(aConnectionInfo, sQODetailsXML, oDetailDOM)
		If lErr <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, CStr(lErr), CStr(Err.Description), Err.source, "ServicesConfigCuLib.asp", "ReadPromptQuestionFromCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString: sQODetailsXML", LogLevelTrace)
		Else
			Set oProperty = oDetailDOM.selectSingleNode("/mi/qos/oi/prs/pr[@n = 'definition']")
			sEncodedData = oProperty.text
			oProperty.text = oDecoder.Decode(sEncodedData)

			Set oProperty = oDetailDOM.selectSingleNode("/mi/in/oi/prs/pr[@id = 'ISM_connInfo']")
			sEncodedData = oProperty.text
			oProperty.text = oDecoder.Decode(sEncodedData)

			sQODetailsXML = oDetailDOM.xml
		End If
	End If

    If lErr = NO_ERR Then
	    lErr = GetPromptsInfo(aConnectionInfo, aSvcConfigInfo(SVCCFG_AQ_ID), sQODetailsXML, aPromptList)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, Err.description, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling GetPromptsInfo", LogLevelTrace)
    End If

    Set oDecoder = Nothing
    Set oDetailDOM = Nothing
    Set oProperty = Nothing

    GetPromptsForQuestion = lErr
    Err.Clear

End Function

Function GetQuestionStorageType(aSvcConfigInfo, aPromptList, nStorageType)
'********************************************************
'*Purpose: Return the type of storage for the question object selected.
'           The storage is based on que number and type of prompts of this question
'*Inputs:  aSvcConfigInfo: Informationa bout the configuration of the service;
'          aPromptList: The list of prompts of the question object.
'*Outputs: nStorageType = The type of storage
'********************************************************
Const PROCEDURE_NAME = "GetQuestionStorageType"
Dim lErr

Dim i
Dim lCount
Dim bValidPrompt

    On Error Resume Next
    lErr = NO_ERR

    If IsEmpty(aPromptList) Then
        nStorageType = STORAGE_NONE

    Else
        lCount = UBound(aPromptList) + 1

        If lCount = 1 Then
            If ValidPromptType(aPromptList(0, PROMPT_TYPE)) Then
                If aPromptList(0, PROMPT_TYPE) = DssXmlPromptElements And _
                   CLng(aPromptList(0, PROMPT_MIN)) <= 1  Then
                    nStorageType = STORAGE_ALL
                ElseIf aPromptList(0, PROMPT_TYPE) = DssXmlPromptLong Or _
                       aPromptList(0, PROMPT_TYPE) = DssXmlPromptString Or _
                       aPromptList(0, PROMPT_TYPE) = DssXmlPromptDouble Or _
                       aPromptList(0, PROMPT_TYPE) = DssXmlPromptDate Then
                    nStorageType = STORAGE_ALL
                Else
                    nStorageType = STORAGE_MAP_ONLY
                End If
            Else
                nStorageType = STORAGE_NONE
            End If

        Else
            bValidPrompt = True

            For i = 0 To lCount - 1
                If Not ValidPromptType(aPromptList(i, PROMPT_TYPE)) Then
                    bValidPrompt = False
                    Exit For
                End If
            Next

            If bValidPrompt Then
                nStorageType = STORAGE_MAP_ONLY
            Else
                nStorageType = STORAGE_NONE
            End If

        End If
    End If

    GetQuestionStorageType = lErr
    Err.Clear

End Function


Function GetQuestionMap(aSvcConfigInfo, sMapId)
'********************************************************
'*Purpose: Returns the map id that this question will use for storing its answers.
'*Inputs:  sQuestionId: The question we want the map from
'*Outputs: sMapId: The map this question uses.
'********************************************************
Const PROCEDURE_NAME = "GetQuestionMap"
Dim lErr
Dim sCacheXML
Dim sCacheName
Dim oDOM
Dim oQuestion
Dim bIsExtra

    On Error Resume Next
    lErr = NO_ERR

    sCacheName = GetSvcConfigCacheName(aSvcConfigInfo)

    If lErr = NO_ERR Then
        lErr = ReadCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, sCacheXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ReadCache", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sCacheXML", LogLevelTrace)
        End If
    End If

    If lErr = NO_ERR Then
        'See if the question is already in the list:
        Set oQuestion = oDOM.selectSingleNode("//oi[@id='" & aSvcConfigInfo(SVCCFG_AQ_ID) &"']")
        If Not oQuestion Is Nothing Then
            'Check if it is a question that should have a map:
            If oQuestion.getAttribute("tp") <> "custom" Then
                'Questions that are not custom, should use standard question configuration:
                lErr = ERR_QUESTION_IN_SERVICE_DEF

            Else
                'Extra questions can either be only alternate or only addtional:
                bIsExtra = (GetPropertyValue(oQuestion, PROP_VALUE) = "false")

                If (bIsExtra And (aSvcConfigInfo(SVCCFG_QO_ID) <> NEW_OBJECT_ID)) Or _
                   (Not bIsExtra And (aSvcConfigInfo(SVCCFG_QO_ID) = NEW_OBJECT_ID)) Then

                    lErr = ERR_QUESTION_ALREADY_USED

                Else
                    sMapId = GetPropertyValue(oQuestion, PROP_STORAGE_MAPPING_ID)

                End If
            End If
        End If

    End If

    Set oDOM = Nothing
    Set oQuestion = Nothing

    GetQuestionMap = lErr
    Err.Clear

End Function

Function ValidPromptType(nType)
'********************************************************
'*Purpose: Returns if the given prompt type is valid
'*Inputs:  nType: The type of prompt
'*Outputs: True, if valid, False otherwise
'********************************************************
    Select Case nType
    Case DssXmlPromptLong, DssXmlPromptString, DssXmlPromptDouble, DssXmlPromptDate, DssXmlPromptElements
        ValidPromptType = True
    Case Else
        ValidPromptType = False
    End Select

End Function

Function GetMapInfo(sMapId, aMapInfo)
'********************************************************
'*Purpose: Return information of the given Map
'*Inputs:  sMapId: The Id of the map
'*Outputs: oMapDefinitionDOM: The DOM object of the map definition.
'********************************************************
Const PROCEDURE_NAME = "GetMapInfo"
Dim lErr

Dim sMapDefinitionXML
Dim sSiteId

Dim oMapDefinitionDOM
Dim oMap
Dim oTables

Dim i
Dim lCount

    On Error Resume Next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    If lErr = NO_ERR Then
        If sMapId = NEW_OBJECT_ID Then
            lErr = LoadXMLDOMFromString(aConnectionInfo, "<mi><in></in></mi>", oMapDefinitionDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sMapDefinitionXML", LogLevelTrace)
        Else
            lErr = co_getMappingDefinition(sSiteId, sMapId, sMapDefinitionXML)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getMapDefinition", LogLevelTrace)
            Else
                lErr = LoadXMLDOMFromString(aConnectionInfo, sMapDefinitionXML, oMapDefinitionDOM)
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sMapDefinitionXML", LogLevelTrace)
            End If
        End If
    End If

    If lErr = NO_ERR Then
        Set oMap = oMapDefinitionDOM.selectSingleNode("//mapping")

        ReDim aMapInfo(MAX_MAP_INFO)
        aMapInfo(MAP_ID)   = oMap.getAttribute("id")
        aMapInfo(MAP_NAME) = oMap.getAttribute("n")
        aMapInfo(MAP_FILTER) = oMap.getAttribute("f")

        Set oTables = oMapDefinitionDOM.selectNodes("//table")
        lCount = oTables.length

        If lCount > 0 Then

            For i = 0 To lCount - 1
                aMapInfo(MAP_DESC) = aMapInfo(MAP_DESC) & oTables(i).getAttribute("id") & ", "
            Next

            aMapInfo(MAP_DESC) = Left(aMapInfo(MAP_DESC), Len(aMapInfo(MAP_DESC)) - 2)
            aMapInfo(MAP_DBALIAS) = oTables(0).getAttribute("connection")

        End If
    End If

    GetMapInfo = lErr
    Err.Clear

End Function


Function GetMapDOM(sMapId, oMapDOM)
'********************************************************
'*Purpose: Return information of the given Map
'*Inputs:  sMapId: The Id of the map
'*Outputs: oMapDefinitionDOM: The DOM object of the map definition.
'********************************************************
Const PROCEDURE_NAME = "GetMapDOM"
Dim lErr

Dim sMapDefinitionXML
Dim sSiteId

    On Error Resume Next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    If lErr = NO_ERR Then
        If sMapId = NEW_OBJECT_ID Then
            lErr = LoadXMLDOMFromString(aConnectionInfo, "<mapping f='' />", oMapDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sMapDefinitionXML", LogLevelTrace)
        Else
            lErr = co_getMappingDefinition(sSiteId, sMapId, sMapDefinitionXML)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getMapDefinition", LogLevelTrace)
            Else
                lErr = LoadXMLDOMFromString(aConnectionInfo, sMapDefinitionXML, oMapDOM)
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sMapDefinitionXML", LogLevelTrace)
            End If
        End If
    End If


    GetMapDOM = lErr
    Err.Clear

End Function

Function GetTables(sDBAlias, sFilter, aTables)
'********************************************************
'*Purpose: Returns the Tables of a given Database
'*Inputs:  sDBAlias: The dbalias where the tables are.
'          sFilter: A valid SQL expression to restrict the search
'*Outputs: aTables: The list of tables.
'********************************************************
Const PROCEDURE_NAME = "GetTables"
Dim lErr

Dim sTablesXML
Dim oDOM
Dim oTables
Dim oPr

Dim lCount
Dim i

Dim lDotPosition
Dim sOwner
Dim sPrefix

    lDotPosition = InStr(1,sFilter, ".")

    If lDotPosition > 0 Then
    	sOwner = Left(sFilter, lDotPosition-1)
    	sPrefix = Mid(sFilter, lDotPosition +1)
    Else
    	sOwner = ""
    	sPrefix = sFilter
    End If

    On Error Resume Next
    lErr = NO_ERR

    If lErr = NO_ERR Then
        lErr = co_getTables(sDBAlias, sOwner, sPrefix, sTablesXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getTables", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sTablesXML, oDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sTablesXML", LogLevelTrace)
        End If
    End If

    If lErr = NO_ERR Then
        Set oTables = oDOM.selectNodes("//oi")
        lCount = oTables.length

        If lCount > 0 Then
            Redim aTables(lCount - 1)

            For i = 0 To lCount - 1
            	Set oPr = oTables(i).selectNodes("prs/pr[@id='TABLE_SCHEM']")
                aTables(i) = oPr(0).getAttribute("v") & "." & oTables(i).getAttribute("n")
            Next
        End If
    End If

    GetTables = lErr
    Err.Clear

End Function

Function GetTablesInfo(sDBAlias, aTables, aTablesInfo)
'********************************************************
'*Purpose: Return information of the given Tables
'*Inputs:  sDBAlias, the database connection
'          aTables: the tables to return back the information
'*Outputs: aTablesInfo: An array with the Tables information:
'*          0: Table Name
'*          1: A ; delimited columns name string
'********************************************************
Const PROCEDURE_NAME = "GetTablesInfo"
Dim lErr

Dim sColumnsXML
Dim oDOM
Dim oColumns

Dim lDotPosition
Dim sOwner
Dim sTableName

Dim lCount, lColumnCount
Dim i, j


    On Error Resume Next
    lErr = NO_ERR

    If lErr = NO_ERR Then
        lCount = UBound(aTables)
        Redim aTablesInfo(lCount, MAX_TABLE_INFO)

        For i = 0 To lCount
            aTablesInfo(i, TABLE_ID) = aTables(i)
            aTablesInfo(i, TABLE_DBALIAS) = sDBAlias

			lDotPosition = InStr(1,aTables(i), ".")

			If lDotPosition > 0 Then
				sOwner = Left(aTables(i), lDotPosition-1)
				sTableName = Mid(aTables(i), lDotPosition +1)
			Else
				sOwner = ""
				sTableName = aTables(i)
			End If

            lErr = co_getColumns(sDBAlias, sOwner, sTableName, sColumnsXML)
            If lErr <> NO_ERR Then
                Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_getColumns", LogLevelTrace)
            Else
                lErr = LoadXMLDOMFromString(aConnectionInfo, sColumnsXML, oDOM)
                If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sColumnsXML", LogLevelTrace)
            End If
            If lErr <> NO_ERR Then Exit For

            Set oColumns = oDOM.selectNodes("//oi")

            If oColumns.length > 0 Then
                lColumnCount = oColumns.length
                For j = 0 to lColumnCount
                    aTablesInfo(i, TABLE_COLUMNS) = aTablesInfo(i, TABLE_COLUMNS) & oColumns(j).getAttribute("n") & ";"
                    aTablesInfo(i, TABLE_COLUMN_GUIDS) = aTablesInfo(i, TABLE_COLUMN_GUIDS) & "t" & i & "c" & j & ";"
                Next
            End If
        Next
    End If

    Set oColumns = Nothing
    Set oDOM = Nothing

    GetTablesInfo = lErr
    Err.Clear

End Function

Function SaveStorageMapping(aSvcConfigInfo, aMapInfo, aTablesInfo)
'********************************************************
'*Purpose: Creates and saves a mapping definition
'*Inputs:  aSvcConfigInfo: the information of the map configuration
'          aTablesInfo: An array with the Tables information:
'*Outputs: None
'********************************************************
Const PROCEDURE_NAME = "SaveStorageMapping"
Dim lErr

Dim sDefinitionXML
Dim sMapXML

Dim bAddTable
Dim aColumns
Dim aValues
Dim sField

Dim lCount, lColumnCount
Dim i, j

Dim sSiteId

    On Error Resume Next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    If lErr = NO_ERR Then

        If aMapInfo(MAP_ID) = NEW_OBJECT_ID Then
            aMapInfo(MAP_ID) = GetGUID()
            lErr = AddStorageMapToCache(aSvcConfigInfo, aMapInfo)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling AddStorageMapToCache", LogLevelTrace)
        Else
            lErr = co_deleteMappingDefinition(sSiteId, aMapInfo(MAP_ID))
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_deleteMapDefinition", LogLevelTrace)
        End If

    End If

    If lErr = NO_ERR Then

        lCount = UBound(aTablesInfo)
        sDefinitionXML = "<mapping n=""" & Server.HTMLEncode(aMapInfo(MAP_NAME)) & """ f=""" & Server.HTMLEncode(aMapInfo(MAP_FILTER)) & """>"

        For i = 0 To lCount
            bAddTable = True

            aColumns = Split(aTablesInfo(i, TABLE_COLUMNS), ";")
            aValues = Split(aTablesInfo(i, TABLE_COLUMN_VALUES), ";")

            lColumnCount = Ubound(aColumns) - 1

            For j = 0 To lColumnCount
                If (Len(Trim(aColumns(j))) > 0) And (Len(Trim(aValues(j))) > 0) Then
                    If bAddTable Then sDefinitionXML = sDefinitionXML & "<table id=""" & Server.HTMLEncode(aTablesInfo(i, TABLE_ID)) & """ connection=""" & Server.HTMLEncode(aTablesInfo(i, TABLE_DBALIAS)) & """ >"
                    sField = Trim(aValues(j))
                    If (Left(sField, 13) = "subscription.") Or Left(sField, 3) = "qo." Then
                        sDefinitionXML = sDefinitionXML & "<col field=""" & Server.HTMLEncode(Trim(aValues(j))) & """ id=""" & Server.HTMLEncode(Trim(aColumns(j))) & """ />"
                    Else
                        sDefinitionXML = sDefinitionXML & "<col field=""" & ANSWER_CONSTANT & """ id=""" & Server.HTMLEncode(Trim(aColumns(j))) & """ value=""" & Server.HTMLEncode(Trim(aValues(j))) & """ />"
                    End If
                    bAddTable = False
                End If
            Next

            If bAddTable = False Then sDefinitionXML = sDefinitionXML & "</table>"

        Next

        sDefinitionXML = sDefinitionXML & "</mapping>"

        lErr = co_createMappingDefinition(sSiteId, aMapInfo(MAP_ID), aMapInfo(MAP_NAME), sDefinitionXML)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_createMappingDefinition", LogLevelTrace)

    End If


    If lErr = NO_ERR Then
        If Len(aSvcConfigInfo(SVCCFG_QO_ID)) = 0 Then
            aSvcConfigInfo(SVCCFG_SS_MAP_ID) = aMapInfo(MAP_ID)
        End If
    End If

    Erase aColumns
    Erase aValues

    SaveStorageMapping = lErr
    Err.Clear

End Function

Function AddStorageMapToCache(aSvcConfigInfo, aMapInfo)
'********************************************************
'*Purpose: Adds a map ot the cache XML
'*Inputs:  aSvcConfigInfo: the information of the map configuration
'          aMapInfo: An array with the Tables information:
'*Outputs: None
'********************************************************
Const PROCEDURE_NAME = "AddStorageMapToCache"
Dim lErr

Dim sCacheName
Dim sCacheXML

Dim oDOM
Dim oMap
Dim oAttr

Dim oMaps

    On Error Resume Next
    lErr = NO_ERR

    sCacheName = GetSvcConfigCacheName(aSvcConfigInfo)

    If lErr = NO_ERR Then
        lErr = ReadCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, sCacheXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ReadCache", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sCacheXML", LogLevelTrace)
        End If
    End If

    If lErr = NO_ERR Then
        Set oMaps = oDOM.selectSingleNode("mi/in/maps")
        Set oMap = oMaps.appendChild(oDOM.createElement("oi"))
        oMap.setAttribute "id", aMapInfo(MAP_ID)
        oMap.setAttribute "tp", TYPE_STORAGE_MAPPING
    End If

    'Save the XML into the cache:
    If lErr = NO_ERR Then
        lErr = WriteCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, oDOM.xml)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling WriteCache", LogLevelTrace)
    End If


    AddStorageMapToCache = lErr
    Err.Clear

End Function

Function RemoveQuestionConfig(aSvcConfigInfo)
'********************************************************
'*Purpose: Deletes a question from the subsset configuration
'*Inputs:  aSvcConfigInfo: the information of the map configuration
'*Outputs: None
'********************************************************
Const PROCEDURE_NAME = "RemoveQuestionConfig"
Dim lErr
Dim sSiteId

Dim sCacheXML
Dim sCacheName

Dim oDOM
Dim oQuestion
Dim oExtras

    On Error Resume Next
    lErr = NO_ERR

    sCacheName = GetSvcConfigCacheName(aSvcConfigInfo)

    If lErr = NO_ERR Then
        lErr = ReadCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, sCacheXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling ReadCache", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oDOM)
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling LoadXMLDOMFromString: sCacheXML", LogLevelTrace)
        End If
    End If

    If lErr = NO_ERR Then
        Set oExtras = oDOM.selectSingleNode("//extras")
        Set oQuestion = oExtras.selectSingleNode("oi[prs/pr[@v='" & aSvcConfigInfo(SVCCFG_AQ_ID) &"']]")

        If Not oQuestion Is Nothing Then Call oExtras.removeChild(oQuestion)

    End If

    'Save the XML into the cache:
    If lErr = NO_ERR Then
        lErr = WriteCache(sCacheName, SVC_CONFIG_CACHE_FOLDER, oDOM.xml)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling WriteCache", LogLevelTrace)
    End If


    RemoveSubsSetConfig = lErr
    Err.Clear

End Function

Function DeleteServiceConfig(aSvcConfigInfo)
'********************************************************
'*Purpose: Deletes a service config  from MD
'*Inputs:  aSvcConfigInfo: the information of the map configuration
'*Outputs: None
'********************************************************
Const PROCEDURE_NAME = "DeleteServiceConfig"
Dim lErr
Dim sSiteId


    On Error Resume Next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    'Get configured subscription sets:
    If lErr = NO_ERR Then
        If Len(aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID)) = 0 Then
            lErr = getConfigObjectID(aSvcConfigInfo(SVCCFG_SVC_ID), TYPE_SERVICE, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID))
            If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling getConfigObjectID", LogLevelTrace)
        End If
    End If

    'Delete it:
    If lErr = NO_ERR Then
        lErr = co_deleteObject(sSiteId, aSvcConfigInfo(SVCCFG_SVC_CONFIG_ID))
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_deleteObject", LogLevelTrace)
    End If

    DeleteServiceConfig = lErr
    Err.Clear

End Function

Function DeleteSubsSetConfig(aSvcConfigInfo)
'********************************************************
'*Purpose: Deletes a subscription set from MD
'*Inputs:  aSvcConfigInfo: the information of the map configuration
'*Outputs: None
'********************************************************
Const PROCEDURE_NAME = "DeleteSubsSetConfig"
Dim lErr
Dim sSiteId


    On Error Resume Next
    lErr = NO_ERR

    sSiteId = Application.Value("SITE_ID")

    If lErr = NO_ERR Then
        lErr = co_deleteObject(sSiteId, aSvcConfigInfo(SVCCFG_SS_CONFIG_ID))
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, sErr, Err.Source, "ServicesConfigCuLib.asp", PROCEDURE_NAME, "", "Error calling co_deleteObject", LogLevelTrace)
    End If

    DeleteSubsSetConfig = lErr
    Err.Clear

End Function

Function GetStorageMapName(sMapId)
'********************************************************
'*Purpose: Returns a Name of a storage mapping
'*Inputs:  sMapId: The Storage Mapping we need:
'*Outputs: The name of the mapping
'********************************************************
Const PROCEDURE_NAME = "GetStorageMapName"
Dim lErr
Dim aMapInfo

    On Error Resume Next

    If sMapId = "sbr" Then
        GetStorageMapName = asDescriptors(581) '"subscription book repository"
    Else
        lErr = GetMapInfo(sMapId, aMapInfo)
        If lErr = NO_ERR Then
            GetStorageMapName = aMapInfo(MAP_NAME)
        Else
            GetStorageMapName = sMapId
        End If
    End If

    Err.Clear


End Function

Function GetExtraQuestionIndex(aExtraQuestions, sQuestionId)
'********************************************************
'*Purpose: Returns the index within the extra questions of the question with the given Id
'*Inputs:  aExtraQuestions: The array of extra questions.
'*          sQuestionId: the id of the extra question to find
'*Outputs: The name of the mapping
'********************************************************
Const PROCEDURE_NAME = "GetExtraQuestionIndex"
Dim lIndex
Dim lCount
Dim i

    On Error Resume next

    lCount = UBound(aExtraQuestions)

    lIndex = -1
    For i = 0 To lCount
        If aExtraQuestions(i, QO_ALTERNATE_ID) = sQuestionId Then
            lIndex = i
            Exit For
        End If
    Next

    GetExtraQuestionIndex = lIndex
    Err.Clear

End Function


%>