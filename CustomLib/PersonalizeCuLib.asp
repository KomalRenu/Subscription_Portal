<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!--#include file="../CoreLib/PersonalizeCoLib.asp" -->
<!--#include file="../CoreLib/AddressCoLib.asp" -->
<%
	Function ParseRequestForPersonalize(oRequest, sSubGUID, sFolderID, sQOID)
	'********************************************************
	'*Purpose:
	'*Inputs: oRequest
	'*Outputs: sReqUserSubs
	'********************************************************
		On Error Resume Next
		Dim lErrNumber

		lErrNumber = NO_ERR

		sSubGUID = ""
		sFolderID = ""
		sQOID = ""

	    sSubGUID = Trim(CStr(oRequest("eSGUID")))
	    sFolderID = Trim(CStr(oRequest("folderID")))
	    sQOID = Trim(CStr(oRequest("QOID")))

		If Err.number <> NO_ERR Then
		    lErrNumber = Err.number
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "ParseRequestForPersonalize", "", "Error setting variables equal to Request variables", LogLevelError)
		Else
	        If Len(sSubGUID) = 0 Then
	        	lErrNumber = URL_MISSING_PARAMETER
	        End If
		End If

		ParseRequestForPersonalize = lErrNumber
		Err.Clear
	End Function

	Function RenderPersonalize(sCacheXML, sSubGUID, sFolderID)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'*TO DO: add error messages
	'********************************************************
	    On Error Resume Next
	    Dim lErrNumber

	    lErrNumber = NO_ERR

		Call RenderQuestions_Personalize(sCacheXML, sSubGUID, sFolderID)

	    RenderPersonalize = lErrNumber
	    Err.Clear
	End Function

	'Function RenderQuestions_Personalize(sGetQuestionsForPublicationXML, sSubGUID, sFolderID)
	Function RenderQuestions_Personalize(sCacheXML, sSubGUID, sFolderID)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'*TO DO: add error handling, messages
	'********************************************************
	    On Error Resume Next
	    Dim lErrNumber
	    Dim oQuestionsDOM		'not used any more
	    Dim oCacheDOM
	    Dim oQuestions
	    Dim oCurrentQuestion
	    Dim oProfiles
	    Dim oCurrentProfile
	    Dim iNumQuestions
	    Dim i
	    Dim oCurrQuestion
	    Dim oAnswer
	    Dim bFirst

	    lErrNumber = NO_ERR

	    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "RenderQuestions_Personalize", "", "Error loading sCacheXML", LogLevelError)
		End If

		Set oQuestions = oCacheDOM.selectNodes("/mi/qos/mi/in/oi")

	    If lErrNumber = NO_ERR Then
	        If oQuestions.length > 0 Then
	            iNumQuestions = CInt(oQuestions.length)
	            Response.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>"
	            Response.Write "<TR><TD><IMG SRC=""images/1ptrans.gif"" HEIGHT=""1"" WIDTH=""40"" ALT="""" BORDER=""0"" /></TD><TD>"
	            Response.Write "<TABLE BORDER=""0"" CELLPADDING=""2"" CELLSPACING=""0"">"

				bFirst = True
	            For i = 1 to (iNumQuestions)
					set oCurrQuestion = oQuestions.item(i-1)
	                If Strcomp(oCurrQuestion.getAttribute("hidden"), "0", vbTextCompare) = 0 Then
						If bFirst Then
						    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""QOID"" VALUE=""" & oCurrQuestion.getAttribute("id") & """ />"
						    bFirst = False
						End If
       					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP"">"
						Response.Write "<IMG WIDTH=""3"" HEIGHT=""8"" ALT="""" BORDER=""0"" SRC=""images/bullet.gif"" />"
						Response.Write "</TD>"
						Response.Write "<TD VALIGN=TOP>"
						Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ color=""#cc0000"" size=""" & aFontInfo(N_MEDIUM_FONT) & """><b>" & oCurrQuestion.getAttribute("n") & "</b></font><BR />"
						Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>" & oCurrQuestion.getAttribute("des") & "</font>"
						Response.Write "</TD></TR>"
					End If
	            Next

	            Response.Write "<TR><TD COLSPAN=2><IMG WIDTH=""1"" HEIGHT=""40"" ALT="""" BORDER=""0"" SRC=""images/1ptrans.gif"" /></TD></TR>"
				Response.Write "</TABLE>"
				Response.Write "</TD></TR></TABLE>"
	        Else
	            'Will this case ever occur?
	        End If
	    End If

	    Set oCacheDOM = nothing
	    Set oQuestionsDOM = Nothing
	    Set oQuestions = Nothing
	    Set oCurrentQuestion = Nothing
	    Set oProfiles = Nothing
	    Set oCurrentProfile = Nothing
		Set oCurrQuestion = nothing
	    Set oAnswer = nothing

        RenderQuestions_Personalize = lErrNumber
        Err.Clear
	End Function

    Function RenderPath_Personalize(sSubGUID, sServiceID, sServiceName, sFolderID, sGetFolderContentsXML, sStatusFlag)
	'********************************************************
	'*Purpose:
	'*Inputs:
	'*Outputs:
	'*TO DO: add error handling, messages
	'********************************************************
        On Error Resume Next
        Dim lErrNumber
        Dim oContentsDOM
        Dim oFolder
        Dim iNumFolders
        Dim i
        Dim sLastFolder

        iNumFolders = 0
        lErrNumber = NO_ERR
        sLastFolder = ""

        If sFolderID <> "" Then
            Set oContentsDOM = Server.CreateObject("Microsoft.XMLDOM")
		    oContentsDOM.async = False
		    If oContentsDOM.loadXML(sGetFolderContentsXML) = False Then
		    	lErrNumber = ERR_XML_LOAD_FAILED
		    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "RenderPath_Personalize", "", "Error loading sGetFolderContentsXML", LogLevelError)
		    	'add error message
            Else
                iNumFolders = CInt(oContentsDOM.selectNodes("//a").length)
                If Err.number <> NO_ERR Then
                    lErrNumber = Err.number
                    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "RenderPath_Personalize", "", "Error retrieving oi nodes", LogLevelError)
                    'add error message
                End If
		    End If
		End If

        Response.Write "<font face=""" & aFontInfo(S_FAMILY_FONT) & """ size=""" & aFontInfo(N_SMALL_FONT) & """>"
        Response.Write asDescriptors(26) & " " 'Descriptor: You are here:
        If lErrNumber = NO_ERR Then
            If iNumFolders > 0 Then
                Response.Write "<A HREF=""services.asp""><font color=""#000000"">" & asDescriptors(362) & "</font></A>" 'Descriptor: Services
                Set oFolder = oContentsDOM.selectSingleNode("/mi/as")
                For i=1 To iNumFolders
                    Set oFolder = oFolder.selectSingleNode("a")
                    Response.Write " > "
                    Response.Write "<A HREF=""services.asp?folderID=" & oFolder.selectSingleNode("fd").getAttribute("id") & """><font color=""#0000"">" & oFolder.selectSingleNode("fd").getAttribute("n") & "</font></A>"
                    If i=iNumFolders Then sLastFolder = oFolder.selectSingleNode("fd").getAttribute("id")
                Next
                Response.Write " > <A HREF=""subscribe.asp?eSGUID=" & sSubGUID & "&serviceID=" & sServiceID & "&folderID=" & sLastFolder & "&serviceName=" & Server.URLEncode(sServiceName) & "&sf=" & sStatusFlag & """><font color=""#000000"">" & asDescriptors(457) & " " & oContentsDOM.selectSingleNode("/mi/fct/oi[@id = '" & sServiceID & "']").getAttribute("n") & "</font></A>" 'Descriptor: Subscribe to:
                Response.Write " > <b>" & asDescriptors(458) & "</b>" 'Descriptor: Personalize your subscription
            Else
                Response.Write "<A HREF=""subscribe.asp?eSGUID=" & sSubGUID & "&serviceID=" & sServiceID & "&folderID=" & sFolderID & "&serviceName=" & Server.URLEncode(sServiceName) & "&sf=" & sStatusFlag & """><font color=""#000000"">" & asDescriptors(457) & " " & sServiceName & "</font></A>" 'Descriptor: Subscribe to:
                Response.Write " > <b>" & asDescriptors(458) & "</b>" 'Descriptor: Personalize your subscription
            End If
        Else
            'add handling
        End If
        Response.Write "</font>"

        Set oContentsDOM = Nothing
        Set oFolder = Nothing

        RenderPath_Personalize = lErrNumber
        Err.Clear
	End Function

Function GetVariablesFromCache_Personalize(sCacheXML, sFolderID, sPublicationIDFromCache, sServiceID, sServiceName, sAddressName, sScheduleName, sSubsEnabled, sStatusFlag, sSubSetID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oSub

    lErrNumber = NO_ERR

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
	oCacheDOM.async = False
    oCacheDOM.loadXML(sCacheXML)

    Set oSub = oCacheDOM.selectSingleNode("/mi/sub")

    sFolderID = oSub.getAttribute("fid")
    sPublicationIDFromCache = oSub.getAttribute("pubid")
    sServiceID = oSub.getAttribute("svcid")
    sServiceName = oSub.getAttribute("svn")
    sAddressName = oSub.getAttribute("adn")
    sScheduleName = oSub.getAttribute("scn")
    sSubsEnabled = oSub.getAttribute("enf")
    sStatusFlag = oSub.getAttribute("sf")
    sSubSetID = oSub.getAttribute("sbstid")

    Set oSub = Nothing
    Set oCacheDOM = Nothing

    GetVariablesFromCache_Personalize = lErrNumber
    Err.Clear
End Function

Function CheckNumberOfQuestions(sQuestionsAndProfilesXML, iNumQuestions)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oQuestionsAndProfilesDOM

    lErrNumber = NO_ERR
    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sQuestionsAndProfilesXML, oQuestionsAndProfilesDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		iNumQuestions = oQuestionsAndProfilesDOM.selectNodes("/mi/in/oi[@tp='" & TYPE_QUESTION & "']").length
	End If

	Set oQuestionsAndProfilesDOM = Nothing

	CheckNumberOfQuestions = lErrNumber
	Err.Clear
End Function

Function CheckNumberOfVisibleQuestions(sCacheXML, iVisibleQOs, sFirstVisibleQOID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'*TO DO: Add error handling!
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oCurrQO
    Dim sISID
    Dim oISMProgID
    Dim oDefaultProfile
    Dim oAnswer
    Dim bHiddenQO

    lErrNumber = NO_ERR
    iVisibleQOs = 0

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		For each oCurrQO in oCacheDOM.selectNodes("/mi/qos/mi/in/oi[@tp='" & TYPE_QUESTION & "']")
			bHiddenQO = False
			If CLng(oCurrQO.getAttribute("qtp")) = QO_TYPE_SLICING Then
				bHiddenQO = True
			End If

			sISID = oCurrQO.getAttribute("isid")
			set oISMProgID = oCacheDOM.selectSingleNode("/mi/in/oi[@tp='" & TYPE_INFORMATION_SOURCE & "' $and$ @id='" & oCurrQO.getAttribute("isid") & "']/prs/pr[@n='ISM_admin_progid']")
			If oISMProgID is nothing Then
				lErrNumber = ERR_CACHE_CONTENT
			ElseIf strcomp(oISMProgID.getAttribute("v"), "UserDetailsISM.cUserDetails", vbBinaryCompare) = 0 Then
				bHiddenQO = True
				'use default to answer
				Call oCurrQO.setAttribute("isid", sISID)
				Set oDefaultProfile = oCurrQO.selectSingleNode("mi/oi[@tp = '" & TYPE_PROFILE & "' and @def='1']")
				If oDefaultProfile is nothing Then
					lErrNumber = ERR_USERDEFAULT_NOTEXIST
				Else
					Set oAnswer = oCacheDOM.createElement("answer")
					Call oCurrQO.appendChild(oAnswer)
					Call oAnswer.setAttribute("n", oDefaultProfile.getAttribute("n"))
					Call oAnswer.setAttribute("prefID", oDefaultProfile.getAttribute("id"))
				End If
			End If

			If bHiddenQO Then
				Call oCurrQO.setAttribute("hidden", "1")
			Else
				Call oCurrQO.setAttribute("hidden", "0")
				iVisibleQOs = iVisibleQOs + 1
				If iVisibleQOs = 1 Then
					sFirstVisibleQOID = oCurrQO.getAttribute("id")
				End If
			End If
		Next
	End If

	sCacheXML = oCacheDOM.xml
    Set oCacheDOM = Nothing
	Set oCurrQO = Nothing
    Set oISMProgID = Nothing
    Set oDefaultProfile = Nothing
    Set oAnswer = Nothing

    CheckNumberOfVisibleQuestions = lErrNumber
    Err.Clear
End Function

Function AddQuestionsToCache_Personalize(sCacheXML, sGetQuestionsAndProfilesForPublicationXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oQuestionsDOM
    Dim oQOS
    Dim sQOID
	Dim oQOFrom
    Dim oProfilesMI
    Dim oQOTo
    Dim oOldProfilesMI

    lErrNumber = NO_ERR

    Set oCacheDOM = Server.CreateObject("Microsoft.XMLDOM")
	oCacheDOM.async = False
    If oCacheDOM.loadXML(sCacheXML) = False Then
        lErrNumber = ERR_XML_LOAD_FAILED
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "AddQuestionsToCache_Personalize", "", "Error loading sCacheXML", LogLevelError)
    Else
        Set oQuestionsDOM = Server.CreateObject("Microsoft.XMLDOM")
        oQuestionsDOM.async = False
        If oQuestionsDOM.loadXML(sGetQuestionsAndProfilesForPublicationXML) = False Then
            lErrNumber = ERR_XML_LOAD_FAILED
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "AddQuestionsToCache_Personalize", "", "Error loading sGetQuestionsAndProfilesForPublicationXML", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
		set oQOS = oCacheDOM.selectSingleNode("/mi/qos")
        If oQOS Is Nothing Then
            Set oQOS = oCacheDOM.createElement("qos")
            Call oCacheDOM.selectSingleNode("/mi").appendChild(oQOS)
            Call oQOS.appendChild(oQuestionsDOM.selectSingleNode("/mi"))
        Else
			For each oQOFrom in oQuestionsDOM.selectNodes("/mi/in/oi[@tp='5']")
				sQOID = oQOFrom.getAttribute("id")
				Set oProfilesMI = oQOFrom.selectSingleNode("mi")

				Set oQOTo = oCacheDOM.selectSingleNode("/mi/qos/mi/in/oi[@id='" & sQOID & "']" )
				If not oQOTo is nothing Then
					Set oOldProfilesMI = oQOTo.selectSingleNode("mi")
					If not oOldProfilesMI is nothing Then
						Call oQOTo.removeChild(oOldProfilesMI)
					End If
					Call oQOTo.appendChild(oProfilesMI)
				End if
			Next
		End If
    End If

    If lErrNumber = NO_ERR Then
        sCacheXML = oCacheDOM.xml
    End If

    Set oCacheDOM = Nothing
    Set oQuestionsDOM = Nothing
    set oQOFrom = Nothing
    set oProfilesMI = Nothing
    set oQOTo = Nothing
    set oOldProfilesMI = Nothing
    set oQOS = nothing

    AddQuestionsToCache_Personalize = lErrNumber
    Err.Clear
End Function

Function cu_GetSubscription(sSubscriptionGUID, sGetSubscriptionXML)
'********************************************************
'*Purpose:
'*Inputs: sSubscriptionGUID
'*Outputs: sGetSubscriptionXML
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetSubscription"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

	lErrNumber = co_GetSubscription(sSessionID, sSubscriptionGUID, sGetSubscriptionXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetSubscription", LogLevelTrace)
	End If

	cu_GetSubscription = lErrNumber
	Err.Clear
End Function

Function cu_GetQuestionsAndProfilesForSubscriptionSet(sSubSetID, sServiceID, sGQAPFSSXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetQuestionsAndProfilesForSubscriptionSet"
	Dim lErrNumber
	Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

	lErrNumber = co_GetQuestionsAndProfilesForSubscriptionSet(sSessionID, sSubSetID, sServiceID, sGQAPFSSXML)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetQuestionsAndProfilesForSubscriptionSet", LogLevelTrace)
	End If

	cu_GetQuestionsAndProfilesForSubscriptionSet = lErrNumber
	Err.Clear
End Function

Function ProcessProfileEdit(oRequest, sSubGUID)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim sQOID
	Dim sPrefID

	sQOID = CStr(oRequest("QOID_1"))
	sPrefID = CStr(oRequest("ProfileList_1"))

	Response.Redirect "preprompt.asp?src=personalization&subGUID=" & sSubGUID & "&qoid=" & sQOID & "&prefid=" & sPrefID

	ProcessProfileEdit = lErrNumber
	Err.Clear
End Function


Function CheckIfAllQOAnswered(oRequest, sSubGUID, sCacheXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim oCacheDOM
	Dim bAllAnswered
	Dim oQuestions
	Dim i
	Dim oCurrQuestion
	Dim oAnswer

	Call GetXMLDOM(aConnectionInfo, oCacheDOM, sErrDescription)
	Call oCacheDOM.loadXML(sCacheXML)
	If Err.number <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "RenderQuestions_Personalize", "", "Error retrieving question nodes", LogLevelError)
	End If

	bAllAnswered = True
	Set oQuestions = oCacheDOM.selectNodes("/mi/qos/mi/in/oi")
	For i = 1 to oQuestions.length
		set oCurrQuestion = oQuestions.item(i-1)
		set oAnswer = oCurrQuestion.selectSingleNode("answer")
		If oAnswer is nothing Then
			If oRequest("ProfileList_" & i).count > 0 Then
				sTemp = CStr(oRequest("ProfileList_" & i))
				'<answer n="XXX" prefID="" />
				Set oAnswer = oCacheDOM.createElement("answer")
            	Call oCurrQuestion.appendChild(oAnswer)
            	Call oAnswer.setAttribute("n", "XXX")
            	Call oAnswer.setAttribute("prefID", sTemp)
			Else
				Call oCurrQuestion.setAttribute("cl", "0")
				bAllAnswered = False
			End If
		End If
	Next

    lErr = WriteCache(sSubGUID, CStr(GetSessionID()), oCacheDOM.xml)
    If lErr = NO_ERR Then
		If bAllAnswered Then
            Set oCacheDOM = Nothing
            Set oQuestions = Nothing
            Set oCurrQuestion = Nothing
            Set oAnswer = Nothing

			Response.Redirect "modify_subscription.asp?subGUID=" & sSubGUID
		End If
    Else
		'LogErrorXML( )
    End If

    Set oCacheDOM = Nothing
    Set oQuestions = Nothing
    Set oCurrQuestion = Nothing
    Set oAnswer = Nothing

	CheckIfAllQOAnswered = lErrNumber
	Err.Clear
End Function

Function AddPreviousAnswerToCache(sSubGUID, sCacheXML, sGetSubscriptionXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
    On Error Resume Next
    Dim lErrNumber
    Dim oCacheDOM
    Dim oQuestions
    Dim oCurrentQuestion
    Dim oAnswer
    Dim oPref
    Dim oSubDOM
    Dim lTempErr
    Dim sPrefID
    Dim sQOID
    Dim sProfileXML
    Dim oProfileDOM
    Dim oProfile
    Dim sProfileName
    Dim sProfileDesc
    Dim oCacheSub

    lErrNumber = NO_ERR

    lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
    If lErrNumber <> NO_ERR Then
    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "AddPreviousAnswerToCache", "", "Error loading sCacheXML", LogLevelError)
    Else
		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetSubscriptionXML, oSubDOM)
        If lErrNumber <> NO_ERR Then
        	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "AddPreviousAnswerToCache", "", "Error calling LoadXMLDOMFromString", LogLevelTrace)
        End If
    End If

    If lErrNumber = NO_ERR Then
		'set subID into cacheXML
		Set oCacheSub = oCacheDOM.selectSingleNode("/mi/sub")
		Call oCacheSub.setAttribute("subid", oSubDOM.selectSingleNode("/SubscriptionInfo/subscription/SUBSCRIPTION_ID").text)

		Set oQuestions = oCacheDOM.selectNodes("/mi/qos/mi/in/oi[@tp = '" & TYPE_QUESTION & "']")
        For Each oCurrentQuestion in oQuestions
            sQOID = oCurrentQuestion.getAttribute("id")
			sPrefID = oSubDOM.selectSingleNode("/SubscriptionInfo/personalization/qo[@id = '" & sQOID & "']/PREFERENCE_ID").text

			Set oAnswer = oCurrentQuestion.selectSingleNode("answer")
			If Not oAnswer is Nothing Then
				Call oCurrentQuestion.removeChild(oAnswer)
			End If
			Set oAnswer = oCurrentQuestion.appendChild(oCacheDOM.createElement("answer"))

            'If Strcomp(oAnswer.getAttribute("prefID"), sPrefID) <> 0 Or oAnswer.selectSingleNode("*") Is Nothing Then
                lErrNumber = cu_GetProfile(sPrefID, sQOID, sProfileXML)
				If lErrNumber <> NO_ERR Then
		        	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PersonalizeCuLib.asp", "AddPreviousAnswerToCache", "", "Error calling cu_GetProfile", LogLevelTrace)
				Else
					Call LoadXMLDOMFromString(aConnectionInfo, sProfileXML, oProfileDOM)

					Set oProfile = oProfileDOM.selectSingleNode("/mi/in/oi")
					If Not (oProfile Is Nothing) Then
					    sProfileName = oProfile.getAttribute("n")
					    sProfileDesc = oProfile.getAttribute("des")
					    'If StrComp(sProfileDesc, "null", vbBinaryCompare) = 0 Then
						'	sProfileDesc = ""
						'End If
					    oAnswer.setAttribute "n", sProfileName
					    oAnswer.setAttribute "desc", sProfileDesc
					Else
					    oAnswer.setAttribute "n", ""
					    oAnswer.setAttribute "desc", ""
					End If

					oAnswer.setAttribute "prefID", sPrefID
					Set oPref = oSubDOM.selectSingleNode("/SubscriptionInfo/personalization/qo[@id = '" & sQOID & "']/PROMPT_ANSWER/*")
					If Not oPref is Nothing Then
						oAnswer.appendChild(oPref)
					End If
				End If
            'End If
        Next
    End If

    sCacheXML = oCacheDOM.xml

    Set oCacheDOM = Nothing
    Set oQuestions = Nothing
    Set oCurrentQuestion = Nothing
    Set oAnswer = Nothing
    Set oPref = Nothing
    Set oSubDOM = Nothing
    Set oProfileDOM = Nothing
    Set oProfile = Nothing


    AddPreviousAnswerToCache = lErrNumber
    Err.Clear
End Function


Function AddQuestionDetailsToCache(sCacheXML, sGetDetailsForQuestionsXML)
'********************************************************
'*Purpose:	Add Question Details
'*Inputs:	sCacheXML, sGetDetailsForQuestionsXML
'*Outputs:	sCacheXML
'********************************************************
	On Error Resume Next
	Const CONNECTION_INFO = "connInfo"
	Const SECURITY_OBJECT = "secObj"
	Const AUTHENTICATION_OBJECT = "authObj"
	Dim lErrNumber
	Dim oCacheDOM
	Dim oDetailsDOM
	Dim oCurrentCacheQuestion
	Dim oDecoder
	Dim oPropDOM
	Dim oPropertyNode
	Dim sGetUserSecurityObjectsXML
	Dim sGetUserAuthenticationObjectsXML
	Dim oDecodeDOM
	Dim sEncodedData
	Dim sQuestionID
	Dim oUserSecurityDOM
	Dim oInfoSourceOI
	Dim oUserAuthenticationDOM
	Dim sInfoSourceID
	Dim oFromInfoSouceOI

	lErrNumber = NO_ERR
	lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
	If lErrNumber <> NO_ERR Then
		Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
	Else
		lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetDetailsForQuestionsXML, oDetailsDOM)
		If lErrNumber <> NO_ERR Then
			Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
		End If
	End If

	If lErrNumber = NO_ERR Then
	    If (oCacheDOM.selectSingleNode("/mi/in")) Is Nothing Then	'add all QO defintion at first time
	        oCacheDOM.selectSingleNode("/mi").appendChild(oDetailsDOM.selectSingleNode("/mi/in"))

			Set oDecoder = Server.CreateObject(PROGID_BASE64)
			Call GetXMLDOM(aConnectionInfo, oDecodeDOM, sErrDescription)

			lErrNumber = cu_GetUserSecurityObjects(sGetUserSecurityObjectsXML)
			If lErrNumber <> NO_ERR Then
			    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "", "Error calling cu_GetUserSecurityObjects", LogLevelTrace)
			Else
				lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetUserSecurityObjectsXML, oUserSecurityDOM)
				If lErrNumber <> NO_ERR Then
					Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
				End If
			End If

			lErrNumber = cu_GetUserAuthenticationObjects(sGetUserAuthenticationObjectsXML)
			If lErrNumber <> NO_ERR Then
			    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "", "Error calling cu_GetUserAuthenticationObjects", LogLevelTrace)
			Else
				lErrNumber = LoadXMLDOMFromString(aConnectionInfo, sGetUserAuthenticationObjectsXML, oUserAuthenticationDOM)
				If lErrNumber <> NO_ERR Then
					Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.Description), Err.source, "PrePromptCuLib.asp", "AddQuestionDetailsToCache", "LoadXMLDOMFromString", "Error calling LoadXMLDOMFromString", LogLevelTrace)
				End If
			End If

			If lErrNumber = NO_ERR Then
				For Each oCurrentCacheQuestion in oCacheDOM.selectNodes("/mi/qos/mi/in/oi[@tp = '" & TYPE_QUESTION & "']")
					sQuestionID = oCurrentCacheQuestion.getAttribute("id")
				    oCurrentCacheQuestion.setAttribute "isid", oDetailsDOM.selectSingleNode("/mi/qos/oi[@id = '" & sQuestionID & "']").getAttribute("isid")
				    oCurrentCacheQuestion.appendChild(oDetailsDOM.selectSingleNode("/mi/qos/oi[@id = '" & sQuestionID & "']/prs"))
				    sEncodedData = oCurrentCacheQuestion.selectSingleNode("prs/pr[@n='definition']").text
				    oCurrentCacheQuestion.selectSingleNode("prs/pr[@n='definition']").text = oDecoder.Decode(sEncodedData)
				Next
				lErrNumber = Err.number
			End If

			If lErrNumber = NO_ERR Then
				For Each oInfoSourceOI in oCacheDOM.selectNodes("/mi/in/oi")
					sInfoSourceID = oInfoSourceOI.getAttribute("id")
					sEncodedData = oInfoSourceOI.selectSingleNode("prs/pr[@n = '" & CONNECTION_INFO & "']").text
					oInfoSourceOI.selectSingleNode("prs/pr[@n = '" & CONNECTION_INFO & "']").text = oDecoder.Decode(sEncodedData)

					set oFromInfoSouceOI = oUserSecurityDOM.selectSingleNode("/mi/in/oi[@id = '" & sInfoSourceID & "']")
					If Not (oFromInfoSouceOI Is Nothing) Then
					    If Len(oFromInfoSouceOI.getAttribute("v")) > 0 Then
					        Set oPropertyNode = oInfoSourceOI.selectSingleNode("prs").appendChild(oCacheDOM.createElement("pr"))
					        oPropertyNode.setAttribute "n", SECURITY_OBJECT
					        oPropertyNode.setAttribute "v", oFromInfoSouceOI.getAttribute("v")
					    End If
					End If

					set oFromInfoSouceOI = oUserAuthenticationDOM.selectSingleNode("/mi/in/oi[@id = '" & sInfoSourceID & "']")
					If Not (oFromInfoSouceOI Is Nothing) Then
					    If Len(oFromInfoSouceOI.getAttribute("v")) > 0 Then
					        Set oPropertyNode = oInfoSourceOI.selectSingleNode("prs").appendChild(oCacheDOM.createElement("pr"))
					        oPropertyNode.setAttribute "n", AUTHENTICATION_OBJECT
					        oPropertyNode.setAttribute "v", oFromInfoSouceOI.getAttribute("v")
					    End If
					End If
					lErrNumber = Err.number
				Next
			End If
		End If
	End If

	If lErrNumber = NO_ERR Then
		sCacheXML = oCacheDOM.xml
	End If

	Set oCacheDOM = Nothing
	Set oDetailsDOM = Nothing
	Set oCacheQuestions = Nothing
	Set oCurrentCacheQuestion = Nothing
	Set oDecoder = Nothing
	Set oPropertyNode = Nothing
	Set oPropDOM = Nothing
	Set oDecodeDOM = Nothing
	Set oUserSecurityDOM = Nothing
	Set oInfoSourceOI = Nothing
	Set oUserAuthenticationDOM = Nothing
	set oFromInfoSouceOI = Nothing

	AddQuestionDetailsToCache = lErrNumber
	Err.Clear
End Function

Function cu_GetDetailsForQuestions(sCacheXML, sGetDetailsForQuestionsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Const PROCEDURE_NAME = "cu_GetDetailsForQuestions"
	Dim lErrNumber
	Dim sSessionID
	Dim asQuestionObjectID()
	Dim oOutputDOM
	Dim oQuestions
	Dim i

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    Set oOutputDOM = Server.CreateObject("Microsoft.XMLDOM")
    oOutputDOM.async = False
    If oOutputDOM.loadXML(sCacheXML) = False Then
    	lErrNumber = ERR_XML_LOAD_FAILED
    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", PROCEDURE_NAME, "", "Error loading sCacheXML", LogLevelError)
    Else
        Set oQuestions = oOutputDOM.selectNodes("/mi/qos//oi[@tp = '" & TYPE_QUESTION & "']")
        If Err.number <> NO_ERR Then
            lErrNumber = Err.number
            Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", PROCEDURE_NAME, "", "Error retrieving question oi nodes", LogLevelError)
        End If
    End If

    If lErrNumber = NO_ERR Then
        If oQuestions.length > 0 Then
            Redim asQuestionObjectID(oQuestions.length - 1)
            For i = 0 To (oQuestions.length - 1)
                asQuestionObjectID(i) = oQuestions.item(i).getAttribute("id")
            Next

			lErrNumber = co_GetDetailsForQuestions(sSessionID, asQuestionObjectID, sGetDetailsForQuestionsXML)
			If lErrNumber <> NO_ERR Then
		    	Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", PROCEDURE_NAME, "", "Error calling co_GetDetailsForQuestions", LogLevelTrace)
		    End If
        End If
    End If

    Set oQuestions = Nothing
	Set oOutputDOM = Nothing

	cu_GetDetailsForQuestions = lErrNumber
	Err.Clear
End Function

Function cu_GetUserSecurityObjects(sGetUserSecurityObjectsXML)
'********************************************************
'*Purpose:
'*Inputs:
'*Outputs:
'********************************************************
	On Error Resume Next
	Dim lErrNumber
    Dim sSessionID

	lErrNumber = NO_ERR
	sSessionID = GetSessionID()

    lErrNumber = co_GetUserSecurityObjects(sSessionID, sGetUserSecurityObjectsXML)
    If lErrNumber <> NO_ERR Then
        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PrePromptCuLib.asp", "cu_GetUserSecurityObjects", "", "Error calling co_GetUserSecurityObjects", LogLevelTrace)
    End If

	cu_GetUserSecurityObjects = lErrNumber
	Err.Clear
End Function


%>