<!-- #include file="../CoreLib/ProcessUserDetailsCoLib.asp" -->
<!-- #include file="../CoreLib/PostPromptCoLib.asp" -->
<!-- #include file="../CoreLib/CommonCoLib.asp" -->
<!-- #include file="../CoreLib/PersonalizeCoLib.asp" -->
<!-- #include file="../CoreLib/PrePromptCoLib.asp" -->
<!--#include file="../CoreLib/SubscribeCoLib.asp" -->
<%

	Function GenerateAnswer(oDict,qDef)
	On Error Resume Next
	Dim oQXML
	Dim oAXML
	Dim newNode
	Dim Root
	Dim aRoot
	Dim elemNode
	Dim attrNode
	Dim name
	Dim nodeName
	Dim METHOD_NAME
	
		METHOD_NAME = "GenerateAnswer"
		
		set oQXML = Server.CreateObject("Microsoft.XMLDOM")
		set oAXML = Server.CreateObject("Microsoft.XMLDOM")
		
		oQXML.loadXML(qDef)		
		
		set Root = oQXML.documentElement
		
		Set aRoot = oAXML.createElement("UserDetail")
		set oAXML.documentElement = aRoot

		For each elemNode in root.childNodes
			Dim tempAttr
			nodeName = elemNode.nodeName
			
			set tempAttr = elemNode.attributes.getNamedItem("Name")
			if not tempAttr is nothing then name = tempAttr.nodeValue
			
			set newNode = oAXML.createElement(nodeName)
			
			set attrNode = oAXML.createAttribute("Name")
            attrNode.Value = name
            newNode.Attributes.setNamedItem attrNode

			set attrNode = oAXML.createAttribute("Value")
            attrNode.Value = oDict(name)
            newNode.Attributes.setNamedItem attrNode
			
			aRoot.appendChild newNode
			call ReportError(METHOD_NAME,"appending a new node")
		Next 
				
		generateAnswer = oAXML.xml
	End Function
	
	Function transformDefinition(qoID,ISID,qDefn,prefID,aXML)
	On Error Resume Next
	Dim METHOD_NAME
	Dim oXML
	Dim oAXML
	Dim oDict
	Dim oADict
		
		METHOD_NAME = "transformDefinition"
		
		set oXML = Server.CreateObject("Microsoft.XMLDOM")
		set oAXML = Server.CreateObject("Microsoft.XMLDOM")

		oXML.loadXML(qDefn)
		call ReportError(METHOD_NAME,"loading question definition")
		oAXML.loadXML(aXML)
		call ReportError(METHOD_NAME,"loading answer XML")
		
		Set oDict = Server.CreateObject("Scripting.Dictionary")
		Set oADict = Server.CreateObject("Scripting.Dictionary")
		
		If oAXML.hasChildNodes() then
			call collectAnswerValues(oAXML,oADict)
		End If
		call ReportError(METHOD_NAME,"collecting answer values")
		
		If oXML.hasChildNodes() then
			call collectQuestionValues(oXML,oDict)
		End If
		call ReportError(METHOD_NAME,"collecting question values")
			
		transformDefinition = generateResult(qoID,ISID,qDefn,prefID,oDict,oADict)
		
	set oXML = nothing
	set oAXML = nothing

	End Function

	Function GenerateResult(qoID,ISID,qDef,prefID,oDict,oADict)
	Dim METHOD_NAME
	Dim result
	Dim answer
	Dim options
	Dim value
	Dim element
		
		On Error Resume Next
	
		METHOD_NAME = "GenerateResult"
		
		For each element in oDict
			answer = oADict(element)
			if oDict(element) <> "" then
				'there are options, so show a list box
				options = getArrayFromString(oDict(element),";")
				dim i
				result = result & "<B>" & GetElementDescriptor(element) & "</B>"
				result = result & "<BR>"
				result = result & "<TABLE>"
				result = result & "<select" & " name=" & """" & element & """" &  " value=" & """" & element & """" & " size=" & """" & "1" & """" & ">"
				for i=0 to UBound(options)
					value = options(i)
					if(answer = value) then
						result = result & "<option selected=" & """" & "1" & """" &  " value=" & """" & value & """" & ">" & value & "</option>"
					else
						result = result & "<option value=" & """" & value & """" & ">" & value & "</option>"
					end if
				next
				result = result & "</select>"
				result = result & "</TABLE>"
				result = result & "<BR>"
			else
				'no options just show a simple text box
				'Filter out the below mentioned hard coded user detail object
				'If strcomp(element,"All (returns XML string with all user details)",vbTextCompare) <> 0 OR strcomp(element,"All User Info (XSL required)",vbTextCompare) <> 0 OR strcomp(element, "Custom Combination",vbTextCompare) <> 0 Then
				If strcomp(element, "Custom Combination",vbTextCompare) <> 0 Then
					result = result & "<B>" & GetElementDescriptor(element) & "</B>"
					result = result & "<BR>"
					result = result & "<TABLE>"
					result = result & "<Input" & " name=" & """" & element & """" & " value=" & """" & answer & """" & ">" & "</Input>"
					result = result & "</TABLE>"
					'result = result & "<Input" & " name=" & """" & element & """" & " value=" & """" & answer & """" & ">" & element & "</Input>"
					result = result & "<BR>"
					result = result & "<BR>"
				End IF
			end if
			
		Next
		result = result & "<BR>"
		result = result & "<Input" & " name=" & """" & "storeUDAns" & """" &  " type=" & """" & "hidden" & """" & " value=" & """" & "1" & """" & " />"
		result = result & "<Input" & " name=" & """" & "qoID" & """" &  " type=" & """" & "hidden" & """" & " value=" & """" & qoID & """" & " />"
		result = result & "<Input" & " name=" & """" & "ISID" & """" &  " type=" & """" & "hidden" & """" & " value=" & """" & ISID & """" & " />"
		If prefID <> "" Then
			result = result & "<Input" & " name=" & """" & "edit" & """" &  " type=" & """" & "hidden" & """" & " value=" & """" & "1" & """" & ">"
			result = result & "<Input" & " name=" & """" & "prefID" & """" &  " type=" & """" & "hidden" & """" & " value=" & """" & prefID & """" & ">"
		End If
		result = result & "<Input" & " name=" & """" & "qDef" & """" &  " type=" & """" & "hidden" & """" & " value=" & """" & Server.HTMLEncode(qDef) & """" & " />"
		result = result & "</FONT>"
		result = result & "<Input" & " name=" & """" & "Submit" & """" &  " type=" & """" & "Submit" & """" & " value=" & """" & asDescriptors(59) & """" &   " class=" & """" & "buttonClass"    &   """"  &  " />" 'Descriptor :Save
		result = result & "<BR>"
		generateResult = result
	End Function
	
	Function GetElementDescriptor(element)
	
		Dim sElementDescriptor
		Dim sElementTrimed
		sElementTrimed = Trim(element)
	
		Select Case sElementTrimed 
			Case "First Name"
				sElementDescriptor = asDescriptors(959)
			Case "Middle Initial"
				sElementDescriptor = asDescriptors(961)
			Case "Last Name"
				sElementDescriptor = asDescriptors(960)
			Case "Suffix"
				sElementDescriptor = asDescriptors(962)
			Case "Title"
				sElementDescriptor = asDescriptors(963)
			Case "Salutation"
				sElementDescriptor = asDescriptors(964)
			Case "Street Address"
				sElementDescriptor = asDescriptors(965)
			Case "City"
				sElementDescriptor = asDescriptors(966)
			Case "State"
				sElementDescriptor = asDescriptors(967)
			Case "Zip Code"
				sElementDescriptor = asDescriptors(968)
			Case "Country"
				sElementDescriptor = asDescriptors(969)
			Case Else
				sElementDescriptor = element
		End Select
	
		GetElementDescriptor = sElementDescriptor	
	End Function
	
	Function SaveProfileAndPreferenceAnswer(qoID,ISID,ansXML)
	Dim sPreferenceID
	Dim lErr
	Dim sessionID
	On Error Resume Next
		sPreferenceID = GetGUID()
		sessionID = GetSessionID()
		lErr = co_CreateProfile(sessionID, sPreferenceID, qoID, ISID, "DefaultUserDetails", "Default profile created for the user details question", true)	
		If lErr <> NO_ERR then
			call ReportError("SaveProfileAndPreferenceAnswer","Calling co_CreateProfile")
		Else
			lErr = SavePreferenceObject(sessionID,qoID,ISID,sPreferenceID,ansXML)
			If lErr <> NO_ERR then
				call ReportError("SaveProfileAndPreferenceAnswer","Calling SavePreferenceObject")
			End IF
		end if
	End Function
	
	Function SavePreferenceAnswer(qoID,ISID,sPreferenceID,ansXML)
	Dim lErr
	Dim sessionID
	On Error Resume Next
		sessionID = GetSessionID()
		lErr = SavePreferenceObject(sessionID,qoID,ISID,sPreferenceID,ansXML)
		If lErr <> NO_ERR then
			call ReportError("SavePreferenceAnswer","Calling SavePreferenceObject")
		End IF
	End Function

	Function SavePreferenceObject(sessionID,qoID,ISID,sPreferenceID,ansXML)
		On Error Resume Next
		Dim sPrefDataXML
		Dim lErr
		
        sPrefDataXML = sPrefDataXML & "<updatePreferenceObjects>"
        sPrefDataXML = sPrefDataXML & "<subsetID>" & "" & "</subsetID>"
        sPrefDataXML = sPrefDataXML & "<serviceID>" & "" & "</serviceID>"
        sPrefDataXML = sPrefDataXML & "<sessionID>" & sessionID & "</sessionID>"
        sPrefDataXML = sPrefDataXML & "<channelID>" & "" & "</channelID>"
        sPrefDataXML = sPrefDataXML & "<personalization>"
        sPrefDataXML = sPrefDataXML & "<qo id='" & qoID & "'>"
        sPrefDataXML = sPrefDataXML & "<INFO_SOURCE_ID>" & ISID & "</INFO_SOURCE_ID>"
        sPrefDataXML = sPrefDataXML & "<PREFERENCE_ID>" & sPreferenceID & "</PREFERENCE_ID>"
        sPrefDataXML = sPrefDataXML & "<PROMPT_ANSWER>"
        sPrefDataXML = sPrefDataXML & Server.HTMLEncode(CStr(ansXML))
        sPrefDataXML = sPrefDataXML & "</PROMPT_ANSWER>"
        sPrefDataXML = sPrefDataXML & "</qo>"
        sPrefDataXML = sPrefDataXML & "</personalization>"
        sPrefDataXML = sPrefDataXML & "</updatePreferenceObjects>"
		
		If lErr = NO_ERR Then
		    lErr = co_UpdatePreferenceObjects(sPrefDataXML)
		    If lErr <> NO_ERR Then
		        Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), CStr(Err.source), "PostPromptCuLib.asp", "UpdatePreferenceObject_Profile", "", "Error while calling co_UpdatePreferenceObjects", LogLevelTrace)
		    End If
		End If
	End Function
	
	Function getArrayFromString(inputString,separator)
	Dim ind
	Dim prevInd
	Dim oArray
	Dim index
	Dim METHOD_NAME
		
	On Error Resume Next
		METHOD_NAME = "getArrayFromString"
		index=0
		'the string starts with the separator
		if mid(inputString,1,1)=separator then
			inputString = mid(inputString,2)
		end if
		ind=0
		prevInd = 0
		while(ind >= 0)
			ind = instr(prevInd+1,inputString,separator)
			if ind > 0 then
				if index=0 then 
					redim oArray(index)
				else
					redim preserve oArray(index)
				end if
				oArray(index) = mid(inputString,prevInd+1,ind-prevInd-1)
				index = index + 1
				prevInd = ind
			else
				redim preserve oArray(index)
				oArray(index) = mid(inputString,prevInd+1)
				ind=-1
			end if
		wend
		call ReportError(METHOD_NAME,"generating options array")
		getArrayFromString = oArray
	End Function
	
	Function ReportError(methodName,stepName)
		If Err.number <> 0 then
            Call LogErrorXML(aConnectionInfo, CStr(Err.number), CStr(Err.Description), Err.Source, "ProcessUserDetailsCuLib.asp", methodName, "", stepName, LogLevelTrace)
		End If
	End Function
	
	Function collectAnswerValues(oXMLDOM,oDict)
	On Error Resume Next
		Dim root
		Dim elemNode
		Dim elemAttr
		Dim METHOD_NAME
		Dim attrString
		Dim name
		Dim value
		Dim found
		Dim oName
		Dim oValue
		
		METHOD_NAME = "collectAnswerValues"
		
		'go through the collection and pick the values
		set root = oXMLDom.documentElement
		If root.nodeName = "UserDetail" then
			For each elemNode in root.childNodes
				attrString = ""
				name=""
				value=""
				found=0
				set oName = elemNode.attributes.getNamedItem("Name")
				set oValue = elemNode.attributes.getNamedItem("Value")
				if not oName is nothing then name = oName.nodeValue
				if not oValue is nothing then value = oValue.nodeValue
				call ReportError(METHOD_NAME,"getting property value from XML")
				call oDict.Add(name,value)
			Next
		Else
			'raise error
			call ReportError(METHOD_NAME,"invalid user details XML")
		End If
		call ReportError(METHOD_NAME,"after adding properties to the collection")
	End Function

	Function collectQuestionValues(oXMLDOM,oDict)
	On Error Resume Next
		Dim root
		Dim elemNode
		Dim elemAttr
		Dim METHOD_NAME
		Dim attrString
		Dim name
		Dim value
		Dim found
		Dim oName
		Dim oValue
		Dim optionNode
		
		METHOD_NAME = "collectQuestionValues"
		
		'go through the collection and pick the values
		set root = oXMLDom.documentElement
		If root.nodeName = "UserDetail" then
			For each elemNode in root.childNodes
				attrString = ""
				name=""
				value=""
				found=0
				set oName = elemNode.attributes.getNamedItem("Name")
				set oValue = elemNode.attributes.getNamedItem("Value")
				if not oName is nothing then name = oName.nodeValue
				if not oValue is nothing then value = oValue.nodeValue
				call ReportError(METHOD_NAME,"getting property value from XML")
				if elemNode.hasChildNodes then
					for each optionNode in elemNode.childNodes
						value = value & ";" & optionNode.text
					next
				end if
				call oDict.Add(name,value)
			Next
		Else
			'raise error
			call ReportError(METHOD_NAME,"invalid user details XML")
		End If
		call ReportError(METHOD_NAME,"after adding properties to the collection")
	End Function

	Function GetQuestionDefinition(qoID,qDef)
	Dim lErr
	Dim aQuestionIDs(0)
		aQuestionIDs(0) = qoID
		lErr = co_GetDetailsForQuestions(getSessionID(), aQuestionIDs, qDef)
		GetQuestionDefinition = lErr
	End Function
	
	Function GetUserDetailPreferenceDefinition(qoID,prefID,sAXML)
	Dim lErr
	Dim pList
	Dim aQuestionIDs(0)
	Dim pDef
	Dim oXML
	Dim udNode
	Dim METHOD_NAME
		On Error Resume Next
	
		METHOD_NAME = GetUserDetailPreferenceDefinition
	
		set oXML = Server.CreateObject("MICROSOFT.XMLDOM")

		aQuestionIDs(0) = qoID
		lErr = co_GetUserDefaultPersonalization(getSessionID(),pList)

		'Now check if there are answers corresponding to this particular question
		call GetPreferenceInformation(pList,qoID,prefID)
		call ReportError(METHOD_NAME,"getting the list of profiles for question " & qoID)

		If prefID <> "" Then 
			call GetPreferenceDefinition(prefID,pDef)
		Else
			pDef = ""
		End If
		call ReportError(METHOD_NAME,"getting the preference definition")
		
		oXML.loadXML pDef
		set udNode = oXML.selectSingleNode("//mi/in/oi/UserDetail")
		if not udNode is nothing then
			sAXML = udNode.xml
		else
			sAXML = ""
		end if
		GetQuestionDefinition = lErr
	End Function

	Function GetPreferenceDefinition(prefID,pDef)
	Dim lErr
	Dim tempArray(0)
		tempArray(0) = prefID
		lErr = co_GetPreferenceObjects(GetSessionID(), tempArray, pDef)
	End Function

	Function GetPreferenceInformation(pList,qoID,prefID)
	Dim lErr
	Dim oXMLDOM
	Dim root
	Dim prefNode
	Dim qNode
	Dim METHOD_NAME
	
	On Error Resume Next
		METHOD_NAME = GetPreferenceInformation
	
		set oXMLDOM = Server.CreateObject("MICROSOFT.XMLDOM")
		oXMLDOM.loadXML pList
		call ReportError(METHOD_NAME,"loading the list of preferences")
		set root = oXMLDOM.documentElement
		If root.hasChildNodes Then
			If root.firstChild.nodeName = "qos" then
				If root.firstChild.hasChildNodes Then
					set qNode = root.firstChild.selectSingleNode("oi[@tp = '" & TYPE_QUESTION & "' $and$ @id='" & qoID & "']")
					if not qNode is nothing then
						set prefNode = qNode.selectSingleNode("mi/in/oi")
						if not prefNode is nothing then
							prefID = prefNode.attributes.getNamedItem("id").nodeValue
						end if
					end if
				End If
			Else
				prefID = ""
			end If
		Else
			prefID = ""
		End If
		call ReportError(METHOD_NAME,"parsing through the question XML definition")
				
		GetPreferenceInformation = lErr
	End Function
	
	Function GetISDefinition(ISID,ISDefn)
		Dim oXML
		Dim xmlDef
		Dim defNode
		Dim oDecoder

		ISDefn = ""
		lErr = co_GetInformationSourceDefinition(GetSessionID(),ISID,xmlDef)
		if lErr = 0 and xmlDef <> "" then
			set oXML = Server.CreateObject("MICROSOFT.XMLDOM")
			oXML.loadXML xmlDef
			set defNode = oXML.selectSingleNode("//mi/in/oi/mstr_item/is_definition")
			if not defNode is nothing then
				ISDefn = defNode.text
				set oDecoder = Server.CreateObject(PROGID_BASE64)
				if not oDecoder is nothing then
					ISDefn = oDecoder.decode(ISDefn)
				end if
			end if
		End if
	End Function
%>