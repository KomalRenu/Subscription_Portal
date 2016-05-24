<%
Public Const SESSION_PROGID = "MSIXMLLib.DSSXMLServerSession.1"
Public Const XMLADMIN_PROGID = "MSIXMLLib.DSSXMLAdmin.1"

Function coLoginUser(objSession, strProjName, strUserName, strPassword, _
				   strNewPassword, lngAuthMode, lngLocale, strClientID, _
				   lngSessFlag, strSessID, strErrDesc)
On Error Resume Next
	strSessID = objSession.CreateSession(strUserName, strPassword, _
					strNewPassword, strProjName, lngLocale, _
					strClientID, lngAuthMode, lngSessFlag)
	If Err.number <> 0 Then
		coLoginUser = Err.Number
		strErrDesc = "Unable to create a Server Session object: " & Err.description
		Call LogErrorXML(aConnectionInfo, Err.Number, Err.description, Err.source, "ISLoginCoLib.asp", "coLoginUser", "CreateSession", "Error calling CreateSession", LogLevelError)
		Err.Clear
	End If

End Function


Function closeSession(sessionObj,strSessID)
	If strSessID <> "" Then
		call sessionObj.closeSession(strSessID,0)
	End If
	If Err.number <> 0 Then
		strErrDesc = "Error closing session"
		closeSession = Err.Number
		Call LogErrorXML(aConnectionInfo, Err.Number, Err.description, Err.source, "ISLoginCoLib.asp", "closeSession", "closeSession", strErrDesc, LogLevelError)
		Err.Clear
	End If
	Set sessionObj = Nothing
	closeSession = 0
End Function

'Get the Server Session object...
Function GetSessionObj(objSession, strErrDesc)

	On Error Resume Next

	'Create a new Server Session object...
	Set objSession = Server.CreateObject(SESSION_PROGID)

	If Err.number <> 0 Then
		strErrDesc = "Error closing session"
		GetSessionObj = Err.Number
		Call LogErrorXML(aConnectionInfo, Err.Number, Err.description, Err.source, "ISLoginCoLib.asp", "GetSessionObj", "GetSessionObj", "Error creating session object", LogLevelError)
		Err.Clear
	End If

	objSession.ApplicationType = APPLICATION_TYPE_PORTAL

	GetSessionObj = Err.number
	Err.Clear
End Function

'Get a Server session object initialized with some basic information...
Function InitSessionObj(strServerName, lngPortNumber, objSession, strErrDesc)

	On Error Resume Next
	'Try to get a Server Session object...
	Dim lngErrCode
	lngErrCode = GetSessionObj(objSession, strErrDesc)

	objSession.ServerName = strServerName

	'Did we get an error?
	If Err.number <> 0 Then
		InitSessionObj = Err.number
		strErrDesc = "Unable to set ServerName property: " & Err.description
		Call LogErrorXML(aConnectionInfo, Err.Number, Err.description, Err.source, "ISLoginCoLib.asp", "GetSessionObj", "GetSessionObj", strErrDesc, LogLevelError)
		Err.Clear
		Exit Function
	End If

	'Set the port number...
	objSession.Port = lngPortNumber

	'Did we get an error?
	If Err.number <> 0 Then
		InitSessionObj = Err.number
		strErrDesc = "Unable to set Port property: " & Err.description
		Call LogErrorXML(aConnectionInfo, Err.Number, Err.description, Err.source, "ISLoginCoLib.asp", "GetSessionObj", "GetSessionObj", strErrDesc, LogLevelError)
		Err.Clear
		Exit Function
	End If

	'Set the Application Type to DSSApplicationPortal...
	objSession.ApplicationType = APPLICATION_TYPE_PORTAL

	'Did we get an error?
	If Err.number <> 0 Then
		InitSessionObj = Err.number
		strErrDesc = "Unable to set ApplicationType property: " & Err.description
		Call LogErrorXML(aConnectionInfo, Err.Number, Err.description, Err.source, "ISLoginCoLib.asp", "GetSessionObj", "GetSessionObj", strErrDesc, LogLevelError)
		Err.Clear
		Exit Function
	End If

	InitSessionObj = Err.number
	Err.Clear
End Function



Function GetProjects(sessionObj,userName,passwd,lAuthMode,projectsXML)
On Error Resume Next
	projectsXML = sessionObj.getUserProjects(userName,passwd,lAuthMode,0)
	If Err.number <> 0 Then
		GetProjects = Err.number
		strErrDesc = "Unable to get the list of projects." & Err.description
		Call LogErrorXML(aConnectionInfo, Err.Number, Err.description, Err.source, "ISLoginCoLib.asp", "GetProjects", "GetProjects", strErrDesc, LogLevelError)
		Err.Clear
		Exit Function
	End If
	GetProjects = 0
End Function


Function GetProjectName(projectsXML,asProjectName)
Dim strErrDesc
'On Error Resume Next
If projectsXML <> "" Then
	Dim oXML
	Dim oProjects
	Dim oProject
	Dim i
	Dim lProjectCount

	Set oXML = Server.createObject("MICROSOFT.XMLDOM")
	oXML.loadXML(projectsXML)
	if not oXML is nothing then
		Set oProjects = oXML.selectNodes("//mi/srps//sp[@ps='0']")
		If Err.Number <> 0 Then
			GetProjectName = Err.number
			strErrDesc = "Error parsing through the project list"
			Call LogErrorXML(aConnectionInfo, Err.Number, Err.description, Err.source, "ISLoginCoLib.asp", "GetProjectName", "GetProjectName", strErrDesc, LogLevelError)
			Err.Clear
			Exit Function
		End If

		lProjectCount = oProjects.length
		If lProjectCount > 0 Then
			Redim asProjectName(lProjectCount)
			For i = 1 to lProjectCount
				Set oProject = oProjects.nextNode
				asProjectName(i) = oProject.attributes.getNamedItem("pn").nodeValue
			Next
		End if

		If Err.Number <> 0 Then
			GetProjectName = Err.number
			strErrDesc = "Error reading the project node"
			Call LogErrorXML(aConnectionInfo, Err.Number, Err.description, Err.source, "ISLoginCoLib.asp", "GetProjectName", "GetProjectName", strErrDesc, LogLevelError)
			Err.Clear
			Exit Function
		End If
	end if
End If
GetProjectName = Err.number
End Function


Function GetUserSession(sessionObj,strServerName,strPortNumber,strUserName,strPassword,lAuthMode,sessionID,strErrDesc)
Dim projectsXML
Dim asProjectName
Dim lngErrCode
Dim strErrMsg
Dim i


	lngErrCode = InitSessionObj(strServerName, strPortNumber, sessionObj, _
								   strErrMsg)

	if lngErrCode = 0 then

		lngErrCode = GetProjects(sessionObj,strUserName,strPassword,lAuthMode,projectsXML)
		if lngErrCode = 0 then
			lngErrCode = GetProjectName(projectsXML,asProjectName)
			if lngErrCode = 0 then
				For i = 1 to UBound(asProjectName)
					lngErrCode = coLoginUser(sessionObj, asProjectName(i), strUserName, strPassword, "",_
								   lAuthMode, GetLng(), "127.0.0.1", 0, sessionID, StrErrDesc)
					If lngErrcode = 0 then
						Exit For
					End If
				Next
			end if
		end if
	end if
	GetUserSession = lngErrCode
End Function



Function GetUserInfo(sessionObj,sessionID,userInfo,strErrDesc)
Dim userID
Dim userName
Dim oInfo
On Error Resume Next
	Redim userInfo(2)
	If sessionID <> "" Then
		call sessionObj.getUserInfo(sessionID,userID,userName,oInfo)
		If Err.number <> 0 Then
			GetUserInfo = Err.Number
			strErrDesc = "Unable to create a Server Session object: " & Err.description
			Call LogErrorXML(aConnectionInfo, getUserInfo, Err.description, Err.source, "ISLoginCoLib.asp", "getUserInfo", "getUserInfo", "Error calling getUserInfo", LogLevelError)
			Err.Clear
			Exit Function
		End If
		userInfo(0) = userID
		userInfo(1) = userName
		userInfo(2) = oInfo
	End IF
	GetUserInfo = Err.number
End Function

'Dim sessionID
'Dim userInfo()
'Dim strErrDesc
'Dim sessionObj

'	Response.Write GetSessionObj(sessionObj, strErrDesc)
'	Response.Write "<BR>"
'	Response.Write strErrDesc
'	Response.write GetUserSession(sessionObj,"malkovich","administrator","",1,sessionID,strErrDesc)
'	Response.Write "<BR>"
'	Response.Write strErrDesc
'	Response.Write "session = " & sessionID
'	Response.Write "<BR>"
'	Response.write GetUserInfo(sessionObj,sessionID,userInfo,strErrDesc)
'	Response.Write "<BR>"
'	Response.Write strErrDesc
'	Response.Write "userID = " & userInfo(0)
'	Response.Write "<BR>"
'	Response.Write "userName = " & userInfo(1)
'	Response.Write "<BR>"
'	Response.Write "oInfo = " & userInfo(2)
'	Response.Write "<BR>"
'	Response.write closeSession(sessionObj,sessionID)
'	Response.Write "<BR>"
%>