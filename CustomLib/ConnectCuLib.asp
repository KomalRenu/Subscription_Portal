<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Private Const ERR_CONNECTION_LOST = 31
Private Const ERR_NOT_LOGGED =286

Function SetConnectionInfo(oRequest, aConnectionInfo)

	On Error Resume Next

	aConnectionInfo(S_IP_ADDRESS_CONNECTION) = CStr(Request.ServerVariables("REMOTE_ADDR"))
	aConnectionInfo(S_PROJECT_URL_CONNECTION) = "Server=" & Server.URLEncode(aConnectionInfo(S_SERVER_NAME_CONNECTION)) & "&Project=" & Server.URLEncode(aConnectionInfo(S_PROJECT_CONNECTION)) & "&Port=" & Server.URLEncode(aConnectionInfo(N_PORT_CONNECTION)) & "&Uid=" & Server.URLEncode(aConnectionInfo(S_UID_CONNECTION)) & "&UMode=" & Server.URLEncode(aConnectionInfo(N_USER_MODE_CONNECTION))

	aConnectionInfo(S_UID_CONNECTION) = GetSessionID()
	aConnectionInfo(S_TOKEN_CONNECTION) = GetSessionID()   'For the moment, we use the UserId as Connection Token
	aConnectionInfo(S_SITEID_CONNECTION) = "site1"

	SetConnectionInfo = Err.number
	Err.Clear

End Function

%>