<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/PrePromptCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sCacheXML
	'Dim sGetDetailsForQuestionsXML
	Dim sGetPreferenceObjectsXML
	Dim sSubGUID
	Dim sQOID
	Dim sSRC
	Dim sFolderID
	Dim sStatusFlag
	Dim asPrefObj()
	Dim sPrefObjID
	Dim bHiddenQO
	Dim sISID

	sSubscriptionsStyle = ""

	lErr = ParseRequestForPrePrompt(oRequest, sSubGUID, sQOID, sSRC, sFolderID, sPrefObjID)

	If lErr = NO_ERR Then
	    lErr = ReadCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
        If lErr = NO_ERR Then
            lErr = GetStatusFlag(sCacheXML, sStatusFlag)
        End If
	End If

	'If lErr = NO_ERR Then
	    'lErr = cu_GetDetailsForQuestions(sCacheXML, sGetDetailsForQuestionsXML)
	    'If lErr = NO_ERR Then
	    '   lErr = GetQuestionProperty(sGetDetailsForQuestionsXML, sQOID, bHiddenQO, sISID)
			'If lErr = NO_ERR Then
			'	If bHiddenQO Then
			'		lErr = AnswerHiddenQO(sCacheXML, sQOID, sISID)
			'		If lErr = NO_ERR Then
			'		    lErr = WriteCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
			'			If lErr = NO_ERR Then
			'				Response.Redirect "PostPrompt.asp?action=next&subGUID=" & sSubGUID & "&qoid=" & sQOID
			'			End If
			'		End If
			'	Else
					'lErr = AddQuestionDetailsToCache(sCacheXML, sGetDetailsForQuestionsXML)
					If lErr = NO_ERR Then
					    If Len(sPrefObjID) > 0 Then
					        Redim asPrefObj(0)
					        asPrefObj(0) = sPrefObjID
					        lErr = cu_GetPreferenceObjects(asPrefObj, sGetPreferenceObjectsXML)
					        If lErr = NO_ERR Then
					            lErr = AddProfileAnswerToCache(sCacheXML, sQOID, sPrefObjID, sGetPreferenceObjectsXML)
					        End If
					    End If
					    If lErr = NO_ERR Then
					        lErr = WriteCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
					        If lErr = NO_ERR Then
					            Response.Redirect "prompt.asp?subGUID=" & sSubGUID & "&qoid=" & sQOID & "&src=" & sSRC
					        End If
					    End If
					End If
				'End If
			'End If
	    'End If
	'End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE></TITLE>
</HEAD>
<BODY BGCOLOR="ffffff" TOPMARGIN="0" LEFTMARGIN="0" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin search box -->
				<!-- #include file="searchbox.asp" -->
			<!-- end search box -->
			<!-- begin left menu -->
			<BR /><BR />
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" VALIGN="TOP">
			<!-- begin center panel -->
			<BR />
			<%
			    If lErr <> NO_ERR Then
			        Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(383), "services.asp") 'Descriptor: Back to Services
			    End If
			%>
			<!-- end center panel -->
		</TD>
		<TD WIDTH="1%">
			<IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
	</TR>
</TABLE>
<BR />
<!-- begin footer -->
	<!-- #include file="footer.asp" -->
<!-- end footer -->
</BODY>
</HTML>
<%
    Erase asPrefObj
%>