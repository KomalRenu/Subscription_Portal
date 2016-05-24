<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/PostPromptCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sCacheXML
	Dim sSubGUID
	Dim sQOID
	Dim sSource
	Dim sNextQOID
	Dim sProfileName
	Dim sOriginalProfileName
	Dim sPrefID
	Dim sNextPrefID
	Dim bIsExistingProfile
	Dim sPrefDesc
	Dim URLString
	Dim sISID
	Dim bIsDefault
	Dim bHasPrefDef
	Dim sServiceID

	sSubscriptionsStyle = ""
	bIsExistingProfile = False

	lErr = ParseRequestForPostPrompt(oRequest, sSubGUID, sQOID, sSource)

	If lErr = NO_ERR Then
	    lErr = ReadCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
	    If lErr = NO_ERR Then
	        lErr = CheckForProfile(sCacheXML, sServiceID, sQOID, sISID, sProfileName, sOriginalProfileName, sPrefID, sPrefDesc, bIsDefault, bIsExistingProfile, bHasPrefDef)
	        If lErr = NO_ERR Then
	            If Len(sProfileName) > 0 Then
                    If bIsExistingProfile = True Then
					    If bHasPrefDef = True Then	'pick by profile
							lErr = cu_UpdatePreferenceObjects(sCacheXML, sQOID, sPrefID, sServiceID)	'update Pref Object Defintion
							If lErr = NO_ERR Then
								lErr = cu_UpdateProfile(sPrefID, sQOID, sISID, sProfileName, sPrefDesc, bIsDefault)	'update profile definition
							End If
						End If
	                Else
	                    lErr = cu_CreateProfile(sPrefID, sQOID, sISID, sProfileName, sPrefDesc, bIsDefault)
	                    If lErr = NO_ERR Then
	                        lErr = UpdateCache_CreateProfile(sCacheXML, sQOID, sPrefID, sProfileName, sPrefDesc, bIsDefault)
	                    End If
	                End If
	                If lErr = NO_ERR Then
						lErr = WriteCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
                    End If
	            End If
	        End If

	        If lErr = NO_ERR Then
	            If StrComp(sSource, "personalization", vbBinaryCompare) = 0 Then
	                Response.Redirect "personalize.asp?eSGUID=" & sSubGUID
	            Else
	                If StrComp(CStr(oRequest("action")), "next", vbBinaryCompare) = 0 Then
	                    lErr = GetNextQuestionObject(sQOID, sCacheXML, sNextQOID, sNextPrefID)
	                    If lErr = NO_ERR Then
	                        If StrComp(sNextQOID, "last", vbBinaryCompare) = 0 Then
	                            Response.Redirect "modify_subscription.asp?subGUID=" & sSubGUID
	                        Else
	                            URLString = "PrePrompt.asp?subGUID=" & sSubGUID & "&qoid=" & sNextQOID
	                            If Len(sNextPrefID) > 0 Then
	                                URLString = URLString & "&prefID=" & sNextPrefID
	                            End If
	                            Response.Redirect URLString
	                        End If
	                    End If
	                ElseIf StrComp(CStr(oRequest("action")), "back", vbBinaryCompare) = 0 Then
	                    lErr = GetPreviousQuestionObject(sQOID, sCacheXML, sNextQOID, sNextPrefID)
	                    If lErr = NO_ERR Then
	                        If StrComp(sNextQOID, "first", vbBinaryCompare) = 0 Then
	                            Response.Redirect "personalize.asp?eSGUID=" & sSubGUID & "&action=back"
	                        Else
                                URLString = "PrePrompt.asp?subGUID=" & sSubGUID & "&qoid=" & sNextQOID
                                If Len(sNextPrefID) > 0 Then
                                    URLString = URLString & "&prefID=" & sNextPrefID
                                End If
                                Response.Redirect URLString
	                        End If
	                    End If
	                ElseIf StrComp(CStr(oRequest("action")), "finish", vbBinaryCompare) = 0 Then
	                    Response.Redirect "modify_subscription.asp?subGUID=" & sSubGUID
	                End If
	            End If
	        End If
	    End If
	End If
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