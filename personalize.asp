<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/PersonalizeCuLib.asp" -->
<!-- #include file="CustomLib/FoldersCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

    'Dim sGetFolderContentsXML
    Dim sGQAPFSSXML 'GetQuestionsAndProfilesForSubscriptionSetXML
    Dim sGetSubscriptionXML
    Dim sCacheXML

	Dim sServiceID
	Dim sServiceName
	Dim sSubGUID
	Dim sScheduleName
	Dim sFolderID
	Dim sPublicationID

	Dim sSubSetID

	Dim sAddressName
	Dim sStatusFlag
    Dim iNumQuestions
	Dim iNumVisibleQuestions
    Dim sSubsEnabled
    Dim sQOID
    Dim asPrefObj()
    Dim bFinishEnabled
    Dim sGetDetailsForQuestionsXML
	Dim sFirstVisibleQOID

	sSubscriptionsStyle = ""
	iSubscribeWizardStep = 3
	bFinishEnabled = False

	lErr = ParseRequestForPersonalize(oRequest, sSubGUID, sFolderID, sQOID)

	If lErr = NO_ERR Then
	    lErr = ReadCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
        If Len(sCacheXML) > 0 Then
            lErr = GetVariablesFromCache_Personalize(sCacheXML, sFolderID, sPublicationID, sServiceID, sServiceName, sAddressName, sScheduleName, sSubsEnabled, sStatusFlag, sSubSetID)
        End If
	End If

    'If lErr = NO_ERR Then
        'If Len(sFolderID) > 0 Then
        '    lErr = cu_GetFolderContents(sFolderID, sGetFolderContentsXML)
        'End If
    'End If

	If oRequest("persCancel").Count > 0 Then
	    lErr = DeleteCache(sSubGUID, CStr(GetSessionID()))
        If lErr = NO_ERR Then
            If Len(sFolderID) = 0 Then
                Response.Redirect "subscriptions.asp"
            Else
                Response.Redirect "services.asp?folderID=" & sFolderID
            End If
        End If
	ElseIf oRequest("persNext").Count > 0 Then
		Response.Redirect "PrePrompt.asp?qoid=" & sQOID & "&subGUID=" & sSubGUID
    ElseIf oRequest("persBack").Count > 0 Then
        Response.Redirect "subscribe.asp?eSGUID=" & sSubGUID & "&folderID=" & sFolderID & "&serviceID=" & sServiceID & "&serviceName=" & Server.URLEncode(sServiceName) & "&sf=" & sStatusFlag
	ElseIf oRequest("persFinish").Count > 0 Then
	    Response.Redirect "modify_subscription.asp?subGUID=" & sSubGUID
	End If

    If lErr = NO_ERR Then
        lErr = cu_GetQuestionsAndProfilesForSubscriptionSet(sSubSetID, sServiceID, sGQAPFSSXML)
        If lErr = NO_ERR Then
			lErr = CheckNumberOfQuestions(sGQAPFSSXML, iNumQuestions)
			If iNumQuestions > 0 Then
				lErr = AddQuestionsToCache_Personalize(sCacheXML, sGQAPFSSXML)
				If lErr = NO_ERR Then
					lErr = cu_GetDetailsForQuestions(sCacheXML, sGetDetailsForQuestionsXML)
					If lErr = NO_ERR Then
						lErr = AddQuestionDetailsToCache(sCacheXML, sGetDetailsForQuestionsXML)
						If lErr = NO_ERR Then		'visible QOs only
				            lErr = CheckNumberOfVisibleQuestions(sCacheXML, iNumVisibleQuestions, sFirstVisibleQOID)
						End If
					End If
				End If
			End If
		End If
    End If

   	If lErr = NO_ERR Then
        If iNumVisibleQuestions > 0 Then
            'If editing, load up previous answer
            If StrComp(sStatusFlag, "0", vbBinaryCompare) = 0 Then
                lErr = cu_GetSubscription(sSubGUID, sGetSubscriptionXML)
                If lErr = NO_ERR Then
                    lErr = AddPreviousAnswerToCache(sSubGUID, sCacheXML, sGetSubscriptionXML)
                End If
            End If
        End If
    End If

    If lErr = NO_ERR Then
        lErr = WriteCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
        If lErr = NO_ERR Then
            If iNumVisibleQuestions = 0 Then
                Response.Redirect "modify_subscription.asp?subGUID=" & sSubGUID
            End If
            lErr = IsFinishEnabled(sCacheXML, bFinishEnabled)
        End If
    End If

    If lErr = NO_ERR Then
		Select Case Clng(GetSummaryPageSetting())
		Case SITE_PROPVALUE_SUMMARY_PAGE_ALWAYS

		Case SITE_PROPVALUE_SUMMARY_PAGE_WHENMORETHANONEQO
			If iNumVisibleQuestions <= 1 Then
				If Strcomp(CStr(oRequest("action")),"back") = 0 Then
				    Response.Redirect "subscribe.asp?eSGUID=" & sSubGUID & "&folderID=" & sFolderID & "&serviceID=" & sServiceID & "&serviceName=" & Server.URLEncode(sServiceName) & "&sf=" & sStatusFlag
				Else
					Response.Redirect "PrePrompt.asp?qoid=" & sFirstVisibleQOID & "&subGUID=" & sSubGUID
				End If
			End If
		Case SITE_PROPVALUE_SUMMARY_PAGE_NEVER
			If Strcomp(CStr(oRequest("action")),"back") = 0 Then
			    Response.Redirect "subscribe.asp?eSGUID=" & sSubGUID & "&folderID=" & sFolderID & "&serviceID=" & sServiceID & "&serviceName=" & Server.URLEncode(sServiceName) & "&sf=" & sStatusFlag
			Else
				Response.Redirect "PrePrompt.asp?qoid=" & sFirstVisibleQOID & "&subGUID=" & sSubGUID
			End If
		End Select
	End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(354)'Descriptor: Subscriptions%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin left menu -->
		    <TABLE BORDER="0" CELLPADDING="3" CELLSPACING="0">
				<TR>
				    <TD>
				        <!-- #include file="_toolbar_Subscribe.asp" -->
				    </TD>
				</TR>
                <TR>
                    <TD><IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT=""></TD>
                </TR>
		    </TABLE>
			<!-- end left menu -->
		</TD>
		<TD WIDTH="1%">
			<IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="97%" VALIGN="TOP">
			<!-- begin center panel -->
			<%
			    If lErr <> NO_ERR Then
			        Response.Write "<BR />"
			    	Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(383), "services.asp") 'Descriptor: Back to Home
			    Else
			%>
			    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.write asDescriptors(540) 'Descriptor: Personalize your service content.%></B></FONT>
			    <% 'Call RenderPath_Personalize(sSubGUID, sServiceID, sServiceName, sFolderID, sGetFolderContentsXML, sStatusFlag)%>

				<TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0">
				    <TR>
				        <TD><IMG SRC="images/1ptrans.gif" WIDTH="20" HEIGHT="1" BORDER="0" ALT=""></TD>
				        <TD>
				            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
				                <%Response.Write asDescriptors(541) 'Descriptor: You can provide the following information to personalize the content for this subscription:%><BR />
				            </FONT>
				        </TD>
				    </TR>
				</TABLE>
				<BR />
	            <FORM ACTION="Personalize.asp" METHOD="POST">
				<INPUT TYPE="HIDDEN" NAME="eSGUID" VALUE="<%Response.Write sSubGUID%>" />
	            <INPUT TYPE="HIDDEN" NAME="folderID" VALUE="<%Response.Write sFolderID%>" />

                <%Call RenderPersonalize(sCacheXML, sSubGUID, sFolderID)%>

	            <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
                    <TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="2" WIDTH="1" ALT="" BORDER="0" /></TD></TR>
	                <TR><TD COLSPAN="2" BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" ALT="" BORDER="0" /></TD></TR>
	                <TR><TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" HEIGHT="3" WIDTH="1" ALT="" BORDER="0" /></TD></TR>
	                <TR>
	                    <TD><INPUT TYPE="SUBMIT" CLASS="buttonClass" NAME="persBack" VALUE="<%Response.Write asDescriptors(149) 'Descriptor: Back%>" /> <INPUT CLASS="buttonClass" TYPE="SUBMIT" NAME="persNext" VALUE="<%Response.Write asDescriptors(335) 'Descriptor: Next%>" /></TD>
	                    <TD ALIGN="RIGHT"><%If bFinishEnabled = True Then%><INPUT TYPE="SUBMIT" NAME="persFinish" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(442) 'Descriptor: Finish%>" /><%End If%> <INPUT CLASS="buttonClass" TYPE="SUBMIT" NAME="persCancel" VALUE="<%Response.Write asDescriptors(120) 'Descriptor: Cancel%>" /></TD>
	                </TR>
	            </TABLE>
                </FORM>
			<% End If %>
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