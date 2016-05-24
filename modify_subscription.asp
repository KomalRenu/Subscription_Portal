<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/ModifySubscriptionCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sCacheXML
	Dim sSubGUID
	Dim iNumQuestions
	Dim sStatusFlag
	Dim sSubsSetID
	Dim sServiceID
	Dim sAddressID
	Dim sFolderID
	Dim sSubsEnabled
	Dim sPublicationID
	Dim sTransPropsID
	Dim sOriginalPublicationID
	Dim sOriginalSubsSetID
	Dim sSubID

	sSubscriptionsStyle = ""
	iNumQuestions = 0

	lErr = ParseRequestForModifySubscription(oRequest, sSubGUID, sStatusFlag)

	If oRequest("subsCancel").Count > 0 Then
	    Response.Redirect "subscriptions.asp?serviceID=" & oRequest("serviceID") & "&folderID=" & oRequest("folderID")
	End If

    If lErr = NO_ERR Then
        If StrComp(sStatusFlag, "2", vbBinaryCompare) <> 0 Then
            lErr = ReadCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
            If lErr = NO_ERR Then
                lErr = ReadSubscriptionProperties(sCacheXML, iNumQuestions, sStatusFlag, sSubsSetID, sServiceID, sAddressID, sFolderID, sSubsEnabled, sPublicationID, sOriginalPublicationID, sOriginalSubsSetID, sTransPropsID, sSubID)
            End If
        End If
    End If

    If lErr = NO_ERR Then
        If StrComp(sStatusFlag, "1", vbBinaryCompare) = 0 Then 'New subscription
            lErr = cu_AddSubscription(sSubsSetID, sServiceID, sAddressID, sTransPropsID, sSubGUID, sSubsEnabled, sCacheXML)
            If lErr = NO_ERR Then
				lErr = WriteCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
			End If
			If lErr = NO_ERR Then
				Response.Redirect "SubsConfirm.asp?subGUID=" & sSubGUID & "&status=success"
			End If
        ElseIf StrComp(sStatusFlag, "0", vbBinaryCompare) = 0 Then 'Edit subscription
            lErr = cu_EditSubscription(sSubsSetID, sServiceID, sAddressID, sTransPropsID, sSubGUID, sSubsEnabled, sSubID, sCacheXML)
            If lErr = NO_ERR Then
				lErr = WriteCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
			End If
			If lErr = NO_ERR Then
                Response.Redirect "SubsConfirm.asp?subGUID=" & sSubGUID & "&status=success"
            End If
        ElseIf StrComp(sStatusFlag, "2", vbBinaryCompare) = 0 Then 'Delete subscription
            lErr = cu_DeleteSubscriptions()
            If lErr = NO_ERR Then
                Response.Redirect "subscriptions.asp?serviceID=" & oRequest("serviceID") & "&folderID=" & oRequest("folderID")
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
			<BR />
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