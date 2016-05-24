<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/SubscribeCuLib.asp" -->
<!-- #include file="CustomLib/FoldersCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

    Dim sGetUserAddressesForServiceXML
    Dim sGetNamedSchedulesForServiceXML
    Dim sGetFolderContentsXML
    Dim sCacheXML

    Dim bFlagValid
    Dim bHasAddresses
	Dim bHasSchedules

	Dim sServiceID
	Dim sFolderID
	Dim sServiceName

	Dim sAddressID
	Dim sAddressName
	Dim sPublicationID
	Dim sSubsSetID
	Dim sScheduleName
	Dim sStatusFlag
	Dim sESGUID 'Edit SubscriptionGUID
	Dim sEAID 'Edit AddressID
	Dim sESSID 'Edit SubscriptionSetID
	Dim sEPUBID ' Edit PublicationID
	Dim sSubID
	Dim sSelectedAddressID
	Dim sSelectedSubSetID
	Dim sEnabledFlag
	Dim sNewAddressValue
	Dim sNewSubsAddrID
	Dim sTransPropsID
	Dim sNewAddressTRPS
	Dim bFinishEnabled

	sSubscriptionsStyle = ""
	iSubscribeWizardStep = 2
	bFlagValid = True
	bHasAddresses = False
	bHasSchedules = False
	bFinishEnabled = False

	lErr = ParseRequestForSubscription(oRequest, sServiceID, sServiceName, sAddressID, sAddressName, sPublicationID, sSubsSetID, sScheduleName, sFolderID, sESGUID, sEAID, sESSID, sEPUBID, sStatusFlag, sEnabledFlag, sNewAddressValue, sTransPropsID, sSubID)

    If lErr = NO_ERR Then
        sSelectedAddressID = sEAID
	    sSelectedSubSetID = sESSID
        'If Len(sFolderID) > 0 Then
        '    lErr = cu_GetFolderContents(sFolderID, sGetFolderContentsXML)
        'End If
    End If

	If lErr = NO_ERR Then
	    lErr = ReadCache(sESGUID, CStr(GetSessionID()), sCacheXML)
        If Len(sCacheXML) > 0 Then
            lErr = GetVariablesFromCache_Subscribe(sCacheXML, sSelectedAddressID, sSelectedSubSetID, sEnabledFlag)
            If lErr = NO_ERR Then
                lErr = IsFinishEnabled(sCacheXML, bFinishEnabled)
            End If
        End If
	End If

    If lErr = NO_ERR Then
        If oRequest("subsSave").Count > 0 Then
            If StrComp(CStr(oRequest("ServAdd")), "n", vbBinaryCompare) = 0 Then
                Select Case CStr(Application("Device_Validation"))
                    Case S_DEVICE_VALIDATION_EMAIL
                        lValidationError = ValidateEmailAddress(sNewAddressValue)
                    Case S_DEVICE_VALIDATION_NUMBER
                        lValidationError = ValidateNumberAddress(sNewAddressValue)
                    Case S_DEVICE_VALIDATION_NONE
                        'Do Nothing
                    Case Else
                End Select

                If (Len(sNewAddressValue) = 0) Or (lValidationError = -1) Then
                    lValidationError = ERR_EMAIL_ADDR_INVALID
                Else
	                lErr = AddNewSubscriptionAddress(sNewAddressValue, sNewSubsAddrID, sNewAddressTRPS)
	                If lErr = NO_ERR Then
	                    sAddressName = sNewAddressValue
	                    sAddressID = sNewSubsAddrID
	                    sTransPropsID = sNewAddressTRPS
	                End If
                End If
	        End If

            If (lErr = NO_ERR) And (lValidationError = NO_ERR) Then
				If Len(sCacheXML) = 0 Then
        			If StrComp(sStatusFlag, "1", vbBinaryCompare) = 0 Then
					    sESGUID = GetGUID()
					End If

					lErr = GenerateCacheXML(sESGUID, sServiceID, sServiceName, sFolderID, sStatusFlag, sSubsSetID, sAddressID, sTransPropsID, sPublicationID, sAddressName, sScheduleName, sESSID, sEAID, sEPUBID, sEnabledFlag, sSubID, sCacheXML)
					If lErr = NO_ERR Then
						lErr = WriteCache(sESGUID, CStr(GetSessionID()), sCacheXML)
					End If
				Else	'modify selections
					lErr = UpdateCache_Subscribe(sSubsSetID, sAddressID, sTransPropsID, sPublicationID, sAddressName, sScheduleName, sEnabledFlag, sStatusFlag, sSubID, sCacheXML)
					If lErr = NO_ERR Then
						lErr = WriteCache(sESGUID, CStr(GetSessionID()), sCacheXML)
					End If
				End If

            End If

			If (lErr = NO_ERR) And (lValidationError = NO_ERR) Then
			    Response.Redirect "personalize.asp?eSGUID=" & sESGUID
			End If

        ElseIf oRequest("subsCancel").Count > 0 Then
	        If Len(sFolderID) = 0 Then
	            Response.Redirect "subscriptions.asp"
	        Else
	            Response.Redirect "services.asp?folderID=" & sFolderID
	        End If
        ElseIf oRequest("subsBack").Count > 0 Then
            'BACK currently does the same thing as cancel
	        If Len(sFolderID) = 0 Then
	            Response.Redirect "subscriptions.asp"
	        Else
	            Response.Redirect "services.asp?folderID=" & sFolderID
	        End If
        ElseIf oRequest("subsFinish").Count > 0 Then
            Response.Redirect "modify_subscription.asp?subGUID=" & sESGUID
        End If
    End If

    If lErr = NO_ERR Then
        lErr = cu_GetUserAddressesForService(sServiceID, sGetUserAddressesForServiceXML)
        If lErr = NO_ERR Then
            lErr = cu_GetNamedSchedulesForService(sServiceID, bFlagValid, sGetNamedSchedulesForServiceXML)
        End If

        If lErr = NO_ERR Then
            lErr = CheckNumberOfAddresses(sGetUserAddressesForServiceXML, bHasAddresses)
        End If
        If lErr = NO_ERR Then
            lErr = CheckNumberOfSchedules(sGetNamedSchedulesForServiceXML, bHasSchedules)
        End If
    End If

    If lErr = NO_ERR Then
        If Len(sEAID) > 0 Then
            lErr = GetOriginalPublication(sEAID, sGetUserAddressesForServiceXML, sEPUBID)
        End If
    End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE>MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<%If StrComp(GetJavaScriptSetting(), "1", vbBinaryCompare) = 0 Then%>

<SCRIPT LANGUAGE=javascript>
<!--
  function autoSelect(sValue) {
  var i;

    for (i = 0; i < document.subscribeForm.ServAdd.length; i++) {
      if (document.subscribeForm.ServAdd[i].value == sValue) {
        document.subscribeForm.ServAdd[i].click();
      }
    }

  }
//-->
</SCRIPT>

<%End If%>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
<!-- #include file="header_multi.asp" -->
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
	<TR>
		<TD WIDTH="1%" valign="TOP">
			<!-- begin left menu -->
		    <TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0>
				<TR>
				    <TD>
				        <!-- #include file="_toolbar_Subscribe.asp" -->
				    </TD>
				</TR>
                <TR>
                    <TD><img src="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT=""></TD>
                </TR>
		    </TABLE>
			<!-- end left menu -->
		</TD>
		<TD WIDTH=1%>
			<img src="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="97%" valign="TOP">
			<!-- begin center panel -->
			<% 'Call RenderPath_Subscribe(sServiceID, sServiceName, sFolderID, sGetFolderContentsXML)%>
			<!-- <BR /> -->
			<%
			    If lValidationError = ERR_EMAIL_ADDR_INVALID Then
			        Response.Write "<BR />"
	        	    sErrorHeader = asDescriptors(390) 'Descriptor: Error during subscription operation
                    Select Case CStr(Application("Device_Validation"))
                        Case S_DEVICE_VALIDATION_EMAIL
                            sErrorMessage = asDescriptors(419) 'Descriptor: Please enter an address in the form of: user@server.com
                        Case S_DEVICE_VALIDATION_NUMBER
                            sErrorMessage = asDescriptors(614) 'Descriptor: Please enter a value for the address in the following form: any numbers and the following characters - ( )
                        Case S_DEVICE_VALIDATION_NONE
                            sErrorMessage = asDescriptors(635) 'Descriptor: Please enter an address in the following form: any text or numeric characters
                        Case Else
                    End Select
			    	Call DisplayLoginError(sErrorHeader, sErrorMessage)
			    End If
			    If lErr <> NO_ERR Then
                    Response.Write "<BR />"
                    Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(383), "services.asp") 'Descriptor: Back to Services
			    Else
			%>
				<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
					<form NAME="subscribeForm" METHOD="POST" ACTION="subscribe.asp">
					<input type="HIDDEN" name="serviceID" value="<%Response.Write sServiceID%>" />
					<input type="HIDDEN" name="folderID" value="<%Response.Write sFolderID%>" />
					<input type="HIDDEN" name="serviceName" value="<%Response.Write Server.HTMLEncode(sServiceName)%>" />
					<input type="HIDDEN" name="sf" value="<%Response.Write sStatusFlag%>" />
					<input type="HIDDEN" name="eSGUID" value="<%Response.Write sESGUID%>" />
        			<input type="HIDDEN" name="eSubID" value="<%Response.Write sSubID%>" />
                    <input type="HIDDEN" name="eAID" value="<%Response.Write sEAID%>" />
                    <input type="HIDDEN" name="eSSID" value="<%Response.Write sESSID%>" />
                    <input type="HIDDEN" name="ePUBID" value="<%Response.Write sEPUBID%>" />

					<TR><TD><img src="images/1ptrans.gif" WIDTH="1" HEIGHT="10" BORDER="0" ALT=""></TD></TR>
					<TR>
					    <TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><b><%Response.Write asDescriptors(533) 'Descriptor: Select a schedule for this service to be delivered.%></b></font></TD>
					</TR>
					<TR>
						<TD VALIGN=TOP>
						    <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
						        <TR>
						            <TD><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="20" ALT="" BORDER="0" /></TD>
						            <TD>
							        <% Call RenderSchedulesForService(sSelectedSubSetID, sGetNamedSchedulesForServiceXML)%>
                                    </TD>
                                </TR>
                            </TABLE>
						</TD>
					</TR>
					<TR><TD><img src="images/1ptrans.gif" WIDTH="1" HEIGHT="25" BORDER="0" ALT=""></TD></TR>
					<TR>
					    <TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><b><%Response.Write asDescriptors(535) 'Descriptor: Select where you would like this service to be delivered.%></b></font></TD>
					</TR>
					<TR>
						<TD VALIGN=TOP>
						    <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
						        <TR>
						            <TD><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="20" ALT="" BORDER="0" /></TD>
						            <TD>
        							<% Call RenderAddressesForService(sSelectedAddressID, sGetUserAddressesForServiceXML)%>
                                    </TD>
                                </TR>
                            </TABLE>
						</TD>
					</TR>
					<TR><TD><img src="images/1ptrans.gif" WIDTH="1" HEIGHT="25" BORDER="0" ALT=""></TD></TR>
                    <TR>
                        <TD>
                            <INPUT TYPE="HIDDEN" NAME="enfCheck" VALUE="1" />
                            <INPUT TYPE="CHECKBOX" NAME="subsEnabled" <%If ((sEnabledFlag = "") OR (sEnabledFlag = "1")) Then%>CHECKED<%End If%> /> <font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(465) 'Descriptor: Enabled (subscription will be delivered on the specified schedule)%></font>
                        </TD>
                    </TR>
                    <TR><TD><img src="images/1ptrans.gif" WIDTH="1" HEIGHT="10" BORDER="0" ALT=""></TD></TR>
                </TABLE>
                <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
                    <TR>
                        <TD COLSPAN=2 BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" ALT="" BORDER="0" /></TD>
                    </TR>
                    <TR>
                        <TD COLSPAN=2><IMG SRC="images/1ptrans.gif" HEIGHT="3" WIDTH="1" ALT="" BORDER="0" /></TD>
                    </TR>
					<TR>
						<TD>
						    <INPUT TYPE=SUBMIT CLASS="buttonClass" NAME="subsBack" VALUE="<%Response.Write asDescriptors(149) 'Descriptor: Back%>" /> <%If (bHasAddresses = True) And (bHasSchedules = True) Then%><input type="SUBMIT" class="buttonClass" NAME="subsSave" VALUE="<%Response.Write asDescriptors(335) 'Descriptor: Next%>" /><%End If%>
						</TD>
						<TD ALIGN=RIGHT>
							<%If bFinishEnabled = True Then%><INPUT TYPE=SUBMIT NAME="subsFinish" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(442) 'Descriptor: Finish%>" /><%End If%> <input type="SUBMIT" class="buttonClass" NAME="subsCancel" VALUE="<%Response.Write asDescriptors(120) 'Descriptor: Cancel%>" />
						</TD>
					</TR>
					</form>
				</TABLE>
			<%End If%>
			<!-- end center panel -->
		</TD>
		<TD WIDTH=1%>
			<img src="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
	</TR>
</TABLE>
<BR>
<!-- begin footer -->
	<!-- #include file="footer.asp" -->
<!-- end footer -->
</BODY>
</HTML>