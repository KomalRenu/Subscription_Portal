<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/SubsConfirmCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sCacheXML
	Dim sSubGUID
	Dim sStatus
	Dim sFolderID
	Dim sServiceID
	Dim sServiceName
	Dim sScheduleName
	Dim sAddressID
	Dim sAddressName
	Dim sSubSetID
	Dim sPublicationID
	Dim sPersonalized

	Dim sURLString

	sSubscriptionsStyle = ""

	lErr = ParseRequestForSubsConfirm(oRequest, sSubGUID, sStatus)

	sFolderID = CStr(oRequest("folderID"))
	sServiceID = CStr(oRequest("serviceID"))
	sServiceName = CStr(oRequest("serviceName"))
	sScheduleName = CStr(oRequest("scheduleName"))
	sAddressID = CStr(oRequest("addressID"))
	sAddressName = CStr(oRequest("addressName"))
	sSubSetID = CStr(oRequest("subSetID"))
	sPublicationID = CStr(oRequest("pubID"))
	sPersonalized = CStr(oRequest("personalized"))

	If lErr = NO_ERR Then
	    lErr = ReadCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
	End If

	If lErr = NO_ERR Then
	    If oRequest("SubsConfirmOK").Count > 0 Then
            lErr = DeleteCache(sSubGUID, CStr(GetSessionID()))
			Response.Redirect "subscriptions.asp"
	    ElseIf oRequest("editSubs").Count > 0 Then
			Call ChangeStatusFlagToEdit(sCacheXML)
            lErr = WriteCache(sSubGUID, CStr(GetSessionID()), sCacheXML)
			Response.Redirect "subscribe.asp?serviceID=" & sServiceID & "&eSGUID=" & sSubGUID & "&eAID=" & sAddressID & "&eSSID=" & sSubSetID  & "&ePUBID=" & sPublicationID & "&serviceName=" & Server.URLEncode(sServiceName)
	    End If
	End If

	If lErr = NO_ERR Then
	        'If Len(sCacheXML) > 0 Then
	            lErr = ReadCacheVariables_SubsConfirm(sCacheXML, sFolderID, sServiceID, sServiceName, sScheduleName, sAddressID, sAddressName, sSubSetID, sPublicationID, sPersonalized)
	            'If lErr = NO_ERR Then
	            '    lErr = DeleteCache(sSubGUID, CStr(GetSessionID()))
	            'End If
	            'If lErr = NO_ERR Then
	            '   sURLString = ""
	            '    sURLString = sURLString & "SubsConfirm.asp?subGUID=" & sSubGUID & "&status=" & sStatus
	            '    sURLString = sURLString & "&folderID=" & sFolderID & "&serviceID=" & sServiceID
	            '    sURLString = sURLString & "&serviceName=" & Server.URLEncode(sServiceName) & "&scheduleName=" & Server.URLEncode(sScheduleName)
	            '    sURLString = sURLString & "&addressID=" & sAddressID & "&addressName=" & Server.URLEncode(sAddressName)
	            '    sURLString = sURLString & "&subSetID=" & sSubSetID & "&personalized=" & sPersonalized
	            '    Response.Redirect sURLString
	            'End If
	        'Else
	        'End If
	    'End If
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
			<!-- begin left menu -->
            <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
            	<TR>
            		<TD>
            			<IMG SRC="images/1ptrans.gif" WIDTH="2" HEIGHT="1" BORDER="0" ALT="">
            		</TD>
            		<TD VALIGN="TOP">
            			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
            				<TR>
            					<TD>
            						<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(773) 'Descriptor: Related links%><%Else%><%Response.Write UCase(asDescriptors(773)) 'Descriptor: Related links%><%End If%></B></FONT>
            					</TD>
            				</TR>
            				<TR>
            				    <TD>
            				        <TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0">
            				            <TR>
            				                <TD VALIGN="TOP"><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
            				                <TD><A HREF="subscribe.asp?serviceID=<%Response.Write sServiceID%>&serviceName=<%Response.Write Server.URLEncode(sServiceName)%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(774) 'Descriptor: Sign up for this service again%></FONT></A></TD>
            				            </TR>
            				            <TR>
            				                <TD VALIGN="TOP"><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
            				                <TD><A HREF="services.asp"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(775) 'Descriptor: Sign up for another service%></FONT></A></TD>
            				            </TR>
            				            <TR>
            				                <TD VALIGN="TOP"><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
            				                <TD><A HREF="subscriptions.asp"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(776) 'Descriptor: View all your subscriptions%></FONT></A></TD>
            				            </TR>
            				        </TABLE>
            				    </TD>
            				</TR>
            			</TABLE>
            		</TD>
            	</TR>
            </TABLE>
			<BR />
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="1%">
			<IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="97%" VALIGN="TOP">
			<!-- begin center panel -->
			<%
			    If lErr <> NO_ERR Then
			    	Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(383), "services.asp") 'Descriptor: Back to Services
			    Else
			%>
				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="500">
                    <FORM METHOD="POST" ACTION="SubsConfirm.asp">
                    <INPUT TYPE="HIDDEN" NAME="subGUID" VALUE="<%Response.Write sSubGUID%>" />
                    <INPUT TYPE="HIDDEN" NAME="folderID" VALUE="<%Response.Write sFolderID%>" />
                    <INPUT TYPE="HIDDEN" NAME="serviceID" VALUE="<%Response.Write sServiceID%>" />
                    <INPUT TYPE="HIDDEN" NAME="serviceName" VALUE="<%Response.Write Server.HTMLEncode(sServiceName)%>" />
                    <INPUT TYPE="HIDDEN" NAME="scheduleName" VALUE="<%Response.Write Server.HTMLEncode(sScheduleName)%>" />
                    <INPUT TYPE="HIDDEN" NAME="addressID" VALUE="<%Response.Write sAddressID%>" />
                    <INPUT TYPE="HIDDEN" NAME="addressName" VALUE="<%Response.Write Server.HTMLEncode(sAddressName)%>" />
                    <INPUT TYPE="HIDDEN" NAME="subSetID" VALUE="<%Response.Write sSubSetID%>" />
                    <INPUT TYPE="HIDDEN" NAME="pubID" VALUE="<%Response.Write sPublicationID%>" />
                    <INPUT TYPE="HIDDEN" NAME="personalized" VALUE="<%Response.Write sPersonalized%>" />

                    <%If StrComp(sStatus, "success", vbBinaryCompare) = 0 Or (Len(sStatus) = 0) Then%>
                    <TR>
                        <TD>
                            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>">
                                <FONT SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(544) 'Descriptor: Success!%></B></FONT><BR />
                                <FONT SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(545) 'Descriptor: Your subscription has been saved successfully.%></FONT>
                            </FONT>
                        </TD>
                    </TR>
                    <%End If%>
				    <TR>
				        <TD><IMG SRC="images/1ptrans.gif" HEIGHT="10" WIDTH="1" ALT="" BORDER="0" /></TD>
				    </TR>
				    <TR>
				        <TD>
				            <TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0">
				                <TR>
				                    <TD><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="40" ALT="" BORDER="0" /></TD>
				                    <TD BGCOLOR="#000000">
				                        <TABLE BORDER="0" CELLPADDING="3" CELLSPACING="0">
				                            <TR>
				                                <TD BGCOLOR="#ffffff">
				                                    <TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0" WIDTH="300">
				                                        <TR>
				                                            <TD>
				                                                <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
				                                                <%Response.Write asDescriptors(366) & ":" 'Descriptor: Service%><BR />
				                                                <B><%Response.Write sServiceName%></B><BR/><BR/>
				                                                </FONT>
				                                            </TD>
				                                        </TR>
				                                        <TR>
				                                            <TD>
				                                                <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
				                                                <%Response.Write asDescriptors(351) & ":" 'Descriptor: Schedule%><BR />
				                                                <B><%Response.Write sScheduleName%></B><BR/><BR />
				                                                </FONT>
				                                            </TD>
				                                        </TR>
				                                        <TR>
				                                            <TD>
				                                                <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
				                                                <%Response.Write asDescriptors(451) 'Descriptor: Send to:%><BR />
				                                                <B><%Response.Write sAddressName%></B><BR /><BR />
				                                                </FONT>
				                                            </TD>
				                                        </TR>
				                                        <TR>
				                                            <TD>
				                                                <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
				                                                <%Response.Write asDescriptors(352) & ":" 'Descriptor: Personalized%><BR />
				                                                <B><%Response.Write sPersonalized%></B>
				                                                </FONT>
				                                            </TD>
				                                        </TR>
				                                    </TABLE>
				                                </TD>
				                            </TR>
				                        </TABLE>
				                    </TD>
				                </TR>
				            </TABLE>
				        </TD>
				    </TR>
				    <TR>
				        <TD><IMG SRC="images/1ptrans.gif" HEIGHT="10" WIDTH="1" ALT="" BORDER="0" /></TD>
				    </TR>
				    <TR>
				        <TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" ALT="" BORDER="0" /></TD>
				    </TR>
				    <TR>
				        <TD><IMG SRC="images/1ptrans.gif" HEIGHT="3" WIDTH="1" ALT="" BORDER="0" /></TD>
				    </TR>
				    <TR>
				        <TD ALIGN="RIGHT">
				            <INPUT TYPE="SUBMIT" CLASS="buttonClass" NAME="SubsConfirmOK" VALUE="   <%Response.Write asDescriptors(543) 'Descriptor: OK%>   " />
				            <INPUT TYPE="SUBMIT" CLASS="buttonClass" NAME="editSubs" VALUE="<%Response.Write asDescriptors(772) 'Descriptor: Change selections%>" />
				        </TD>
				    </TR>
			        </FORM>
				</TABLE>
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