<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/OptionsCuLib.asp" -->
<!-- #include file="CustomLib/LoginCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sSiteLanguage
	Dim sUseJavaScript
	Dim sStartPage
	Dim sOptSection
	Dim sLocale
	Dim bSaved
	Dim sHeader
	Dim sGetLocalesForSiteXML
	Dim sSummaryPage
	Dim sPortalDeviceID

	sOptionsStyle = ""
	bSaved = False
	sPortalDeviceID = Application.Value("Portal_device")

	lErr = ParseRequestForOptions(oRequest, sSiteLanguage, sOptSection, sUseJavaScript, sStartPage, sLocale, sSummaryPage)

	'Check if a change was posted to this page
	If lErr = NO_ERR Then
		If oRequest("saveOpt").count > 0 Then
			lErr = ChangeLanguage(sSiteLanguage, asDescriptors, aFontInfo)
			If lErr = NO_ERR Then
				lErr = ChangeJavaScript(sUseJavaScript)
			End If
			If lErr = NO_ERR Then
			    lErr = ChangeStartPage(sStartPage)
			End If
			If lErr = NO_ERR Then
			    lErr = ChangeLocale(sLocale)
			End If
			If lErr = NO_ERR Then
			    lErr = ChangeSummaryPage(sSummaryPage)
			End If
			If lErr = NO_ERR Then
			    bSaved = True
			End If
		End If
	End If

	If lErr = NO_ERR Then
	    lErr = cu_GetLocalesForSiteByCurrentLocale(sGetLocalesForSiteXML)
	End If

	Select Case sOptSection
		Case "1"
			sHeader = asDescriptors(402) 'Descriptor: User options
        Case Else
            sHeader = asDescriptors(402) 'Descriptor: User options
	End Select
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(365)'Descriptor: Options%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY BGCOLOR="ffffff" TOPMARGIN="0" LEFTMARGIN="0" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin left menu -->
				<!-- #include file="toolbarUserOptions.asp" -->
			<BR />
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" VALIGN="TOP">
			<!-- begin center panel -->
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
			    <TR>
			        <TD VALIGN="CENTER">
			            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(26) & " " 'Descriptor: You are here:%> <%Response.Write asDescriptors(286) 'Descriptor: Preferences%> > <B><%Response.Write sHeader%></B></FONT>
			        </TD>
			        <TD ALIGN="RIGHT"><IMG SRC="images/desktop_preferences.gif" WIDTH="60" HEIGHT="60" BORDER="0" ALT="" /></TD>
			    </TR>
			</TABLE>
		<%
		    If lErr <> NO_ERR Then
			    Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(380), "home.asp") 'Descriptor: Back to Home
		    Else
		%>
				<% If bSaved = True Then %>
					<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
						<TR>
							<TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#cc0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(403) 'Descriptor: Your options have been saved.%></B></FONT></TD>
						</TR>
						<TR>
							<TD><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
						</TR>
						<TR>
							<TD BGCOLOR="#000000"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
						</TR>
					</TABLE>
					<BR />
				<% End If %>

            <!-- BEGIN: Web 7 Preferences -->
			<TABLE BGCOLOR="#CCCCCC" BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
    		<FORM METHOD="GET" ACTION="options.asp">
			<INPUT TYPE="HIDDEN" NAME="optSection" VALUE="<%Response.Write sOptSection%>">
			<% If StrComp(sOptSection, "1", vbBinaryCompare) = 0 Or Len(sOptSection) = 0 Then 'QUESTION: Should this logic include other pages, or just have everything here? %>
				    <TR>
					    <TD ALIGN="LEFT" VALIGN="TOP"><IMG SRC="Images/loginUpperLeftCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
						<TD WIDTH="100%"><FONT SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(402) 'Descriptor: User options%></B></FONT></TD>
						<TD ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="Images/loginUpperRightCorner.gif" WIDTH="11" HEIGHT="11" ALT="" BORDER="0" /></TD>
                    </TR>
				</TABLE>
				<TABLE BGCOLOR="#CCCCCC" BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="1">
				    <TR><TD COLSPAN="3">
					        <TABLE BGCOLOR="#FFFFFF" BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="3"><TR>
							    <TD WIDTH="1%">&nbsp;&nbsp;</TD>
								<TD><BR />
                                    <TABLE WIDTH="98%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
                                        <!-- BEGIN: Start page Options -->
                                        <TR>
		                                    <TD VALIGN="TOP">
			                                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(138) 'Descriptor: Default start page:%></B>&nbsp;</FONT>
		                                    </TD>
		                                    <TD>
		                                        <SELECT CLASS="pullDownClass" NAME="startPage">
		                                            <OPTION VALUE="<%Response.Write S_PAGE_HOME%>" <%If StrComp(GetStartPage(), S_PAGE_HOME, vbBinaryCompare) = 0 Then%>SELECTED<%End If%>><%Response.Write asDescriptors(1) 'Descriptor: Home%></OPTION>
		                                            <OPTION VALUE="<%Response.Write S_PAGE_SUBSCRIPTIONS%>" <%If StrComp(GetStartPage(), S_PAGE_SUBSCRIPTIONS, vbBinaryCompare) = 0 Then%>SELECTED<%End If%>><%Response.Write asDescriptors(354) 'Descriptor: Subscriptions%></OPTION>
		                                            <%If Len(sPortalDeviceID) > 0 Then %><OPTION VALUE="<%Response.Write S_PAGE_REPORTS%>" <%If StrComp(GetStartPage(), S_PAGE_REPORTS, vbBinaryCompare) = 0 Then%>SELECTED<%End If%>><%Response.Write asDescriptors(360) 'Descriptor: Reports%></OPTION> <%End If%>
		                                            <OPTION VALUE="<%Response.Write S_PAGE_ADDRESSES%>" <%If StrComp(GetStartPage(), S_PAGE_ADDRESSES, vbBinaryCompare) = 0 Then%>SELECTED<%End If%>><%Response.Write asDescriptors(361) 'Descriptor: Addresses%></OPTION>
		                                            <OPTION VALUE="<%Response.Write S_PAGE_SERVICES%>" <%If StrComp(GetStartPage(), S_PAGE_SERVICES, vbBinaryCompare) = 0 Then%>SELECTED<%End If%>><%Response.Write asDescriptors(362) 'Descriptor: Services%></OPTION>
		                                        </SELECT>
                                                <BR />
		                                    </TD>
	                                    </TR>
	                                    <!-- END: Start page Options -->
	                                    <TR><TD COLSPAN="2"><HR SIZE="1" /></TD></TR>
	                                    <!-- BEGIN: Locale Options -->
                                        <TR>
		                                    <TD VALIGN="TOP">
		                                    	<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(404) 'Descriptor: Locale:%></B></FONT>
		                                    </TD>
		                                    <TD>
		                                        <% Call RenderLocaleChoices(sGetLocalesForSiteXML, "", "", GetSiteLocale())  %>
	                                    	</TD>
	                                    </TR>
	                                    <!-- END: Locale Options -->
		                                <TR><TD COLSPAN="2"><HR SIZE="1" /></TD></TR>
		                                <!-- BEGIN: JavaScript Options -->
		                                <TR>
		                                	<TD VALIGN="TOP">
		                                		<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write "Dynamic HTML:"%></B></FONT>
		                                	</TD>
		                                	<TD VALIGN="TOP">
		                                		<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>">
		                                			<%Response.Write asDescriptors(70) 'Descriptor: Use Dynamic HTML?%>
		                                		</FONT>
                                                <SELECT CLASS="pullDownClass" NAME="useJavaScript">
                                                    <OPTION VALUE="2" <%If StrComp(GetJavaScriptPreference(), "2", vbBinaryCompare) = 0 Then%>SELECTED<%End If%>><%Response.Write asDescriptors(304) 'Descriptor: Determine automatically%></OPTION>
                                                    <OPTION VALUE="1" <%If StrComp(GetJavaScriptPreference(), "1", vbBinaryCompare) = 0 Then%>SELECTED<%End If%>><%Response.Write asDescriptors(119) 'Descriptor: Yes%></OPTION>
                                                    <OPTION VALUE="0" <%If StrComp(GetJavaScriptPreference(), "0", vbBinaryCompare) = 0 Then%>SELECTED<%End If%>><%Response.Write asDescriptors(118) 'Descriptor: No%></OPTION>
                                                </SELECT>
		                                		<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
		                                			<BR />
		                                			<%Response.write asDescriptors(236) 'Descriptor: Note: Netscape 4, Internet Explorer 4 and newer versions of these browsers support DHTML.%>
		                                			<BR />
		                                		</FONT>
		                                	</TD>
		                                </TR>
		                                <!-- END: JavaScript Options -->

		                                <!-- BEGIN: Summary Page Options -->
                                        <!--
											<TR>
		                                    <TD VALIGN="TOP">
		                                    	<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B>
		                                    	<%Response.Write(asDescriptors(0)) 'Descriptor: Display summary page for personalized service:%></B></FONT>
		                                    </TD>
		                                    <TD>
		                                        <% Call RenderSummaryPageChoices() %>
	                                    	</TD>
	                                    </TR>
	                                    -->
	                                    <!-- END: Summary Page Options -->

                                    </TABLE><BR />
								</TD>
                            </TR></TABLE>
					</TD></TR>
					<TR><TD COLSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
					<TR>
						<TD COLSPAN="2">
							<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
								<TR>
								    <TD>&nbsp;&nbsp;</TD>
									<TD>
									    <INPUT TYPE="SUBMIT" CLASS="buttonClass" NAME="saveOpt" VALUE="<%Response.Write asDescriptors(74) 'Descriptor: Apply%>" />
									</TD>
								</TR>
							</TABLE>
                        </TD>
						<TD>&nbsp;</TD>
					</TR>
					<TR><TD COLSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
			<% End If %>
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