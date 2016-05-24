<%@LANGUAGE=VBSCRIPT%>
<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Option Explicit
On Error Resume Next

Dim oBinaryRequest
Dim iRequestSize

Set oRequest = Nothing

If Request.Form.Count > 0 Then
	Set oRequest = Request.Form
Else
	Set oRequest = Request.QueryString
End If

If IsEmpty(oRequest("xml")) Then
	Response.CacheControl = "no-cache"		'enforce go back to reexecute the page
	Response.AddHeader "Pragma", "no-cache"
End If
Response.Expires = -1
%>

<!-- #include file="commonDeclarations.asp" -->
<!-- #include file="CustomLib/PromptCuLib.asp" -->

<%
	Dim aObjectInfo()
	Dim aFolderInfo()
	Dim aPromptGeneralInfo()
	Dim aPromptInfo
	'Dim sStartPageURL
	'Dim aScheduleInfo()
	'Dim aSubscriptionInfo()

    'Wizard Left Panel
    Dim sCacheXML
    Dim sServiceName
    Dim sScheduleName
    Dim sAddressName
	Dim sMessage

	sMessage = ""

	iSubscribeWizardStep = 3

	sSubscriptionsStyle = ""

	Erase asDescriptors
	asDescriptors = asWebDescriptors

	lErrNumber = ReceivePromptRequest(oRequest, aObjectInfo, aFolderInfo, aPromptGeneralInfo)

	Erase asDescriptors
	asDescriptors = asHydraDescriptors

	If lErrNumber <> NO_ERR Then
		lErr = lErrNumber
	Else
		Erase asDescriptors
		asDescriptors = asWebDescriptors
		lErrNumber = GetPrompt(oRequest, aConnectionInfo, oSession, oObjServer, aObjectInfo, aFolderInfo, aPromptGeneralInfo, sErrDescription)
		Erase asDescriptors
		asDescriptors = asHydraDescriptors

		If lErrNumber <> NO_ERR Then
			lErr = lErrNumber
			If lErrNumber = ERR_UNSUPPORTED_PROMPTS Then
				sMessage = asDescriptors(941) 'This service contains an unsupported prompt type.
			End If
		Else
			sCacheXML = aPromptGeneralInfo(PROMPT_O_HYDRAPROMPTS).xml
			Call IsFinishEnabled(sCacheXML, aPromptGeneralInfo(PROMPT_B_FINISH_ENABLED))
		End If
	End If



%>
<!-- #include file="CheckError.asp" -->

<HTML>
	<HEAD>
		<%Response.Write(putMETATagWithCharSet())%>
		<TITLE>MicroStrategy Narrowcast Server<%'Response.Write aObjectInfo(S_NAME_OBJECT)%></TITLE>
		<%
		Erase asDescriptors
		asDescriptors = asWebDescriptors

		If aPromptGeneralInfo(PROMPT_B_DHTML) Then%>

		<SCRIPT language="JavaScript"><!--
			var aDescriptor = new Array();

			aDescriptor[1] = "<%Response.Write asDescriptors(696)%>"; //Descriptor: Between
			aDescriptor[2] = "<%Response.Write asDescriptors(746)%>"; //Descriptor: Not between
			aDescriptor[3] = "<%Response.Write asDescriptors(525)%>"; //Descriptor: Like
			aDescriptor[4] = "<%Response.Write asDescriptors(526)%>"; //Descriptor: Not Like
			aDescriptor[5] = "<%Response.Write asDescriptors(529)%>"; //Descriptor: Highest
			aDescriptor[6] = "<%Response.Write asDescriptors(530)%>"; //Descriptor: Lowest
			aDescriptor[7] = "<%Response.Write asDescriptors(587)%>"; //Descriptor: In
			aDescriptor[8] = "<%Response.Write asDescriptors(701)%>"; //Descriptor: and
			aDescriptor[9] = "<%Response.Write Replace(asDescriptors(ERR_BETWEEN_EXPECTS_TWOVALUES), """", "\""")%>"; //Descriptor: The correct syntax for the Between operator is value1;value2
			aDescriptor[10] = "<%Response.Write Replace(asDescriptors(ERR_NOTBETWEEN_EXPECTS_TWOVALUES), """", "\""")%>"; //Descriptor: The correct syntax for the Not Between operator is value1;value2
			aDescriptor[11] = "<%Response.Write Replace(asDescriptors(ERR_IN_EXPECTS_VALUE), """", "\""")%>"; //Descriptor: Please type valid values separated by semicolon(s).
			aDescriptor[12] = "<%Response.Write Replace(asDescriptors(ERR_OPERATOR_EXPECTS_VALUE), """", "\""")%>"; //Descriptor: Please enter a value in the text box.
			aDescriptor[13] = "<%Response.Write Replace(asDescriptors(ERR_NOT_QUALIFY_TWO), """", "\""")%>"; //Descriptor: You can not qualify on more than one item at the same time.
			aDescriptor[14] = "<%Response.Write Replace(asDescriptors(ERR_NOT_SUPPORT_DATA_TYPE_NONTEXT), """", "\""")%>"; //Descriptor: The operators "Like" and "Not Like" do not support non-text data type.
			aDescriptor[15] = "<%Response.Write Replace(asDescriptors(ERR_TOOMANY_SELECTIONS_EXPRESSIONPROMPT), """", "\""")%>"; //Descriptor: You have selected too many qualifications for an expression prompt.
			aDescriptor[16] = "<%Response.Write Replace(asDescriptors(980), """", "\""")%>"; //Descriptor: Follow the instructions marked below by a red flag.
			aDescriptor[17] = "<%Response.Write Replace(asDescriptors(1261), "##", ReadUserOption(ALLOWED_FILE_EXTENSION_OPTION))%>"; //Descriptor: The file contains an invalid extension. The file extensions allowed are: ##
			aDescriptor[18] = "<%Response.Write asDescriptors(181)%>"; //Descriptor: - none -
			aDescriptor[19] = "<%Response.Write asDescriptors(267)%>"; //Descriptor: (default)
			aDescriptor[20] = "<%Response.Write asDescriptors(2204)%>"; //Descriptor: Not In

			var bDefault = new Array();
			var sValidExtensions = ',' + '<%Response.Write Replace(Replace(ReadUserOption(ALLOWED_FILE_EXTENSION_OPTION), " ", ""), ".", "")%>' + ',';
			var sFontFamily = '<%Response.Write aFontInfo(S_FAMILY_FONT)%>';
			var sSmallFontSize = '<%Response.Write aFontInfo(N_SMALL_FONT)%>';
			<%call SetDefaultArrayforJS(aPromptGeneralInfo, aPromptInfo)%>
		//--></SCRIPT>
		<SCRIPT language="JavaScript" SRC="PromptFunctions.js"><!-- //--></SCRIPT>
		<SCRIPT language="JavaScript" SRC="Calendar.js"><!-- //--></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="URLManagement.js"><!-- // --></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="DHTMLapi.js"><!-- //--></SCRIPT>

		<%
			Call PopulateClientDescriptors(asDescriptors, True)
		Else
			Call DisplayGoToAnchorDHTML(True)
		End If%>

	</HEAD>
	<BODY BGCOLOR="#FFFFFF" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0"<%
	If Len(aPromptGeneralInfo(PROMPT_S_CURORDER)) > 0 And aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) And aPromptGeneralInfo(PROMPT_B_DHTML) Then
		Response.Write " onLoad=""gotoAnchor(" & CStr(aPromptGeneralInfo(PROMPT_S_CURORDER)) & ")"""
	End If%>>
	<%
	Erase asDescriptors
	asDescriptors = asHydraDescriptors
	%>
		<!-- #include file="header_multi.asp" -->
		<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
			<TR>
				<TD NOWRAP WIDTH="1%" valign="TOP">
				    <!-- include file="_toolbar_prompt.asp" -->
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
				<TD WIDTH="1%">
					<img src="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
				</TD>

				<TD VALIGN="TOP" WIDTH="97%">

				<%'Call DisplayObjectPath(aConnectionInfo, aObjectInfo)

				If lErr <> NO_ERR Then
					Response.Write "<BR /><BR />"
			        Call DisplayError(sErrorHeader, sErrorMessage & " " & sMessage, asDescriptors(383), "services.asp") 'Descriptor: Back to Services
					Response.End
			    End If

				If aPromptGeneralInfo(PROMPT_B_SPECIAL_FORM) Then%>
					<FORM ACTION="ProcessingPrompt.asp" METHOD="POST" NAME="PromptForm" ENCTYPE="MULTIPART/FORM-DATA" onSubmit="return(BuildUserSelections(<%Response.Write aPromptGeneralInfo(PROMPT_L_MAXPIN)%>));">
				<%Elseif aPromptGeneralInfo(PROMPT_B_DHTML) Then %>
					<FORM ACTION="prompt.asp" METHOD="POST" NAME="PromptForm" onSubmit="return(BuildUserSelections(<%Response.Write aPromptGeneralInfo(PROMPT_L_MAXPIN)%>));">
				<%Else%>
					<FORM ACTION="prompt.asp" METHOD="POST" NAME="PromptForm">
				<%End If%>

					<!-- General Error -->
					<%If aPromptGeneralInfo(PROMPT_B_DHTML) Then%>
						<DIV Name="GeneralErrorDisplay" ID="GeneralErrorDisplay"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></DIV>
					<%End If

					If Len(sErrDescription) > 0 Then
						Call WritePromptGeneralError(sErrDescription)
					ElseIf lErrNumber <> NO_ERR Or aPromptGeneralInfo(PROMPT_B_ANYERROR) Then
						Call WritePromptGeneralError(asDescriptors(321)) 'Descriptor: Follow the instructions marked below by a red flag.
					End If

					If aPromptGeneralInfo(PROMPT_B_ALLOW_PROFILE) Then
						Call DisplayProfileList(aConnectionInfo, aPromptGeneralInfo)
					End If
					%>

					<TABLE BORDER="0" COLS="2" WIDTH="100%">
					<TR>
						<TD WIDTH="1%"><IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" ALT="" BORDER="0" /></TD>
						<TD>

						<!-- BEGIN: new prompt widget -->
						<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
							<!-- BEGIN:  REPORT TITLE BAR -->
							<TR>
								<TD WIDTH="100%"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>" COLOR="#CC0000"><B><%Response.Write aObjectInfo(S_NAME_OBJECT)%></B></FONT></TD>
								<!--
								<%If aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) >= 3 And aPromptGeneralInfo(PROMPT_B_ALLPROMPTSINONEPAGE) And Not aPromptGeneralInfo(PROMPT_B_SUMMARY) Then%>
								<TD><INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON" NAME="PromptGO" VALUE="<%
									If aPromptGeneralInfo(PROMPT_B_ISDOC) Then
										Response.write asDescriptors(303) 'Descriptor: Execute Document
									Else
										Response.write asDescriptors(215) 'Descriptor: Execute Report
									End If
								%>" /></TD>
								<TD><IMG SRC="images/1ptrans.gif" WIDTH="3" HEIGHT="1" BORDER="0" ALT="" /></TD>
									<%If aPageInfo(N_ALIAS_PAGE) = DssXmlFolderNameTemplateReports Then 'hide save button for users don't have either save or publish privilege
										If (aConnectionInfo(N_PRIVILEGES_CONNECTION) And PRIVILEGE_WEBSAVEREPORT) Or (aConnectionInfo(N_PRIVILEGES_CONNECTION) And PRIVILEGE_WEBPUBLISH) Then%>
											<TD><INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON" NAME="SaveAsBtn" VALUE="<%Response.Write asDescriptors(216) 'Descriptor: Save Report%>" /></TD>
											<TD><IMG SRC="images/1ptrans.gif" WIDTH="3" HEIGHT="1" BORDER="0" ALT="" /></TD>
										<%End If
									End If%>
								<TD><INPUT TYPE="SUBMIT" CLASS="REDBUTTON" NAME="cancel" VALUE="<%Response.Write asDescriptors(120) 'Descriptor: Cancel%>" /></TD>
								<%End If%> -->
							</TR>
							<TR><TD COLSPAN="5"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT="" /></TD></TR>
							<!-- END:  REPORT TITLE BAR -->
						</TABLE>

						<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%">
							<TR>
								<!-- BEGIN: Prompt Index -->
								<%If aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) >= 2 Then%>
								<TD BGCOLOR="#CCCC99" WIDTH="100" VALIGN="TOP" ROWSPAN="3">
									<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
										<TR>
											<TD ALIGN="LEFT" VALIGN="TOP"><IMG SRC="Images/loginUpperLeftCorner.gif" WIDTH="11" HEIGHT="11" /></TD>
											<TD><IMG SRC="Images/1ptrans.gif" WIDTH="89" HEIGHT="1" /></TD>
										</TR>
										<TR>
											<TD ALIGN="LEFT" COLSPAN="2">
												<%

												Erase asDescriptors
												asDescriptors = asWebDescriptors
												Call DisplayPromptIndex(aPromptGeneralInfo, aPromptInfo)
												Erase asDescriptors
												asDescriptors = asHydraDescriptors
												%>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<%Else%>
								<TD BGCOLOR="#CCCC99" WIDTH="1" ROWSPAN="4"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
								<%End If%>
								<!-- END: Prompt Index -->

								<TD WIDTH="100%">
									<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
										<TR><TD BGCOLOR="#CCCC99"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
										<TR><TD><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="5" ALT="" BORDER="0" /></TD></TR>
									</TABLE>
								</TD>
								<%If aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) >= 2 Then%>
								<TD BGCOLOR="#CCCC99" WIDTH="1" ROWSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
								<%Else%>
								<TD BGCOLOR="#CCCC99" ROWSPAN="4"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
								<%End If%>
							</TR>

							<!-- BEGIN: Each Prompt -->
							<%If Not aPromptGeneralInfo(PROMPT_B_SUMMARY) Then%>
							<TR>
								<TD VALIGN="TOP">
								<%Erase asDescriptors
								asDescriptors = asWebDescriptors
								Call DisplayAllPrompts(aConnectionInfo, aPromptInfo, aPromptGeneralInfo, sErrDescription)
								Erase asDescriptors
								asDescriptors = asHydraDescriptors

								%>
								<BR />
								</TD>
							</TR>
							<%Else %>
							<TR>
								<TD VALIGN="TOP" COLSPAN="2">
								<!-- BEGIN:  Summary -->

								<%
								Erase asDescriptors
								asDescriptors = asWebDescriptors
								Call DisplayPromptSummary(aPromptGeneralInfo, aPromptInfo)
								Erase asDescriptors
								asDescriptors = asHydraDescriptors

								%>
								<!-- END:  Summary -->
								<BR />
								</TD>
							</TR>
							<%End If%>
							<!-- END: Each Prompt -->

							<TR>
								<TD VALIGN="BOTTOM">
								</TD>
							</TR>

							<%If aPromptGeneralInfo(PROMPT_L_ACTIVEPROMPT) >= 2 Then%>
							<TR>
								<TD BGCOLOR="#CCCC99" ALIGN="LEFT" VALIGN="BOTTOM"><IMG SRC="Images/loginLowerLeftCorner.gif" WIDTH="11" HEIGHT="11" /></TD>
								<TD VALIGN="BOTTOM">
									<TABLE BGCOLOR="#CCCC99" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
										<TR><TD><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
									</TABLE>
								</TD>
								<TD BGCOLOR="#CCCC99"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
							</TR>
							<%Else%>
							<TR>
								<TD><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="10" ALT="" BORDER="0" /></TD></TR>
							</TR>
							<TR>
								<TD BGCOLOR="#CCCC99" COLSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
							</TR>
							<%End If%>
						</TABLE>
						<!-- END: new prompt widget -->

						</TD>
					</TR>
					</TABLE>

					<%If aPromptGeneralInfo(PROMPT_B_ALLOW_PROFILE) Then
						Call DisplayProfileNameDesc(aConnectionInfo, aPromptGeneralInfo)
					End If%>

					<!-- BEGIN:  EXECUTE & CANCEL BAR -->
					<TABLE BORDER="0" COLS="3" WIDTH="100%" CELLPADDING="0">
						<TR>
							<TD WIDTH="1%" ROWSPAN="2"><IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" ALT="" BORDER="0" /></TD>
							<TD COLSPAN="2" BGCOLOR="#cccccc"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
						</TR>

						<TR>
							<%If aPromptGeneralInfo(PROMPT_B_SUMMARY) Then%>
							<TD><INPUT TYPE="SUBMIT" CLASS="buttonClass" NAME="PromptBack" VALUE="<%Response.write asDescriptors(340) 'Descriptor: Back to Prompt%>" /></TD>
							<%End If%>

							<TD ALIGN="LEFT" NOWRAP="1">
								<INPUT TYPE="SUBMIT" CLASS="buttonClass" NAME="HydraBack2" VALUE="<%Response.write asDescriptors(149) 'Descriptor: Back%>" />
								<INPUT TYPE="SUBMIT" CLASS="buttonClass" NAME="HydraNext2" VALUE="<%Response.write asDescriptors(335) 'Descriptor: Next%>" />
							</TD>
							<TD ALIGN="RIGHT" NOWRAP="1">
								<%If aPromptGeneralInfo(PROMPT_B_FINISH_ENABLED) Then%>
								<INPUT TYPE="SUBMIT" CLASS="buttonClass" NAME="HydraFinish2" VALUE="<%Response.write asDescriptors(442) 'Descriptor: Back to Prompt%>" />
								<%End If%>
								<INPUT TYPE="SUBMIT" CLASS="buttonClass" NAME="cancel" VALUE="<%Response.Write asDescriptors(120) 'Descriptor: Cancel%>" />
							</TD>
						</TR>
					</TABLE>
					<!-- END:  EXECUTE & CANCEL BAR -->

					<!-- hidden inputs -->
					<%Call DisplayHiddenValues(oRequest, aPromptGeneralInfo)%>
					</FORM>
				</TD>
			</TR>
		</TABLE>
		<!-- #include file="footer.asp" -->
	</BODY>
</HTML>
<%
Call Clean(aConnectionInfo)
%>