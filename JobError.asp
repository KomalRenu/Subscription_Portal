<%@LANGUAGE=VBSCRIPT%>
<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="CommonDeclarations.asp" -->
<!-- #include file="CustomLib/JobErrorCuLib.asp" -->
<%
	On Error Resume Next

	Dim aJobErrorRequest()
	Dim sJobErrorMessage
	Dim aObjectInfo()
	Dim aFolderInfo()

    lErrNumber = NO_ERR
	lErrNumber = ReceiveJobErrorRequest(oRequest, aConnectionInfo, aJobErrorRequest, aObjectInfo, aFolderInfo, sErrDescription)
	If lErrNumber = NO_ERR Then
		If Len(CStr(oRequest("FolderID").Item)) = 0 Then
			lErrNumber = GetObject(oRequest, aConnectionInfo, oSession, oObjServer, aObjectInfo, sErrDescription)
		End If
		If lErrNumber = NO_ERR Then
			'*** If there is an error,  cancel the message and go back to a diffenret page ***'
			If Len(aJobErrorRequest(S_OK_JOBERROR)) > 0 Then
				'If Len(aJobErrorRequest(S_PREVIOUS_PAGE_JOBERROR)) > 0 Then
				'	Call CancelMessage(aJobErrorRequest, aConnectionInfo, sErrDescription)
				'	Response.Redirect aJobErrorRequest(S_PREVIOUS_PAGE_JOBERROR) & "?" & aConnectionInfo(S_PROJECT_URL_CONNECTION)
				'Else
					'Response.Redirect "Folder.asp?Page=" & aPageInfo(N_ALIAS_PAGE) & "&FolderID=" & CStr(oRequest("FolderID").Item) & "&" & aConnectionInfo(S_PROJECT_URL_CONNECTION)
				'End If
				'''''clear cache file
				Response.Redirect "Services.asp?folderID=" & CStr(oRequest("FolderID"))

			'*** Get the Error description from the XML ***'
			Else
				Call GetErrorMessageForMsgID(aJobErrorRequest, aConnectionInfo, sJobErrorMessage, sErrDescription)
			End If
			aFolderInfo(S_OBJECT_ID_OBJECT) = aObjectInfo(S_PARENT_ID_OBJECT)
			lErrNumber = GetObject(oRequest, aConnectionInfo, oSession, oObjServer, aFolderInfo, sErrDescription)
		End If
	End If
	'Check for all the different errors we might find across all the pages
	Call CheckCommonErrors(lErrNumber, aConnectionInfo, aPageInfo, oRequest, sErrDescription)

	'Initialization of GenericHeader.asp variables
	Select Case aPageInfo(N_ALIAS_PAGE)
		Case DssXmlFolderNameProfileReports
			aPageInfo(N_CURRENT_OPTION_PAGE) = MY_REPORTS_LINK
		Case DssXmlFolderNameTemplateReports
			aPageInfo(N_CURRENT_OPTION_PAGE) = NEW_REPORT_LINK
		Case DssXmlFolderNamePublicReports
			aPageInfo(N_CURRENT_OPTION_PAGE) = SHARED_REPORTS_LINK
		Case Else
			aPageInfo(N_CURRENT_OPTION_PAGE) = NO_LINK
	End Select

	aPageInfo(S_FOLDER_ID_PAGE) = aObjectInfo(S_PARENT_ID_OBJECT)
	aPageInfo(S_ROOT_FOLDER_ID_PAGE) = aObjectInfo(S_ROOT_ID_OBJECT)
%>
<HTML>
	<HEAD>
		<%Response.Write(putMETATagWithCharSet())%>
		<TITLE><%Response.Write asDescriptors(148) 'Descriptor: Error executing the report%>. MicroStrategy Web.</TITLE>
	</HEAD>
	<BODY BGCOLOR="#FFFFFF" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="header_multi.asp" -->
		<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
			<TR>
				<TD VALIGN="TOP" WIDTH="1%">
					<!-- #include file="_toolbar_jobError.asp" -->
					<!-- include file="_toolbar_Login.asp" -->
				</TD>
				<TD VALIGN="TOP" WIDTH="100%">
					<FONT SIZE="1"><BR /><FONT>
					<%'<** Begin path for this Report **>
					Call DisplayObjectPath(aConnectionInfo, aObjectInfo)
					'<** End path for this Report **>
					%><BR /><BR />
					<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%">
						<TR>
							<TD VALIGN="TOP" WIDTH="1%">
								<IMG SRC="Images/jobError.gif" WIDTH="55" HEIGHT="65" ALT="" BORDER="0" />
							</TD>
							<TD>&nbsp;</TD>
							<TD ALIGN="LEFT">
								<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>">
									<FONT COLOR="#CC0000"><B><%If aObjectInfo(L_TYPE_OBJECT) = DssXmlTypeDocumentDefinition Then
										Response.Write asDescriptors(302) 'Descriptor: Error in document results.
									Else
										Response.Write asDescriptors(273) 'Descriptor: Error in report results.
									End If%></B></FONT><BR /><BR />
									<%Response.Write asDescriptors(253) & " " 'Descriptor: Your request could not be processed due to a server error.
									Response.Write asDescriptors(254) 'Descriptor: Please try again.  If the error persists, contact the server administrator.%>
									<BR /><BR />
									<%Response.Write DisplayContactInformation("")
									If Len(aJobErrorRequest(S_MSG_JOBERROR)) > 0 Then
										Response.Write DisplayError("", aJobErrorRequest(S_MSG_JOBERROR))
									ElseIf Len(sJobErrorMessage) > 0 Then
										Response.Write DisplayError("", asDescriptors(274) & " " & sJobErrorMessage) 'Descriptor: MicroStrategy Server error:
									ElseIf Len(aJobErrorRequest(S_ERROR_DESCRIPTION_JOBERROR)) > 0 Then
										If aJobErrorRequest(L_ERROR_NUMBER_JOBERROR) <> NO_ERR Then
											Response.Write DisplayError("", aJobErrorRequest(S_ERROR_DESCRIPTION_JOBERROR))
										End If
									ElseIf lErrNumber <> NO_ERR Then
										If aObjectInfo(L_TYPE_OBJECT) = DssXmlTypeDocumentDefinition Then
											Response.Write DisplayError(asDescriptors(302), sErrDescription) 'Descriptor: Error in document results.
										Else
											Response.Write DisplayError(asDescriptors(273), sErrDescription) 'Descriptor: Error in report results.
										End If
									End If%>
								</FONT>
							</TD>
						</TR>
					</TABLE>
					<%If Len(aObjectInfo(S_OBJECT_ID_OBJECT)) > 0 And Len(aObjectInfo(S_PARENT_ID_OBJECT)) > 0 Then%>
						<FORM ACTION="JobError.asp" METHOD="<%Response.Write DecideFormMethod()%>">
							<INPUT TYPE="HIDDEN" NAME="MsgID" VALUE="<%Response.Write aJobErrorRequest(O_RESULTSET_JOBERROR).MessageID%>" />
							<%If aObjectInfo(L_TYPE_OBJECT) = DssXmlTypeReportDefinition Then%>
								<INPUT TYPE="HIDDEN" NAME="ReportID" VALUE="<%Response.Write aObjectInfo(S_OBJECT_ID_OBJECT)%>" />
							<%Else%>
								<INPUT TYPE="HIDDEN" NAME="DocumentID" VALUE="<%Response.Write aObjectInfo(S_OBJECT_ID_OBJECT)%>" />
							<%End If%>
							<INPUT TYPE="HIDDEN" NAME="FolderID" VALUE="<%Response.Write oRequest("FolderID")%>" />
							<INPUT TYPE="HIDDEN" NAME="Server" VALUE="<%Response.Write aConnectionInfo(S_SERVER_NAME_CONNECTION)%>" />
							<INPUT TYPE="HIDDEN" NAME="Project" VALUE="<%Response.Write aConnectionInfo(S_PROJECT_CONNECTION)%>" />
							<INPUT TYPE="HIDDEN" NAME="Port" VALUE="<%Response.Write aConnectionInfo(N_PORT_CONNECTION)%>" />
							<INPUT TYPE="HIDDEN" NAME="Uid" VALUE="<%Response.Write aConnectionInfo(S_UID_CONNECTION)%>" />
							<INPUT TYPE="HIDDEN" NAME="UMode" VALUE="<%Response.Write aConnectionInfo(N_USER_MODE_CONNECTION)%>" />
							<INPUT TYPE="HIDDEN" NAME="Page" VALUE="<%Response.Write aPageInfo(N_ALIAS_PAGE)%>" />
							<INPUT TYPE="HIDDEN" NAME="PreviousPage" VALUE="<%Response.Write aJobErrorRequest(S_PREVIOUS_PAGE_JOBERROR)%>" />
							<%If Len(aJobErrorRequest(S_PREVIOUS_PAGE_JOBERROR)) > 0 Then%>
								<INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON" NAME="Ok" VALUE="<%Response.Write asDescriptors(114)'Descriptor: Continue%>" />
							<%Else%>
								<INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON" NAME="Ok" VALUE="<%Response.Write asDescriptors(383) 'Descriptor: Back to Services%>" />
							<%End If%>
						</FORM>
					<%Else%>
						<FORM ACTION="Desktop.asp" METHOD="<%Response.Write DecideFormMethod()%>">
							<INPUT TYPE="HIDDEN" NAME="Server" VALUE="<%Response.Write aConnectionInfo(S_SERVER_NAME_CONNECTION)%>" />
							<INPUT TYPE="HIDDEN" NAME="Project" VALUE="<%Response.Write aConnectionInfo(S_PROJECT_CONNECTION)%>" />
							<INPUT TYPE="HIDDEN" NAME="Port" VALUE="<%Response.Write aConnectionInfo(N_PORT_CONNECTION)%>" />
							<INPUT TYPE="HIDDEN" NAME="Uid" VALUE="<%Response.Write aConnectionInfo(S_UID_CONNECTION)%>" />
							<INPUT TYPE="HIDDEN" NAME="UMode" VALUE="<%Response.Write aConnectionInfo(N_USER_MODE_CONNECTION)%>" />
							<INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON" NAME="Ok" VALUE="<%Response.Write asDescriptors(1)'Descriptor: Home%>" />
						</FORM>
					<%End If%>
					<BR />
					<%If aObjectInfo(L_TYPE_OBJECT) = DssXmlTypeReportDefinition Then
						If Len(aJobErrorRequest(S_REQUEST_JOBERROR)) > 0 Then%>
							<A HREF="Report.asp?<%Response.Write aJobErrorRequest(S_REQUEST_JOBERROR)%>"><%Reponse.Write asDescriptors(247) 'Descriptor: Go to Previous Report State%></A>
						<%End If
					End If%>
				</TD>
			</TR>
			<!-- #include file="footer.asp" -->
		</TABLE>
	</BODY>
</HTML>
<%
Set oRequest = Nothing
Set oSession = Nothing
Set oObjServer = Nothing
Set aObjectInfo(O_CONTENTS_XML_OBJECT) = Nothing
Set aFolderInfo(O_CONTENTS_XML_OBJECT) = Nothing
Erase aObjectInfo
Erase aFolderInfo
Erase aiQuestions
%>