<%@LANGUAGE=VBSCRIPT%>
<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/ShowErrorsCuLib.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<%
Dim oFS
Dim oFile
Dim oDoc, oSortedXML, opreXML
Dim oNode, oElement
Dim oSortXSL, oStyleXSL
Dim sWholeFile, sSorted, sResult, sSelected
Dim sLogFileName
Dim sOrderBy, sSortOrder, sWarn, sErr, sMessage, sPrevious, sNext, iFirstIndex, iLastIndex, sDay, sMonth, sYear, bReturn
Dim i, iErrors
Dim bValidDate
Dim sLanguage
Dim asDisplayDate(5)
Dim lStatus

    lStatus = checkSiteConfiguration()

'Added to use General Header
sChannel = ""

Call RecieveShowErrorsRequest(oRequest, sOrderBy, sSortOrder, sWarn, sErr, sMessage, iFirstIndex, iLastIndex, sDay, sMonth, sYear, bReturn, asDisplayDate)

If bReturn Then
	Erase aiQuestions
	If Len(CStr(oRequest("src"))) > 0 Then
		Response.Redirect CStr(oRequest("src"))
	Else
		Response.Redirect "."
	End If
End If

bValidDate = IsValidDate(sDay, sMonth, sYear)
If bValidDate Then
	sLogFileName = sYear & "-" & sMonth & "-" & sDay & "err.log"

	Set oFS = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFS.OpenTextFile(server.MapPath("..\Logs\" & sLogFileName), FOR_READING_SHOWERRORS)
	sWholeFile = oFile.readAll()
	sWholeFile = "<ERRORS>" + sWholeFile + "<options>"
	sWholeFile = sWholeFile + "<OrderBy>" + sOrderBy + "</OrderBy>"
	sWholeFile = sWholeFile + "<SortOrder>" + sSortOrder + "</SortOrder>"
	sWholeFile = sWholeFile + "<Error>" + sErr + "</Error>"
	sWholeFile = sWholeFile + "<Warning>" + sWarn + "</Warning>"
	sWholeFile = sWholeFile + "<Message>" + sMessage + "</Message>"
	sWholeFile = sWholeFile + "</options>" + "</ERRORS>"
	oFile.close()

	Set oFS = Nothing
	Set oFile = Nothing
End If

'Initialization of GenericHeader.asp variables
aPageInfo(N_CURRENT_OPTION_PAGE) = ADMIN_LINK
'aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = HOME_LINK + PROJECTS_LINK
sBGColor = "000000"
aPageInfo(S_TITLE_PAGE) = asDescriptors(173) 'Description: Error Message List
aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForShowErrors(oRequest)
%>
<HTML>
	<HEAD>
		<%Response.Write(putMETATagWithCharSet())%>
		<TITLE><%
			Response.Write asDescriptors(173) 'Descriptor: Error Message List
			If bValidDate Then
				Response.Write(": " & sLogFileName)
			End If
		%>. MicroStrategy DSS Web.</TITLE>
	</HEAD>
	<BODY BGCOLOR="#FFFFFF" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<FORM ACTION="ShowErrors.asp" METHOD="GET">
			<!-- #include file="admin_header.asp" -->
			<TABLE WIDTH="98%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
				<TR>
					<TD VALIGN="TOP" WIDTH="1%">
						<!-- #include file="_toolbar_showerrors.asp"-->
					</TD>
					<TD WIDTH="1%" ROWSPAN="2">&nbsp;&nbsp;&nbsp;</TD>
					<TD VALIGN="TOP" WIDTH="98%">
						<%If Not bValidDate Then Response.Write asDescriptors(294) 'Descriptor: Please enter a valid date.
						If bValidDate Then
							Set oDoc = Server.CreateObject("Microsoft.XMLDOM")
							oDoc.async = False
							oDoc.loadXML(sWholeFile)
							oDoc.save(server.MapPath("saved.xml"))
							If oDoc.parseError.errorCode <> 0 Then
								Call ReportError(sLogFileName, oDoc.parseError)
							Else 'Load the stylesheet
								Set oSortXSL = Server.CreateObject("Microsoft.XMLDOM")
								oSortXSL.async = False
								oSortXSL.load(Server.MapPath("SortErrors.xsl"))
								If oSortXSL.parseError.errorCode <> 0 Then
									Call ReportError("SortErrors.xsl", oSortXSL.parseError)
								Else 'sort the errors
									Set oSortedXML = Server.CreateObject("Microsoft.XMLDOM")
									oSortedXML.async = False
									Set opreXML = Server.CreateObject("Microsoft.XMLDOM")
									opreXML.async = False

								    sSorted = oDoc.transformNode(oSortXSL)
								    oSortedXML.loadXML(sSorted)
								    oSortedXML.save(Server.MapPath("sorted.xml"))

									Set oNode = opreXML.createElement("ERRORS")
									Set opreXML.documentElement = oNode

									'create the correct top bound
								    iErrors = oSortedXML.documentElement.childNodes.length
								    If iLastIndex > iErrors Then iLastIndex = iErrors
								    If iErrors = 0 Then iFirstIndex = 0

									For i = iFirstIndex To iLastIndex
										'appending the node from oSortedXML removes it, so the index remains the same, but is a different node
										Set oNode = oSortedXML.documentElement.childNodes.item(iFirstIndex - 1)
										opreXML.documentElement.appendChild(oNode)
								    Next

									If lErrNumber <> 0 Then
										Call ReportError(sLogFileName, oDoc.parseError)
										Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "ShowErrors.asp", "ShowErrors page", "", "Error in call to AddInputsToXML function", LogLevelTrace)
									Else%>
										<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2" ALIGN="CENTER">
											<TR>
												<TD>
													<%If iFirstIndex > 1 Then%>
														<INPUT TYPE="IMAGE" NAME="bPrev" SRC="../Images/arrow_left_inc_fetch.gif" WIDTH="5" HEIGHT="10" ALT="<%Response.Write(Replace(asDescriptors(279), "##", MESSAGES_SHOWERRORS)) 'Descriptor: Previous ##%>" BORDER="0" /></TD>
													<%Else%>
														<IMG SRC="../Images/arrow_left_inc_fetch_disabled.gif" WIDTH="5" HEIGHT="10" ALT="" BORDER="0" />
													<%End If%>
												</TD>
												<TD><B><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write Replace(Replace(Replace(asDescriptors(57), "####", iErrors), "###", iLastIndex), "##", iFirstIndex) 'Descriptor: Messages ## - ### of ####%></FONT></B></TD>
												<TD>
													<%If iLastIndex < iErrors Then%>
														<INPUT TYPE="IMAGE" NAME="bNext" SRC="../Images/arrow_right_inc_fetch.gif" WIDTH="5" HEIGHT="10" ALT="<%Response.Write(Replace(asDescriptors(280), "##", MESSAGES_SHOWERRORS)) 'Descriptor: Next ##%>" BORDER="0" /></TD>
													<%Else%>
														<IMG SRC="../Images/arrow_right_inc_fetch_disabled.gif" WIDTH="5" HEIGHT="10" ALT="" BORDER="0" />
													<%End If%>
												</TD>
											</TR>
										</TABLE>
										<%
										Call ShowErrors(opreXML, aFontInfo, sOrderBy, sSortOrder)
										If lErrNumber <> 0 Then
											Call ReportError("ShowErrors.xsl", sErrDescription)
											Call LogErrorXML(aConnectionInfo, lErrNumber, sErrDescription, Err.source, "ShowErrors.asp", "ShowErrors page", "", "Error in call to AddInputsToXML function", LogLevelTrace)
									    End If
									End If
								End If
							End If%>
							<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2" ALIGN="CENTER">
								<TR>
									<INPUT TYPE="HIDDEN" NAME="hFirst" VALUE="<%Response.Write iFirstIndex%>" />
									<INPUT TYPE="HIDDEN" NAME="hLast" VALUE="<%Response.Write iLastIndex%>" />
									<INPUT TYPE="HIDDEN" NAME="OrderBy" VALUE="<%Response.Write sOrderBy%>" />
									<INPUT TYPE="HIDDEN" NAME="SortOrder" VALUE="<%Response.Write sSortOrder%>" />
									<INPUT TYPE="HIDDEN" NAME="src" VALUE="<%=CStr(oRequest("src"))%>" />

									<TD>
										<%If iFirstIndex > 1 Then%>
											<INPUT TYPE="IMAGE" NAME="bPrev" SRC="../Images/arrow_left_inc_fetch.gif" WIDTH="5" HEIGHT="10" ALT="<%Response.Write(Replace(asDescriptors(279), "##", MESSAGES_SHOWERRORS)) 'Descriptor: Previous ##%>" BORDER="0" /></TD>
										<%Else%>
											<IMG SRC="../Images/arrow_left_inc_fetch_disabled.gif" WIDTH="5" HEIGHT="10" ALT="" BORDER="0" />
										<%End If%>
									</TD>
									<TD><B><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><% Response.Write Replace(Replace(Replace(asDescriptors(57), "####", iErrors), "###", iLastIndex), "##", iFirstIndex) 'Descriptor: Messages ## - ### of ####%></FONT></B></TD>
									<TD>
										<%If iLastIndex < iErrors Then%>
											<INPUT TYPE="IMAGE" NAME="bNext" SRC="../Images/arrow_right_inc_fetch.gif" WIDTH="5" HEIGHT="10" ALT="<%Response.Write(Replace(asDescriptors(280), "##", MESSAGES_SHOWERRORS)) 'Descriptor: Next ##%>" BORDER="0" /></TD>
										<%Else%>
											<IMG SRC="../Images/arrow_right_inc_fetch_disabled.gif" WIDTH="5" HEIGHT="10" ALT="" BORDER="0" />
										<%End If%>
									</TD>
								</TR>
							</TABLE>
						<%End If%>
					</TD>

					<TD WIDTH="1%">
				    </TD>
				</TR>
			</TABLE>
		</FORM>
	</BODY>
</HTML>
<%
Set oDoc = Nothing
Set oSortXSL = Nothing
Set oSortedXML = Nothing
Set opreXML = Nothing
Set oNode = Nothing
Set oElement = Nothing
Set oStyleXSL = Nothing
Erase aiQuestions
%>