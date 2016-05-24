<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/ServicesCuLib.asp" -->
<!-- #include file="CustomLib/FoldersCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sGetFolderContentsXML
    Dim sFolderID

	sSubscriptionsStyle = ""
	iSubscribeWizardStep = 1

	lErr = ParseRequestForServices(oRequest, sFolderID)

	If lErr = NO_ERR Then
		'If no folder requested, select the channel root:
		If Len(sFolderID) = 0 Then
			sFolderID = APP_ROOT_FOLDER
		End If
		lErr = cu_GetFolderContents(sFolderID, sGetFolderContentsXML)
	End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE>
		<%Response.Write asDescriptors(362)'Descriptor: Services%> - MicroStrategy Narrowcast Server
	</TITLE>
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
				    <TD>
				        <!-- #include file="toolbarServiceFolders.asp" -->
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
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
			    <TR>
			        <TD VALIGN="CENTER">
			            <%Call RenderPath_Services(sGetFolderContentsXML)%>
			        </TD>
			        <TD ALIGN="RIGHT"><IMG SRC="images/desktop_NewReport.gif" WIDTH="60" HEIGHT="60" BORDER="0" ALT="" /></TD>
			    </TR>
			</TABLE>
			<%
			    If lErr <> NO_ERR Then
			    	Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(383), "services.asp") 'Descriptor: Back to Services
			    Else
			    	If StrComp(GetServiceViewMode(), "1", vbBinaryCompare) = 0 Then
                        Call RenderLargeIcon_Services(sFolderID, sGetFolderContentsXML)
			    	Else
                        Call RenderList_Services(sFolderID, sGetFolderContentsXML)
			    	End If
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