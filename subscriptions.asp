<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/SubscriptionsCuLib.asp" -->
<!-- #include file="CustomLib/FoldersCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<!-- #include file="CustomLib/HomeCuLib.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sGetFolderContentsXML
	Dim sGetUserSubscriptionsXML
	Dim sGetAvailableSubscriptionsXML
	Dim sServiceID
	Dim sFolderID
	Dim sDeliv_SortOrder
	Dim sDeliv_OrderBy
	Dim sRep_SortOrder
	Dim sRep_OrderBy
	Dim asSubscriptionGUIDS()
	Dim sPortalDeviceID

    Dim sSitesXML
    Dim aSiteInfo(2)

	sSubscriptionsStyle = ""
	aPageInfo(S_NAME_PAGE) = "subscriptions.asp"
	sPortalDeviceID = Application.Value("Portal_device")

	lErr = ParseRequestForSubscriptions(oRequest, sServiceID, sFolderID, sDeliv_SortOrder, sDeliv_OrderBy, sRep_SortOrder, sRep_OrderBy)

	If lErr = NO_ERR Then
		If sFolderID <> "" Then
		    lErr = cu_GetFolderContents(sFolderID, sGetFolderContentsXML)
		End If
	End If

	If lErr = NO_ERR Then
        lErr = cu_GetUserSubscriptions(sGetUserSubscriptionsXML)
	End If

    If lErr = NO_ERR Then
        If Len(sPortalDeviceID) > 0 Then
            lErr = GetSubscriptionsArray_Reports(sGetUserSubscriptionsXML, asSubscriptionGUIDS)
            If lErr = NO_ERR Then
                If UBound(asSubscriptionGUIDS) <> -1 Then
                    lErr = cu_GetAvailableSubscriptions(asSubscriptionGUIDS, sGetAvailableSubscriptionsXML)
                End If
            End If
        End If
    End If

    If lErr = NO_ERR Then
        lErr = cu_GetChannels(sSitesXML)
    End If

    If lErr = NO_ERR Then
        lErr = cu_GetSiteInfo(sSitesXML, aSiteInfo)
    End If

%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(354)'Descriptor: Subscriptions%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
<!-- #include file="header_multi.asp" -->
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
	<TR>
		<TD WIDTH="1%" valign="TOP">
            <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="140">
                <TR><TD>
                    <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
                        <TR><TD BGCOLOR="#cccccc">
                            <TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 WIDTH="100%">
                                <TR BGCOLOR="#ffffff">
                                    <TD><A HREF="services.asp?folderID=<%=aSiteInfo(2)%>"><IMG SRC="images/desktop_newReport.gif" HEIGHT="60" WIDTH="60" ALT="" BORDER="0" /></A></TD>
                                    <TD><A HREF="services.asp?folderID=<%=aSiteInfo(2)%>"><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#cc0000" size="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><b><%Response.Write asDescriptors(452) 'Descriptor: Sign up for a service%></b></font></A></TD>
                                </TR>
                            </TABLE>
                        </TD></TR>
                    </TABLE>
                </TD></TR>
            </TABLE>
			<!-- begin left menu -->
			    <!-- #include file="_toolbar_Subscriptions.asp" -->
			<BR />
			<!-- end left menu -->
			<img src="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="1%">
			<IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="97%" VALIGN="TOP">
			<!-- begin center panel -->
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
			    <TR>
			        <TD VALIGN="CENTER">
			            <%Call RenderPath_Subscriptions(sServiceID, sGetFolderContentsXML)%>
			        </TD>
			        <TD ALIGN=RIGHT><IMG SRC="images/desktop_ScheduledReports.gif" WIDTH=60 HEIGHT=60 BORDER=0 ALT="" /></TD>
			    </TR>
			</TABLE>
			<% If lErr <> 0 Then %>
				<% Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(380), "home.asp") 'Descriptor: Back to Home%>
			<% Else %>
				<%
				    If Len(sPortalDeviceID) > 0 Then
				        If StrComp(GetSubscriptionsViewMode(), N_VIEW_LARGE_ICONS, vbBinaryCompare) = 0 Then%>
				            <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
				            	<TR>
				            		<TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><b><%Response.Write asDescriptors(360)'Descriptor: Reports%></b></font></TD>
				            	</TR>
				            	<TR>
				            	    <TD bgcolor="#000000"><img src="images/1ptrans.gif" HEIGHT="1" WIDTH="1" ALT="" BORDER="0" /></TD>
				            	</TR>
				            </TABLE>
                            <%Call RenderLargeIcons_Reports(sServiceID, sFolderID, sGetUserSubscriptionsXML, sGetAvailableSubscriptionsXML)
				        Else%>
				            <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
				            	<TR>
				            	    <TD><IMG SRC="images/reports.gif" WIDTH="50" HEIGHT="42" BORDER="0" ALT="" /></TD>
				            		<TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><b><%Response.Write asDescriptors(360)'Descriptor: Reports%></b></font></TD>
				            	</TR>
				            </TABLE>
                            <%Call RenderList_Reports(sRep_SortOrder, sRep_OrderBy, sServiceID, sFolderID, sGetUserSubscriptionsXML, sGetAvailableSubscriptionsXML)
				        End If
				%>
				<BR />
				<%End If
				    If StrComp(GetSubscriptionsViewMode(), N_VIEW_LARGE_ICONS, vbBinaryCompare) = 0 Then%>
				        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
				        	<TR>
				        		<TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><b><%Response.Write asDescriptors(447) 'Descriptor: Deliveries%></b></font></TD>
				        	</TR>
				        	<TR>
				        	    <TD bgcolor="#000000"><img src="images/1ptrans.gif" HEIGHT="1" WIDTH="1" ALT="" BORDER="0" /></TD>
				        	</TR>
				        </TABLE>
				        <%Call RenderLargeIcons_Deliveries(sServiceID, sFolderID, sGetUserSubscriptionsXML)
				    Else%>
				        <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
				        	<TR>
				        		<TD><IMG SRC="images/deliveries.gif" WIDTH="50" HEIGHT="42" BORDER="0" ALT="" /></TD>
				        		<TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><b><%Response.Write asDescriptors(447) 'Descriptor: Deliveries%></b></font></TD>
				        	</TR>
				        </TABLE>
				        <%Call RenderList_Deliveries(sDeliv_SortOrder, sDeliv_OrderBy, sServiceID, sFolderID, sGetUserSubscriptionsXML)
				    End If
				%>
			<% End If %>
			<!-- end center panel -->
		</TD>
		<TD WIDTH="1%">
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
<%
    Erase asSubscriptionGUIDS
%>