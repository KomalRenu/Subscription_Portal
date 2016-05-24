<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/HomeCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
    If Len(LoggedInStatus()) = 0 Then
        Response.Redirect "login.asp"
    End If

    Dim sSitesXML
    Dim aSiteInfo(2)
    Dim aVersionInfo(1)
    Dim sPortalDeviceID

    sHomeStyle = ""
    sPortalDeviceID = Application.Value("Portal_Device")

    If lErr = NO_ERR Then
        lErr = cu_GetVersions(aVersionInfo)
    End If

    If lErr = NO_ERR Then
        lErr = cu_GetChannels(sSitesXML)
    End If

    If lErr = NO_ERR Then
        lErr = cu_GetSiteInfo(sSitesXML, aSiteInfo)
    End If

    'Filter out everything after the paranthesis which would come for Iserver Authentication users.
    'This is just for disply purposes
    Dim sUserName
    Dim ipos
    sUserName = GetUserName()
    ipos = Instr(1,sUserName,"(")
    If ipos > 0 then
		sUserName = Mid(sUserName, 1, ipos - 1 )
    End IF



%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(1)'Descriptor: Home%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="home_header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
    <TR>
        <TD WIDTH="1%" VALIGN="TOP" ROWSPAN="3">
		    <TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0">
				<TR>
				    <TD>
				       <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(1)'Descriptor: Home%><%Else%><%Response.Write UCase(asDescriptors(1))'Descriptor: Home%><%End If%></B></FONT>
				    </TD>
				</TR>
                <TR>
                    <TD><IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT=""></TD>
                </TR>
		    </TABLE>
		</TD>
		<TD WIDTH="98%" VALIGN="TOP">
			<!-- begin center panel -->
            <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
                <TR>
                    <TD NOWRAP>
                        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%Response.Write asDescriptors(283) & " " & sUserName 'Descriptor: Welcome%></B> (<%Response.Write Replace(asDescriptors(284), "##", sUserName) 'Descriptor: If you are not ##,%> <A HREF="logout.asp"><%Response.Write asDescriptors(285) 'Descriptor: click here%></A>.)</FONT>
                    </TD>
                    <TD ALIGN="RIGHT">
                        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.write DisplayDateAndTime(Date, "")%></FONT>
                    </TD>
			    </TR>
			</TABLE>

			<!-- begin Desktop Content: -->

            <%
            If lErr <> NO_ERR Then
                Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(380), "home.asp") 'Descriptor: Back to Home
            Else
            %>
			<BR />
            <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
                <TR>
                    <!-- Services -->
                    <TD WIDTH="50%" VALIGN="TOP">
                        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
                            <TR>
                                <TD WIDTH="64" VALIGN="TOP" ALIGN="LEFT">
                                    <A HREF="services.asp?folderID=<%=aSiteInfo(2)%>"><IMG SRC="images/desktop_newReport.gif" WIDTH="60" HEIGHT="60" ALT="" BORDER="0" /></A></TD>
								<TD VALIGN="TOP">
                                    <A HREF="services.asp?folderID=<%=aSiteInfo(2)%>" STYLE="text-decoration:none;"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#CC0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(452)'Descriptor:Sign up for a service%></B></FONT></A><BR />
                                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><BR /><%Response.Write  asDescriptors(453)'Descriptor:Browse a list of prepared reports available to all users.%> <BR /></FONT>
                                </TD>
                            </TR>
                        </TABLE>
					</TD>
					<TD><IMG SRC="Images/1ptrans.gif" WIDTH="2" HEIGHT="1" ALT="" BORDER="0" /></TD>
					<!-- subscribe -->
                    <TD WIDTH="50%" VALIGN="TOP">
                        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
                            <TR>
                                <TD WIDTH="64" VALIGN="TOP" ALIGN="LEFT">
                                    <A HREF="subscriptions.asp"><IMG SRC="images/desktop_ScheduledReports.gif" WIDTH="60" HEIGHT="60" ALT="" BORDER="0" /></A></TD>
								<TD VALIGN="TOP">
                                    <A HREF="subscriptions.asp" STYLE="text-decoration:none;"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#CC0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(354)'Descriptor:Subscriptions%></B></FONT></A><BR />
                                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><BR /><%Response.Write  asDescriptors(346)'Descriptor:View a list of the reports to which you are subscribed.%> <BR /></FONT>
                                </TD>
                            </TR>
                        </TABLE>
					</TD>
                </TR>
                <TR><TD COLSPAN="4"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="20" ALT="" BORDER="0" /></TD></TR>
                <TR>
                    <%If Len(sPortalDeviceID) > 0 Then %>
					<!-- reports -->
                    <TD WIDTH="50%" VALIGN="TOP">
                        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
                            <TR>
                                <TD WIDTH="64" VALIGN="TOP" ALIGN="LEFT">
                                    <A HREF="reports.asp"><IMG SRC="images/desktop_SharedReports.gif" WIDTH="60" HEIGHT="60" ALT="" BORDER="0" /></A></TD>
								<TD VALIGN="TOP">
                                    <A HREF="reports.asp" STYLE="text-decoration:none;"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#CC0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(360)'Descriptor:Reports%></B></FONT></A><BR />
                                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><BR /><%Response.Write  asDescriptors(454) 'Descriptor: View your scheduled reports.%> <BR /></FONT>
                                </TD>
                            </TR>
                        </TABLE>
					</TD>
                    <TD><IMG SRC="Images/1ptrans.gif" WIDTH="2" HEIGHT="1" ALT="" BORDER="0" /></TD>
                    <%End If %>
					<!-- addresses -->
                    <TD WIDTH="50%" VALIGN="TOP">
                        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
                            <TR>
                                <TD WIDTH="64" VALIGN="TOP" ALIGN="LEFT">
                                    <A HREF="addresses.asp"><IMG SRC="images/desktop_Addresses.gif" WIDTH="60" HEIGHT="60" ALT="" BORDER="0" /></A></TD>
								<TD VALIGN="TOP">
                                    <A HREF="addresses.asp" STYLE="text-decoration:none;"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#CC0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(361)'Descriptor:Addresses%></B></FONT></A><BR />
                                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><BR /><%Response.Write  asDescriptors(455)'Descriptor:Specify your contact information for your subscription deliveries.%> <BR /></FONT>
                                </TD>
                            </TR>
                        </TABLE>
					</TD>
                    <%If Len(sPortalDeviceID) = 0 Then %>
                    <TD><IMG SRC="Images/1ptrans.gif" WIDTH="2" HEIGHT="1" ALT="" BORDER="0" /></TD>
                    <TD></TD>
                    <%End If %>

                </TR>
                <TR><TD COLSPAN="4"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="20" ALT="" BORDER="0" /></TD></TR>
                <TR>
					<!-- Preferences -->
                    <TD WIDTH="50%" VALIGN="TOP">
                        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
                            <TR>
                                <TD WIDTH="64" VALIGN="TOP" ALIGN="LEFT">
                                    <A HREF="options.asp"><IMG SRC="images/desktop_preferences.gif" WIDTH="60" HEIGHT="60" ALT="" BORDER="0" /></A></TD>
								<TD VALIGN="TOP">
                                    <A HREF="options.asp" STYLE="text-decoration:none;"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#CC0000" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(286)'Descriptor:Preferences%></B></FONT></A><BR />
                                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><BR /><%Response.Write  asDescriptors(456) 'Descriptor:Customize these screens according to your browser & viewing preferences.%> <BR /></FONT>
                                </TD>
                            </TR>
                        </TABLE>
					</TD>
                    <TD><IMG SRC="Images/1ptrans.gif" WIDTH="2" HEIGHT="1" ALT="" BORDER="0" /></TD>
                </TR>
            </TABLE>
            <!-- End: Desktop Content -->
            <% End If %>

		    <!-- end center panel -->
        </TD>
        <TD WIDTH="1%">
            <IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
        </TD>
    </TR>
    <TR><TD COLSPAN="2"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="20" ALT="" BORDER="0" /></TD></TR>
</TABLE>
<BR>
<!-- begin footer -->
	<!-- #include file="footer.asp" -->
<!-- end footer -->
</BODY>
</HTML>
<%
    Erase aSiteInfo
    Erase aVersionInfo
%>