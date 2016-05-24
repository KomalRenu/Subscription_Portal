<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Dim sSelectedTabColor
	Dim homeColor
	Dim sBarTitle
	Dim sPortalDev

	sSelectedTabColor = ""
	homeColor = "000000"
	sBarTitle = ""
	sPortalDev = Application.Value("Portal_Device")

	If Len(sChannel) = 0 Then
		sSelectedTabColor = homeColor
		sBarTitle = asDescriptors(1) 'Descriptor: Home
	End If

%>
<!-- #include file="NSStyleSheet.asp" -->
<TABLE BORDER="0" WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
	<TR style="background-image: url('images/bg_gray.gif'); background-repeat: repeat-x">
		<TD><IMG SRC="images\9_logo.gif"></TD>
		<TD ALIGN="RIGHT">
			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
				<TR>
					<TD ALIGN="RIGHT" VALIGN="BOTTOM" >
						<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" HEIGHT="100%">
							<TR><TD HEIGHT="24" ALIGN="CENTER" NOWRAP="1" WIDTH="100%"
							>&nbsp;&nbsp;<A HREF="default.asp" STYLE="text-decoration:none;"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#<%If Len(sChannel) = 0 Then Response.Write "ffff00" Else Response.Write "ffffff" End If%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B
								><NOBR><%Response.Write asDescriptors(1) 'Descriptor: Home%></NOBR></B></FONT></A>&nbsp;&nbsp;<BR/>
								</TD>
							</TR>
						</TABLE>
					</TD>
					<%
					    Call RenderTabs(sChannel, sSelectedTabColor, sBarTitle, nStart)
					%>
					<TD ALIGN="RIGHT" VALIGN="BOTTOM" >
						<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" HEIGHT="100%">
							<TR><TD HEIGHT="24" ALIGN="CENTER" NOWRAP="1" WIDTH="100%">&nbsp;&nbsp;
								<A HREF="logout.asp" STYLE="text-decoration:none;">
									<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#ffffff" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>">
										<B><NOBR><%Response.Write asDescriptors(4) 'Descriptor: Logout%></NOBR></B></FONT></A>&nbsp;&nbsp;<BR/>
								</TD>
							</TR>
						</TABLE>
					</TD>
			    </TR>
			</TABLE>
		</TD>
	</TR>
	<!-- end header tabs -->

	<!-- begin subheader menu -->
	<TR>
			<TD COLSPAN="3" WIDTH="100%" ALIGN="RIGHT">
				<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" >
					<TR>
						<TD STYLE="<% Response.Write STYLE_BEIGE_BACKGROUND %>" WIDTH="100%"/>
						<TD STYLE="<% Response.Write sHomeStyle %>" NOWRAP>
							<A HREF="home.asp" STYLE="text-decoration:none">
								<FONT COLOR="444444" FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
									&nbsp;&nbsp;<% Response.Write asDescriptors(1)'Descriptor: Home%>&nbsp;&nbsp;
								</FONT>
							</A>
						</TD>

						<TD><IMG SRC="Images/divider_top.gif"/></TD>

						<TD STYLE="<% Response.Write sSubscriptionsStyle %>" NOWRAP>
							<A HREF="subscriptions.asp" STYLE="text-decoration:none">
								<FONT COLOR="444444" FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
									&nbsp;&nbsp;<% Response.Write asDescriptors(354)'Descriptor: Subscriptions %>&nbsp;&nbsp;
								</FONT>
							</A>
						</TD>

						<TD><IMG SRC="Images/divider_top.gif"/></TD>

	 				<%If Len(sPortalDev) > 0 Then %>
						<TD STYLE="<% Response.Write sReportsStyle %>" NOWRAP>
							<A HREF="reports.asp" STYLE="text-decoration:none">
								<FONT COLOR="444444" FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
									&nbsp;&nbsp;<% Response.Write asDescriptors(360)'Descriptor: Reports %>&nbsp;&nbsp;
								</FONT>
							</A>
						</TD>

						<TD><IMG SRC="Images/divider_top.gif"/></TD>

	                <%End If%>

					<TD STYLE="<% Response.Write sAddressesStyle %>" NOWRAP>
						<A HREF="addresses.asp" STYLE="text-decoration:none">
							<FONT COLOR="444444" FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								&nbsp;&nbsp;<% Response.Write asDescriptors(361)'Descriptor: Addresses %>&nbsp;&nbsp;
							</FONT>
						</A>
					</TD>

					<TD><IMG SRC="Images/divider_top.gif"/></TD>

					<TD STYLE="<% Response.Write sOptionsStyle %>" NOWRAP>
						<A HREF="options.asp" STYLE="text-decoration:none">
							<FONT COLOR="444444" FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
								&nbsp;&nbsp;<% Response.Write asDescriptors(286)'Descriptor: Preferences %>&nbsp;&nbsp;
							</FONT>
						</A>
					</TD>
					</TR>
				</TABLE>
			</TD>
	</TR>
	<!-- end subheader menu -->
</TABLE>