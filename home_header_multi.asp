<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Dim sSelectedTabColor
	Dim homeColor
	Dim sBarTitle
	sSelectedTabColor = ""
	homeColor = "000000"
	sBarTitle = ""

	If Len(LoggedInStatus()) = 0 Then
	    sChannel = ""
	End If
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
							>&nbsp;&nbsp;<A HREF="default.asp" STYLE="text-decoration:none;"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#<%If Len(sChannel) = 0 Then Response.Write "eedd82" Else Response.Write "ffffff" End If%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B
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
							<TR><TD HEIGHT="24" ALIGN="CENTER" NOWRAP="1" WIDTH="100%"
							>&nbsp;&nbsp;<A HREF="logout.asp" STYLE="text-decoration:none;"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="ffffff" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B
								><NOBR><%Response.Write asDescriptors(4) 'Descriptor: Logout%></NOBR></B></FONT></A>&nbsp;&nbsp;<BR/>
								</TD>
							</TR>
						</TABLE>
					</TD>
			    </TR>
			</TABLE>
		</TD>
	</TR>
	<!-- end header tabs -->
</TABLE>