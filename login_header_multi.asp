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
					<%
					    Call RenderTabs(sChannel, sSelectedTabColor, sBarTitle, nStart)
					%>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<!-- end header tabs -->
</TABLE>