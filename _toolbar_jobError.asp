<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%If aPageInfo(N_TOOLBARS_PAGE) AND MAIN_TOOLBAR Then %>
 <TABLE WIDTH="158" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<TR>
		<TD><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="3" ALT="" BORDER="0" /></TD>
	</TR>
	<!-- #include file="_toolbar_helpSection.asp" -->
 </TABLE>
<%Else%>
	<!-- #include file="_noToolbar.asp" -->
<%End If%>