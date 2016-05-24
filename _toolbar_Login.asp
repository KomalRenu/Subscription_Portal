<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
	    <TD>
	        <TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0">
	            <TR>
	                <TD VALIGN="TOP"><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
	                <TD><A HREF="about.asp" STYLE="text-decoration:none;"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(363)'Descriptor: About%><%Else%><%Response.Write UCase(asDescriptors(363))'Descriptor: About%><%End If%></B></FONT></A></TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
</TABLE>