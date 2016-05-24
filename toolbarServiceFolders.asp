<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD>
			<img src="images/1ptrans.gif" WIDTH="2" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD VALIGN=TOP>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
				<TR>
					<TD>
						<font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#444444" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><b><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(37)'Descriptor: View mode%><%Else%><%Response.Write UCase(asDescriptors(37))'Descriptor: View mode%><%End If%></b></font>
					</TD>
				</TR>
				<TR>
				    <TD>
				        <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
				        <%If StrComp(GetServiceViewMode(), "1", vbBinaryCompare) = 0 Then%>
				            <TR>
				                <TD><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
				                <TD><b><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>" color="#444444" ><%Response.Write asDescriptors(308) 'Descriptor: Large icons view%></font></b></TD>
				            </TR>
				            <TR>
				                <TD><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
				                <TD><A HREF="services.asp?svm=2<%If Len(sFolderID) > 0 Then%>&folderID=<%Response.Write sFolderID%><%End If%>"><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#444444" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(307) 'Descriptor: List view%></font></A></TD>
				            </TR>
				        <%Else%>
				            <TR>
				                <TD><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
				                <TD><A HREF="services.asp?svm=1<%If Len(sFolderID) > 0 Then%>&folderID=<%Response.Write sFolderID%><%End If%>"><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#444444" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(308) 'Descriptor: Large icons view%></font></A></TD>
				            </TR>
				            <TR>
				                <TD><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
				                <TD><b><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>" color="#444444"><%Response.Write asDescriptors(307) 'Descriptor: List view%></font></b></TD>
				            </TR>
				        <%End If%>
				        </TABLE>
				    </TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>