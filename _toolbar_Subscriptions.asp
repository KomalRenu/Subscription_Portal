<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD>
			<IMG SRC="images/1ptrans.gif" WIDTH="2" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD VALIGN="TOP">
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
				<TR>
					<TD>
						<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(37)'Descriptor: View mode%><%Else%><%Response.Write UCase(asDescriptors(37))'Descriptor: View mode%><%End If%></B></FONT>
					</TD>
				</TR>
				<TR>
				    <TD>
				        <TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0">
				        <%If StrComp(GetSubscriptionsViewMode(), "1", vbBinaryCompare) = 0 Then%>
				            <TR>
				                <TD><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
				                <TD><B><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(308) 'Descriptor: Large icons view%></FONT></B></TD>
				            </TR>
				            <TR>
				                <TD><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
				                <TD><A HREF="subscriptions.asp?suvm=2<%If Len(sServiceID) > 0 Then%>&serviceID=<%Response.Write sServiceID%><%End If%><%If Len(sFolderID) > 0 Then%>&folderID=<%Response.Write sFolderID%><%End If%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(307) 'Descriptor: List view%></FONT></A></TD>
				            </TR>
				        <%Else%>
				            <TR>
				                <TD><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
				                <TD><A HREF="subscriptions.asp?suvm=1<%If Len(sServiceID) > 0 Then%>&serviceID=<%Response.Write sServiceID%><%End If%><%If Len(sFolderID) > 0 Then%>&folderID=<%Response.Write sFolderID%><%End If%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(308) 'Descriptor: Large icons view%></FONT></A></TD>
				            </TR>
				            <TR>
				                <TD><IMG SRC="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
				                <TD><B><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(307) 'Descriptor: List view%></FONT></B></TD>
				            </TR>
				        <%End If%>
				        </TABLE>
				    </TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>