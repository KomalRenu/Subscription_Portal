<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="160">
	<TR>
		<TD>
			<IMG SRC="images/1ptrans.gif" WIDTH="2" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD VALIGN="TOP">
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
				<TR>
				    <TD ROWSPAN="2" VALIGN="TOP">
				        <IMG SRC="images/blacktriDown.gif">
				    </TD>
					<TD VALIGN="TOP">
						<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(360) 'Descriptor: Reports%><%Else%><%Response.Write UCase(asDescriptors(360)) 'Descriptor: Reports%><%End If%></B></FONT>
					</TD>
				</TR>
				<TR>
					<TD>
						<% Call RenderPortalDocumentsList(sSubsXML, sGetAvailableSubscriptionsXML, aDocInfo(DOC_SUBS_ID))%>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>