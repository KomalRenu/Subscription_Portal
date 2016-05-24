<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<TABLE WIDTH="160" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<TR>
		<TD WIDTH="10"  BGCOLOR="#CCCCCC" ALIGN="LEFT" VALIGN="TOP"><IMG SRC="images/corner_grey_topleft.gif" WIDTH="10" HEIGHT="10" BORDER="0" ALT=""/></TD>
		<TD WIDTH="130" BGCOLOR="#CCCCCC" ALIGN="CENTER"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(602) 'Descriptor: Select a Device Style%><%Else%><%Response.Write UCase(asDescriptors(602)) 'Descriptor: Select a Device Style%><%End If%></B></FONT></TD>
		<TD WIDTH="11"  BGCOLOR="#CCCCCC" ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="images/corner_grey_topright.gif" WIDTH="11" HEIGHT="10" BORDER="0" ALT=""/></TD>
		<TD WIDTH="9"><IMG SRC="images/1ptrans.gif" WIDTH="9" HEIGHT="1" BORDER="0" ALT="" /></TD>
	</TR>
</TABLE>
<TABLE WIDTH="160" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<TR>
		<TD ROWSPAN="3" WIDTH="2"   BGCOLOR="#CCCCCC"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT="" /></TD>
		<TD ROWSPAN="3" WIDTH="8"   BGCOLOR="#E7E5DF"><IMG SRC="images/1ptrans.gif" WIDTH="9" HEIGHT="1" BORDER="0" ALT="" /></TD>
		<TD COLSPAN="2" WIDTH="130" BGCOLOR="E7E5DF" ALIGN="LEFT"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(603) 'Descriptor: To find the style of your device, complete the following steps:%><BR /><BR /></FONT></TD>
		<TD WIDTH="9"   BGCOLOR="#E7E5DF"><IMG SRC="images/1ptrans.gif" WIDTH="9" HEIGHT="1" BORDER="0" ALT="" /></TD>
		<TD WIDTH="2"   BGCOLOR="#CCCCCC"><IMG SRC="images/1ptrans.gif" WIDTH="2" HEIGHT="1" BORDER="0" ALT="" /></TD>
		<TD WIDTH="9"><IMG SRC="images/1ptrans.gif" WIDTH="9" HEIGHT="1" BORDER="0" ALT="" /></TD>
	</TR>
	<!-- BEGIN:  Step 1 -->
	<TR>
		<TD WIDTH="18"  BGCOLOR="E7E5DF" ALIGN="LEFT" VALIGN="TOP" NOWRAP="1"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%If iAddressWizardStep = 1 Then%><B><%End If%>1.<%If iAddressWizardStep = 1 Then%></B><%End If%></FONT></TD>
		<TD WIDTH="112" BGCOLOR="E7E5DF" ALIGN="LEFT" VALIGN="TOP">
			<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
			<%If iAddressWizardStep = 1 Then%><B><%End If%>
			<%Response.Write asDescriptors(604) 'Descriptor: Select a category.%>
			<%If iAddressWizardStep = 1 Then%></B><%End If%>
			<BR />
			<%If iAddressWizardStep > 1 Then%>
			    <FONT COLOR="#0000cc"><% Response.Write sCategoryName %></FONT>
			<%End If%>
			<BR /><BR />
			</FONT>
		</TD>
		<TD WIDTH="9"   BGCOLOR="#E7E5DF" ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="images/<%If iAddressWizardStep = 1 Then%>wizardarrow_red_left<%Else%>wizardarrow_left<%End If%>.gif" WIDTH="9" HEIGHT="20" BORDER="0" ALT="" /></TD>
		<TD WIDTH="2"   BGCOLOR="#CCCCCC" VALIGN="TOP"><IMG SRC="images/<%If iAddressWizardStep = 1 Then%>wizardarrow_red_middle<%Else%>wizardarrow_middle<%End If%>.gif" WIDTH="2" HEIGHT="20" BORDER="0" ALT="" /></TD>
		<TD WIDTH="9"  ALIGN="LEFT" VALIGN="TOP"><IMG SRC="images/<%If iAddressWizardStep = 1 Then%>wizardarrow_red_right<%Else%>wizardarrow_right<%End If%>.gif" WIDTH="9" HEIGHT="20" BORDER="0" ALT="" /></TD>
	</TR>
	<!-- END:  Step 1 -->
	<!-- BEGIN:  Step 2 -->
	<TR>
		<TD WIDTH="18"  BGCOLOR="E7E5DF" ALIGN="LEFT" VALIGN="TOP" NOWRAP="1"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%If iAddressWizardStep = 2 Then%><B><%End If%>2.<%If iAddressWizardStep = 2 Then%></B><%End If%></FONT></TD>
		<TD WIDTH="112" BGCOLOR="E7E5DF" ALIGN="LEFT" VALIGN="TOP">
			<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
			<%If iAddressWizardStep = 2 Then%><B><%End If%>
			<%Response.Write asDescriptors(605) 'Descriptor: Select a style.%>
			<%If iAddressWizardStep = 2 Then%></B><%End If%>
			<BR />
			<%If iAddressWizardStep > 2 Then%>

			<%End If%>
			<BR /><BR />
			</FONT>
		</TD>
		<TD WIDTH="9"   BGCOLOR="#E7E5DF" ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="images/<%If iAddressWizardStep = 2 Then%>wizardarrow_red_left<%ElseIf iSubscribeWizardStep = 3 Then%>wizardarrow_left<%Else%>wizardarrow_grey_left<%End If%>.gif" WIDTH="9" HEIGHT="20" BORDER="0" ALT="" /></TD>
		<TD WIDTH="2"   BGCOLOR="#CCCCCC" VALIGN="TOP"><IMG SRC="images/<%If iAddressWizardStep = 2 Then%>wizardarrow_red_middle<%ElseIf iSubscribeWizardStep = 3 Then%>wizardarrow_middle<%Else%>wizardarrow_grey_middle<%End If%>.gif" WIDTH="2" HEIGHT="20" BORDER="0" ALT="" /></TD>
		<TD WIDTH="9"  ALIGN="LEFT" VALIGN="TOP"><IMG SRC="images/<%If iAddressWizardStep = 2 Then%>wizardarrow_red_right<%ElseIf iSubscribeWizardStep = 3 Then%>wizardarrow_right<%Else%>wizardarrow_grey_right<%End If%>.gif" WIDTH="9" HEIGHT="20" BORDER="0" ALT="" /></TD>
	</TR>
	<!-- END:  Step 2 -->
	<TR>
		<TD COLSPAN="6" BGCOLOR="#CCCCCC" HEIGHT="2"><IMG SRC="images/1ptrans.gif" HEIGHT="2" WIDTH="1" BORDER="0" ALT="" /></TD>
	</TR>
</TABLE>