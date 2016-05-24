<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Dim sGetISForSiteXML
Dim bHasProjs
Dim isErr
Dim lNumberOfIS
Const UD_PAGE_DESCRIPTOR = "Define User Details"
%>

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD>
			<img src="images/1ptrans.gif" WIDTH="2" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD VALIGN=TOP>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
				<TR>
					<TD>
						<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><b><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(286)'Descriptor: Preferences%><%Else%><%Response.Write UCase(asDescriptors(286))'Descriptor: Preferences%><%End If%></b></font>
					</TD>
				</TR>
				<TR>
					<TD>
						<!-- QUESTION: Should these options be rendered by a function? -->
						<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
							<TR>
								<TD WIDTH="1%" VALIGN="TOP"><img src="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
								<TD WIDTH="99%"><%If StrComp(sOptSection, "1", vbBinaryCompare) = 0 Or Len(sOptSection) = 0 Then%><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><b><%Response.Write asDescriptors(402) 'Descriptor: User options%></b></font><%Else%><A HREF="options.asp?optSection=1"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(402) 'Descriptor: User options%></font></A><%End If%></TD>
							</TR>

							<% If Session("CastorUserID") = "" Then %>
								<TR>
									<TD WIDTH="1%" VALIGN="TOP"><img src="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
									<TD WIDTH="99%"><%If StrComp(sOptSection, "2", vbBinaryCompare) = 0 Then%><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><b><%Response.Write asDescriptors(143) 'Descriptor: Change my password%></b></font><%Else%><A HREF="change_password.asp"><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#000000" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(143) 'Descriptor: Change my password%></font></A><%End If%></TD>
								</TR>
							<% End If %>

							<TR>
								<TD WIDTH="1%" VALIGN="TOP"><img src="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
								<TD WIDTH="99%"><%If StrComp(sOptSection, "5", vbBinaryCompare) = 0 Then%><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><b><%Response.Write asDescriptors(971) 'Descriptor: Define User Details%></b></font><%Else%><A HREF="ProcessUserDetails.asp?getUD=1"><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#000000" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(971) 'Descriptor: Define User Details%></font></A><%End If%></TD>
							</TR>

							<!-- Do not show this link if all the projects for the site have "Project Credentials"-->
						    <%
								isErr = cu_GetInformationSourcesForSite(sGetISForSiteXML, bHasProjs)
								Call GetNumberOfInformationSourceThatNeedAuthentication(sGetISForSiteXML, lNumberOfIS)
								If (bHasProjs and lNumberOfIS > 0) Then
						    %>
									<TR>
										<TD WIDTH="1%" VALIGN="TOP"><img src="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
										<TD WIDTH="99%"><%If StrComp(sOptSection, "4", vbBinaryCompare) = 0 Then%><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><b><%Response.Write asDescriptors(466) 'Descriptor: Change Information Source credentials%></b></font><%Else%><A HREF="authentications.asp"><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#000000" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(466) 'Descriptor: Change Information Source credentials%></font></A><%End If%></TD>
									</TR>
							<% End If %>

							<% If Session("CastorUserID") = "" Then %>
								<TR>
									<TD WIDTH="1%" VALIGN="TOP"><img src="images/bullet.gif" WIDTH="3" HEIGHT="8" BORDER="0" ALT="" /></TD>
									<TD WIDTH="99%"><%If StrComp(sOptSection, "3", vbBinaryCompare) = 0 Then%><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><b><%Response.Write asDescriptors(460) 'Descriptor: Deactivate my account%></b></font><%Else%><A HREF="deactivate.asp"><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#000000" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(460) 'Descriptor: Deactivate my account%></font></A><%End If%></TD>
								</TR>
							<% End If %>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>