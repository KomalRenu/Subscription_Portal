<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
    On Error Resume Next
    Dim oDeviceTypesDOM
    Dim oDeviceTypes
    Dim oCurrentDeviceType

    Call LoadXMLDOMFromString(aConnectionInfo, sDeviceTypesXML, oDeviceTypesDOM)
    'Add error handling?
    Set oDeviceTypes = oDeviceTypesDOM.selectNodes("//devicetype")
%>
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="160">
	<TR>
		<TD>
			<IMG SRC="images/1ptrans.gif" WIDTH="2" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD VALIGN="TOP">
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
				<TR>
					<TD>
						<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(374)'Descriptor: Create new address%><%Else%><%Response.Write UCase(asDescriptors(374))'Descriptor: Create new address%><%End If%></B></FONT>
					</TD>
				</TR>
				<TR>
				    <TD>
				        <%If oDeviceTypes.length > 0 Then%>
				        <TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0">
				            <%For Each oCurrentDeviceType in oDeviceTypes%>
				            <TR>
				                <TD><A HREF="addresses.asp?action=new&devicetypeID=<%Response.Write oCurrentDeviceType.selectSingleNode("devicetypeID").text%>"><IMG SRC="<%Response.Write oCurrentDeviceType.selectSingleNode("icon").text%>" WIDTH="<%Response.Write oCurrentDeviceType.selectSingleNode("icon").getAttribute("width")%>" HEIGHT="<%Response.Write oCurrentDeviceType.selectSingleNode("icon").getAttribute("height")%>" BORDER="0" ALT="<%Response.Write oCurrentDeviceType.selectSingleNode("name").text%>" /></A></TD>
				                <%If StrComp(CStr(oCurrentDeviceType.selectSingleNode("devicetypeID").text), sDeviceTypeID, vbBinaryCompare) = 0 Then%>
				                    <TD><B><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write oCurrentDeviceType.selectSingleNode("name").text%></FONT></B></TD>
				                <%Else%>
				                    <TD><A HREF="addresses.asp?action=new&devicetypeID=<%Response.Write oCurrentDeviceType.selectSingleNode("devicetypeID").text%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" COLOR="#000000" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write oCurrentDeviceType.selectSingleNode("name").text%></FONT></A></TD>
				                <%End If%>
				            </TR>
				            <%Next%>
				        </TABLE>
				        <%End If%>
				    </TD>
				</TR>
			</TABLE>
			<%
			    If Len(sDeviceTypeID) > 0 Then

			        Set oCurrentDeviceType = oDeviceTypesDOM.selectSingleNode("/devicetypes/devicetype[devicetypeID = '" & sDeviceTypeID & "']")
			%>
			<BR />
			<TABLE BORDER="0" CELLPADDING="3" CELLSPACING="0" WIDTH="100%">
			    <TR>
			        <TD>
			            <TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0" WIDTH="140">
			                <TR>
			                    <TD BGCOLOR="#cccccc">
			                        <TABLE BORDER="0" CELLPADDING="3" CELLSPACING="0" WIDTH="100%">
			                            <TR>
			                                <TD>
			                                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>"><FONT SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
			                                        <B><%Response.Write asDescriptors(530) 'Descriptor: Help with:%></B><BR />
			                                        <B><FONT COLOR="#0000cc"><%Response.Write CStr(oCurrentDeviceType.selectSingleNode("name").text)%></FONT></B>
			                                    </FONT></FONT>
			                                </TD>
			                            </TR>
			                            <TR>
			                                <TD BGCOLOR="#ffffff">
			                                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
			                                        <%
			                                            Select Case oCurrentDeviceType.selectSingleNode("editfields/field[@col = 'v']").getAttribute("vld")
                                                            Case "number"
                                                                Response.Write asDescriptors(614) 'Descriptor: Please enter a value for the address in the following form: any numbers and the following characters - ( )
                                                            Case "email"
                                                                Response.Write asDescriptors(419) 'Descriptor: Please enter an address in the form of: user@server.com
                                                            Case "none"
                                                                Response.Write asDescriptors(635) 'Descriptor: Please enter an address in the following form: any text or numeric characters
                                                            Case Else
			                                            End Select
			                                        %>
			                                    </FONT>
			                                </TD>
			                            </TR>
			                        </TABLE>
			                    </TD>
			                </TR>
			            </TABLE>
			        </TD>
			    </TR>
			</TABLE>
			<%End If%>
		</TD>
	</TR>
</TABLE>
<%
    Set oDeviceTypesDOM = Nothing
    Set oDeviceTypes = Nothing
    Set oCurrentDeviceType = Nothing
%>