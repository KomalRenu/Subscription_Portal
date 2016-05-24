<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
    'QUESTION: Should this be moved into a function?

    Dim oCacheDOM
    Dim oQuestions
    Dim oCurrentQuestion
    Dim sIconImageString

    Call LoadXMLDOMFromString(aConnectionInfo, sCacheXML, oCacheDOM)
    'Add error handling?

	If iSubscribeWizardStep > 2 Then
	    sServiceName = oCacheDOM.selectSingleNode("/mi/sub").getAttribute("svn")
	    sScheduleName = oCacheDOM.selectSingleNode("/mi/sub").getAttribute("scn")
	    sAddressName = oCacheDOM.selectSingleNode("/mi/sub").getAttribute("adn")
	End If
%>

<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<TR>
		<TD WIDTH="130" ALIGN="LEFT" BGCOLOR = "#f5f5f5"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="444444"><%Response.Write "<BR><B>" & asDescriptors(452) & "</B><BR><BR> " &asDescriptors(536) 'Descriptor: To sign up, complete the following steps:%><BR /><BR /></FONT></TD>
		<TD STYLE="background-image: url('images/bg_beige.gif'); background-repeat: repeat-y" WIDTH ="2" ROWSPAN="8"><IMG SRC="images/1ptrans.gif"></TD>
		<TD><IMG SRC="images/1ptrans.gif" BORDER="0" ALT="" /></TD>
	</TR>
	<TR>
		<TD HEIGHT="2"><IMG SRC="images/divider_left.gif" BORDER="0" ALT="" /></TD>
		<TD HEIGHT="2"><IMG SRC="images/1ptrans.gif" BORDER="0" ALT="" /></TD>
	</TR>
	<!-- BEGIN:  Step 1 -->
	<TR>
		<TD WIDTH="130"ALIGN="LEFT" VALIGN="TOP"  BGCOLOR = "#f5f5f5">
			<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="444444">
			<%If iSubscribeWizardStep = 1 Then%><B><%End If%>
			<%Response.Write "1." & asDescriptors(537) 'Descriptor: Select a Service.%>
			<%If iSubscribeWizardStep = 1 Then%></B><%End If%>
			<BR />
			<%If iSubscribeWizardStep > 1 Then%>
			    <FONT COLOR="#0000cc"><% Response.Write sServiceName %></FONT>
			    <BR />
			<%End If%>
			<BR />
			</FONT>
		</TD>
		<TD ALIGN="LEFT"><IMG SRC="images/<%If iSubscribeWizardStep = 1 Then%>indicator<%Else%>1ptrans<%End If%>.gif" BORDER="0" ALT="" /></TD>
	</TR>
	<!-- END:  Step 1 -->
	<TR>
		<TD><IMG SRC="images/divider_left.gif" BORDER="0" ALT="" /></TD>
		<TD><IMG SRC="images/1ptrans.gif" BORDER="0" ALT="" /></TD>
	</TR>
	<!-- BEGIN:  Step 2 -->
	<TR>
		<TD WIDTH="130" ALIGN="LEFT" VALIGN="TOP"  BGCOLOR = "#f5f5f5">
			<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="444444">
			<%If iSubscribeWizardStep = 2 Then%><B><%End If%>
			<%Response.Write "2." & asDescriptors(538) 'Descriptor: Specify a schedule and an address.%>
			<%If iSubscribeWizardStep = 2 Then%></B><%End If%>
			<BR />
			<%If iSubscribeWizardStep > 2 Then%>
			    <FONT COLOR="#0000cc"><%Response.Write sScheduleName%><BR />
			    <%Response.Write Replace(sAddressName, "@", "<BR />@")%></FONT>
			    <BR />
			<%End If%>
			<BR />
			</FONT>
		</TD>
		<TD ALIGN="LEFT"><IMG SRC="images/<%If iSubscribeWizardStep = 2 Then%>indicator<%Else%>1ptrans<%End If%>.gif" BORDER="0" ALT="" /></TD>
	</TR>
	<!-- END:  Step 2 -->
	<TR>
		<TD><IMG SRC="images/divider_left.gif" BORDER="0" ALT="" /></TD>
		<TD><IMG SRC="images/1ptrans.gif" BORDER="0" ALT="" /></TD>
	</TR>
	<!-- BEGIN:  Step 3 -->
	<TR>
		<TD WIDTH="130" ALIGN="LEFT" VALIGN="TOP" BGCOLOR = "#f5f5f5">
			<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="444444">
			<%If iSubscribeWizardStep = 3 Then%>
			    <B><%Response.Write "3." & asDescriptors(540) 'Descriptor: Personalize your service content.%></B>
			<%Else%>
                <%Response.Write "3." & asDescriptors(539) 'Descriptor: Personalize your Service content (if applicable).%>
			<%End If%>
			<BR />
			</FONT>

			<!-- BEGIN:  List of Questions -->
			<%If iSubscribeWizardStep = 3 Then%>
			<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="2">
                <%
		            Set oQuestions = oCacheDOM.selectNodes("/mi/qos/mi/in/oi")

		            If oQuestions.length > 0 Then
		                For Each oCurrentQuestion in oQuestions
		                    If Strcomp(oCurrentQuestion.getAttribute("hidden"), "0", vbTextCompare) = 0 Then
							    sIconImageString = ""
							    If Not (oCurrentQuestion.selectSingleNode("answer") Is Nothing) Then
							        sIconImageString = "qo_check_black"
							    Else
							        sIconImageString = "qo_bullet"
							    End If
							%>
							    <TR>
							    	<TD WIDTH="7" ALIGN="LEFT" VALIGN="TOP"><IMG SRC="images/<%Response.Write sIconImageString%>.gif" WIDTH="7" HEIGHT="15" BORDER="0" ALT="" /></TD>
							    	<TD>
							    		<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
							    		<%Response.Write oCurrentQuestion.getAttribute("n") %><BR />
							    		</FONT>
							    	</TD>
							    </TR>
							<%
							End If
		                Next
		            End If

		            Set oCacheDOM = Nothing
		            Set oQuestions = Nothing
		            Set oCurrentQuestion = Nothing
                %>
			</TABLE>
			<%End If%>
			<!-- END:  List of Questions -->

		</TD>
		<TD ALIGN="LEFT"><IMG SRC="images/<%If iSubscribeWizardStep = 3 Then%>indicator<%Else%>1ptrans<%End If%>.gif" BORDER="0" ALT="" /></TD>
	</TR>
	<!-- END:  Step 3 -->
	<TR>
		<TD><IMG SRC="images/divider_left.gif" BORDER="0" ALT="" /></TD>
		<TD><IMG SRC="images/1ptrans.gif" BORDER="0" ALT="" /></TD>
	</TR>
</TABLE>