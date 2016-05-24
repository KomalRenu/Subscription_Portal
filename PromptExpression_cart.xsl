<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<!-- Copyright 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. -->
   <xsl:template language="JAVASCRIPT" match=".">
	<xsl:apply-templates select="./pif[@pt='8']" />
   </xsl:template>

   <xsl:template match='pif[@pt="8"]'>
	<TABLE WIDTH="99%" BORDER="0" CELLSPACING="0" CELLPADDING="1">
		<!-- BEGIN:  Prompt Title Bar -->
		<TR>
			<xsl:if test="./info/@digit1" >
				<TD VALIGN="TOP" ROWSPAN="3">
					<IMG WIDTH="14" HEIGHT="22" ALT="" BORDER="0">
						<xsl:attribute name="SRC">Images/<xsl:value-of select="./info/@digit1" />_olive.gif</xsl:attribute>
					</IMG>
				</TD>
			</xsl:if>
			<xsl:if test="./info/@digit2">
				<TD VALIGN="TOP" ROWSPAN="3">
					<IMG WIDTH="14" HEIGHT="22" ALT="" BORDER="0">
						<xsl:attribute name="SRC">Images/<xsl:value-of select="./info/@digit2" />_olive.gif</xsl:attribute>
					</IMG>
				</TD>	
			</xsl:if>
			<TD VALIGN="TOP" ROWSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="4" HEIGHT="1" ALT="" BORDER="0" /></TD>
							
			<TD BGCOLOR="#DDDDBB" WIDTH="100%" ALIGN="LEFT">
				<A><xsl:attribute name="NAME"><xsl:value-of select="./info/@order" /></xsl:attribute></A>
				<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/mediumFont").text</xsl:eval></xsl:attribute><xsl:value-of select="@ttl" />  <FONT COLOR="#cc0000"><xsl:value-of select="./info/@step" /></FONT></FONT>
			</TD>
			
			<xsl:choose>
			<xsl:when test="./info[@totop = '1']">
				<TD BGCOLOR="#DDDDBB" ALIGN="RIGHT" VALIGN="TOP">
					<A HREF="#top">
						<IMG SRC="Images/BackToTop.gif" WIDTH="20" HEIGHT="13" BORDER="0" >
							<xsl:attribute name="ALT"><xsl:eval>this.selectSingleNode("/mi/inputs/Desc_330").text</xsl:eval></xsl:attribute>
						</IMG>
					</A>
				</TD>
			</xsl:when>
			<xsl:otherwise>
				<TD BGCOLOR="#DDDDBB" ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
			</xsl:otherwise>
			</xsl:choose>
		</TR>
		
		<!-- BEGIN:  Prompt Description & Content -->
		<TR>
			<TD WIDTH="100%" ALIGN="LEFT" COLSPAN="2">
					<TABLE COLS="2" CELLSPACING="0" BORDER="0" WIDTH="100%">
						<xsl:if expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
						<TR>
						<TD>
							<FONT COLOR="#FFFFFF"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
							<B>
							<DIV>
								<xsl:attribute name="NAME">DetailErrorDisplay_<xsl:value-of select='./@pin' /></xsl:attribute>
								<xsl:attribute name="ID">DetailErrorDisplay_<xsl:value-of select='./@pin' /></xsl:attribute>
								<IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" />
							</DIV>
							</B>
							</FONT>
						</TD>
						</TR>
						</xsl:if>

						<xsl:if test="./info[@error != '']">
							<TR>
								<TD VALIGN="TOP" ALIGN="LEFT" WIDTH="23" ROWSPAN="2"><IMG SRC="Images/promptError_white.gif" WIDTH="23" HEIGHT="23" BORDER="0"><xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_981").text</xsl:eval></xsl:attribute></IMG></TD>
								<TD VALIGN="TOP" ALIGN="LEFT">
									<FONT COLOR="#CC0000"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><B><xsl:value-of select="./info/@error" /></B><BR /></FONT>
								</TD>
							</TR>
						</xsl:if>

						<TR>
						<TD>
							<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><B><xsl:value-of select="@mn" /><BR /><xsl:value-of select="./info/@msg" /></B><BR /></FONT>
						</TD>
						</TR>
					</TABLE>
			</TD>
		</TR>
		<TR>
			<TD COLSPAN="2">
			<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="3">
				<TR>
					<TD WIDTH="80%">

					<!-- Begin cart -->	
					<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
					  <TR>
						<TD ALIGN="LEFT" VALIGN="TOP">
							<!-- attribute or metric : -->
							<xsl:choose>
							   	<xsl:when test="./res[. = '10']" >
									<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_517").text</xsl:eval>:</FONT><BR />
								</xsl:when>
								<xsl:when test="./res[. = '17']" >
									<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_518").text</xsl:eval>:</FONT><BR />
								</xsl:when>
								<xsl:when test="./res[. = '18']" >
									<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_518").text</xsl:eval>:</FONT><BR />
								</xsl:when>
							</xsl:choose>	

							<!-- available -->
							<SELECT SIZE="10">
								<xsl:attribute name="NAME">Available_<xsl:value-of select='@pin' /></xsl:attribute>
								<!--for the calendar button-->
								<xsl:choose>
											<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
												<xsl:attribute name="ID">Available_<xsl:value-of select='@pin' /></xsl:attribute>
												<xsl:attribute name="onChange">showOrHideCalendarButton('Available_<xsl:value-of select='@pin' />','Calendar_button_<xsl:value-of select='@pin' />');</xsl:attribute>
											</xsl:when>
								</xsl:choose>
								<xsl:choose>
									<xsl:when test="./pa[@il='1' $or$ @idl='1']/mi[@pcc!='0']" >
								     	<xsl:choose>
								     	<xsl:when test="./res[. = '10']" >
							     			<xsl:for-each select="./pa[@il='1' $or$ @idl='1']/mi/oi" >	
												<OPTION>
												<xsl:attribute name="VALUE"><xsl:value-of select='@did' /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_n" /></xsl:attribute>
												<xsl:attribute name="DATATYPE"><xsl:value-of select='@ddt' /> </xsl:attribute>
												<xsl:if test=".[@highlight='1']">
													<xsl:if test="/mi/inputs/accessibilityModeOff">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
												</xsl:if>
												<xsl:value-of select="@disp_n" />
												
												</OPTION>
											</xsl:for-each>
										</xsl:when>
										<xsl:when test="./res[. ='17' $or$ . ='18']" >
											<xsl:for-each select="./pa[@il='1' $or$ @idl='1']/mi/oi" >
												<xsl:for-each select="oi">	
													<OPTION>
													<xsl:attribute name="VALUE">
													<xsl:value-of select="context(-2)/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="context(-2)/@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@ddt" /></xsl:attribute>
													<xsl:attribute name="DATATYPE"><xsl:value-of select='@ddt' /> </xsl:attribute>
													<xsl:if test=".[@highlight='1']">
														<xsl:if test="/mi/inputs/accessibilityModeOff">
															<xsl:attribute name="SELECTED">1</xsl:attribute>
														</xsl:if>
													</xsl:if>
													<xsl:value-of select="context(-2)/@disp_n" />(<xsl:value-of select="context()/@disp_n" />)
													
													</OPTION>
												</xsl:for-each>
											</xsl:for-each>
										</xsl:when>
										</xsl:choose>
									</xsl:when>
									<xsl:otherwise>
										<OPTION VALUE="--none--">--- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ---</OPTION>
									</xsl:otherwise>
								</xsl:choose>
							</SELECT><BR />
							<!-- incremental fetch links -->
							<xsl:if test="./increfetch/*" >
					 			<xsl:apply-templates select="./increfetch" />
					 		</xsl:if>	
						</TD>

						<!-- operator -->
						<TD ALIGN="LEFT" valign="top">
							<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
								<TR>
									<TD ALIGN="LEFT">
										<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_528").text</xsl:eval>:</FONT><BR />
										<SELECT SIZE="1">
										<xsl:attribute name="NAME">Operator_<xsl:value-of select='@pin' /></xsl:attribute>
										<!--for calendar control - to append or not -->
										<xsl:choose>
											<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
												<xsl:attribute name="onChange">updateOperator('Operator_<xsl:value-of select='@pin' />');</xsl:attribute>
												<xsl:attribute name="ID">Operator_<xsl:value-of select='@pin' /></xsl:attribute>
											</xsl:when>
										</xsl:choose>
											
											
										<!-- different operators for AQ/MQ -->
										<xsl:choose>
										   	<xsl:when test="./res[. = '10']" >
												<OPTION VALUE="M17">
													<xsl:if test="current[@op='M17']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_519").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M44">
													<xsl:if test="current[@op='M44']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_614").text</xsl:eval>
												</OPTION>											
												<OPTION VALUE="M6">
													<xsl:if test="current[@op='M6']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_520").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M7">
													<xsl:if test="current[@op='M7']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_612").text</xsl:eval>
												</OPTION>											
 												<OPTION VALUE="M8">
 													<xsl:if test="current[@op='M8']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_521").text</xsl:eval>
													
 												</OPTION>
												<OPTION VALUE="M10">
													<xsl:if test="current[@op='M10']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_522").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M9">
													<xsl:if test="current[@op='M9']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_523").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M11">
													<xsl:if test="current[@op='M11']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_524").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="R1">
													<xsl:if test="current[@op='R1']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_529").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="R2">
													<xsl:if test="current[@op='R2']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_530").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="P1">
													<xsl:if test="current[@op='P1']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_529").text</xsl:eval>(%)
												</OPTION>
												<OPTION VALUE="P2">
													<xsl:if test="current[@op='P2']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_530").text</xsl:eval>(%)
												</OPTION>
												<OPTION VALUE="M22">
													<xsl:if test="current[@op='M22']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_898").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M57">
													<xsl:if test="current[@op='M57']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_2394").text</xsl:eval>
												</OPTION>
											</xsl:when>
											<xsl:when test="./res[. = '17' $or$ .='18']" >
												<OPTION VALUE="M17">
													<xsl:if test="current[@op='M17']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_519").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M44">
													<xsl:if test="current[@op='M44']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_614").text</xsl:eval>
												</OPTION>											
												<OPTION VALUE="M6">
													<xsl:if test="current[@op='M6']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_520").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M7">
													<xsl:if test="current[@op='M7']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_612").text</xsl:eval>
												</OPTION>											
 												<OPTION VALUE="M8">
 													<xsl:if test="current[@op='M8']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
 													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_521").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M10">
													<xsl:if test="current[@op='M10']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_522").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M9">
													<xsl:if test="current[@op='M9']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_523").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M11">
													<xsl:if test="current[@op='M11']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_524").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M18">
													<xsl:if test="current[@op='M18']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_525").text</xsl:eval>
													<xsl:if test="available[@datetime='1']"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_611").text</xsl:eval></xsl:if>
												</OPTION>
												<OPTION VALUE="M43">
													<xsl:if test="current[@op='M43']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_526").text</xsl:eval>
													<xsl:if test="available[@datetime='1']"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_611").text</xsl:eval></xsl:if>
												</OPTION>
												<OPTION VALUE="M22">
													<xsl:if test="current[@op='M22']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_898").text</xsl:eval>
												</OPTION>
												<OPTION VALUE="M57">
													<xsl:if test="current[@op='M57']">
														<xsl:attribute name="SELECTED">1</xsl:attribute>
													</xsl:if>
													<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_2394").text</xsl:eval>
												</OPTION>
											</xsl:when>
					     				</xsl:choose>
					     						
										</SELECT>
										<BR />
										<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_527").text</xsl:eval>:</FONT>
										<BR />
										<xsl:choose>
											<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
												<INPUT TYPE="TEXT" SIZE="23" STYLE="font-family: courier">
												<xsl:attribute name="NAME">Input_<xsl:value-of select='@pin' /></xsl:attribute>
												<xsl:attribute name="ID">Input_<xsl:value-of select='@pin' /></xsl:attribute>
												</INPUT>
												<!--code for calendar control-->
												<SCRIPT language="JavaScript">
													var sDateFormat = "<xsl:eval>this.selectSingleNode("/mi/inputs/DATE_FORMAT").text;</xsl:eval>"
												</SCRIPT>
												<IMG ALIGN="top" SRC="Images/calendar.gif" STYLE="cursor: hand;" >
												<xsl:attribute name="ID">Calendar_button_<xsl:value-of select='@pin' /></xsl:attribute>
												<xsl:attribute name="onClick">showCalendar(getMonth('Input_<xsl:value-of select='@pin' />'),getYear('Input_<xsl:value-of select='@pin' />'),'Input_<xsl:value-of select='@pin' />',getObjSumLeft('Calendar_button_<xsl:value-of select='@pin' />')-108, (getObjSumTop('Calendar_button_<xsl:value-of select='@pin' />')+getObjHeight('Calendar_button_<xsl:value-of select='@pin' />')));</xsl:attribute>
												</IMG>
												
												<SCRIPT language="JavaScript">
													createDivForCalendar();
													showOrHideCalendarButton('Available_<xsl:value-of select='@pin' />', 'Calendar_button_<xsl:value-of select='@pin' />');
													updateOperator('Operator_<xsl:value-of select='@pin' />');
												</SCRIPT>

											</xsl:when>
											<xsl:otherwise>
												<INPUT TYPE="TEXT" SIZE="23" STYLE="font-family: courier">
												<xsl:attribute name="NAME">Input_<xsl:value-of select='@pin' /></xsl:attribute>
												</INPUT>											
											</xsl:otherwise>
										</xsl:choose>							
									</TD>
								</TR>
								<TR>
					 				<TD ALIGN="CENTER">
									<xsl:choose>
									<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
										<A>
										<xsl:attribute name="HREF">javascript:AddItemsbyListObjectForExpression(document.PromptForm.Available_<xsl:value-of select="./@pin" />, document.PromptForm.Operator_<xsl:value-of select="./@pin" />, document.PromptForm.Input_<xsl:value-of select="./@pin" />, document.PromptForm.Selected_<xsl:value-of select="./@pin" />, '<xsl:value-of select="./@pin" />')</xsl:attribute>
											<IMG SRC="Images/btn_add.gif" WIDTH="25" HEIGHT="25" BORDER="0" >
												<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_537").text</xsl:eval></xsl:attribute> 
												<xsl:attribute name="NAME">Add_<xsl:value-of select="./@pin" /></xsl:attribute> 
											</IMG
										></A>				
										<BR />
										<A>
										<xsl:attribute name="HREF">javascript:RemoveItemsbyListObjectForExpression(document.PromptForm.Selected_<xsl:value-of select="./@pin" />, document.PromptForm.Operator_<xsl:value-of select="./@pin" />, document.PromptForm.Available_<xsl:value-of select="./@pin" />, '<xsl:value-of select="./@pin" />')</xsl:attribute>
											<IMG SRC="Images/btn_remove.gif" WIDTH="25" HEIGHT="25" BORDER="0">					
												<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_875").text</xsl:eval></xsl:attribute> 
												<xsl:attribute name="NAME">Remove_<xsl:value-of select="./@pin" /></xsl:attribute> 
											</IMG
										></A>				
									</xsl:when>
									<xsl:otherwise>	
					 					<INPUT TYPE="IMAGE" SRC="Images/btn_add.gif" WIDTH="25" HEIGHT="25" BORDER="0">
											<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_537").text</xsl:eval></xsl:attribute>
											<xsl:attribute name="NAME">Add_<xsl:value-of select="./@pin" /></xsl:attribute> 
										</INPUT>
										<BR />
										<INPUT TYPE="IMAGE" SRC="Images/btn_remove.gif" WIDTH="25" HEIGHT="25" BORDER="0">
											<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_875").text</xsl:eval></xsl:attribute>
											<xsl:attribute name="NAME">Remove_<xsl:value-of select="./@pin" /></xsl:attribute> 
										</INPUT>
									</xsl:otherwise>
									</xsl:choose>
					 				</TD>
								</TR>
							</TABLE>
						</TD>

						<!-- selections -->
						<TD ALIGN="LEFT" VALIGN="TOP">
							<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_532").text</xsl:eval>:</FONT><BR />
							<SELECT SIZE="10" MULTIPLE="1" >
								<xsl:attribute name="NAME">Selected_<xsl:value-of select='@pin' /></xsl:attribute>
								<xsl:if expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
									<xsl:attribute name="WIDTH">200</xsl:attribute>
								</xsl:if>
									<xsl:choose>
									<xsl:when test="./pa[@ia='1']/exp/unknowndef" >	
										<OPTION VALUE="-default-"><xsl:value-of select="./pa[@ia='1']/exp/unknowndef/@text" /></OPTION>
									</xsl:when>
									<xsl:when test="./pa[@ia='1']/exp/nd/nd" >	
										<xsl:choose>
										<xsl:when test="./res[. = '10']" >									
											<xsl:for-each select="./pa[@ia='1']/exp/nd/nd" >	
											<OPTION>
											<xsl:choose>
											<xsl:when test=".[@optp = '2']" >
												<xsl:attribute name="VALUE"><xsl:value-of select="./nd[0]/oi[@tp='4']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/@disp_n" /><xsl:eval no-entities="1">ESC</xsl:eval>R<xsl:value-of select="./op/@fnt" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./nd[1]/cst" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="context()/@res" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
											</xsl:when>
											<xsl:when test=".[@optp = '3']" >
												<xsl:attribute name="VALUE"><xsl:value-of select="./nd[0]/oi[@tp='4']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/@disp_n" /><xsl:eval no-entities="1">ESC</xsl:eval>P<xsl:value-of select="./op/@fnt" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./nd[1]/cst" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="context()/@res" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
											</xsl:when>
											<xsl:otherwise>
												<xsl:choose>
												<xsl:when test="./op[@fnt='17' $or$ @fnt='44']" >
													<xsl:attribute name="VALUE"><xsl:value-of select="./nd[0]/oi[@tp='4']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/@disp_n" /><xsl:eval no-entities="1">ESC</xsl:eval>M<xsl:value-of select="./op/@fnt" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./nd[1]/cst" />;<xsl:value-of select="./nd[2]/cst" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="context()/@res" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
												</xsl:when>
												<xsl:when test="./op[@fnt='22' $or$ @fnt='57']" >
													<xsl:attribute name="VALUE"><xsl:value-of select="./nd[0]/oi[@tp='4']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/@disp_n" /><xsl:eval no-entities="1">ESC</xsl:eval>M<xsl:value-of select="./op/@fnt" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./op/@cstvalue" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="context()/@res" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
												</xsl:when>
												<xsl:otherwise>
													<xsl:attribute name="VALUE"><xsl:value-of select="./nd[0]/oi[@tp='4']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/@disp_n" /><xsl:eval no-entities="1">ESC</xsl:eval>M<xsl:value-of select="./op/@fnt" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./nd[1]/cst" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="context()/@res" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
												</xsl:otherwise>
												</xsl:choose>	
											</xsl:otherwise>
											</xsl:choose>
											<xsl:value-of select="./@disp_n" />
											</OPTION>
											</xsl:for-each>											
										</xsl:when>
										<xsl:when test="./res [.='17' $or$ .='18']" >	
											<xsl:for-each select="./pa[@ia='1']/exp/nd/nd" >	
											<OPTION>
												<!-- no need to differentiate Between operator -->
												<!--
												<xsl:choose>
												<xsl:when test="./op[@fnt='17' $or$ @fnt='44']" > 
													<xsl:attribute name="VALUE"><xsl:value-of select="./nd[0]/oi[@tp='12']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/oi[@tp='21']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[1]/@ddt" /><xsl:eval no-entities="1">ESC</xsl:eval>M<xsl:value-of select="./op/@fnt" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./nd[1]" />;<xsl:value-of select="./nd[2]" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="context(-2)/res" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
													<xsl:value-of select="./@disp_n" />
												</xsl:when>
												<xsl:otherwise>
												-->
													<!-- if time node, use <nd>'s ddt and text -->
													<!-- else, use <cst>'s ddt and text -->
													<xsl:choose>
														<xsl:when test="./nd[1]/cst">
															<xsl:attribute name="VALUE"><xsl:value-of select="./nd[0]/oi[@tp='12']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/oi[@tp='21']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[1]/cst/@ddt" /><xsl:eval no-entities="1">ESC</xsl:eval>M<xsl:value-of select="./op/@fnt" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:for-each select="./nd[index()>0]"><xsl:value-of select="./cst" /><xsl:if match="(./nd)[$not$ end()]">;</xsl:if></xsl:for-each><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="context(-2)/res" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
															<xsl:value-of select="./@disp_n" />
														</xsl:when>
														<xsl:otherwise>
															<xsl:attribute name="VALUE"><xsl:value-of select="./nd[0]/oi[@tp='12']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/oi[@tp='21']/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[1]/@ddt" /><xsl:eval no-entities="1">ESC</xsl:eval>M<xsl:value-of select="./op/@fnt" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:for-each select="./nd[index()>0]"><xsl:value-of select="." /><xsl:if match="(./nd)[$not$ end()]">;</xsl:if></xsl:for-each><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="context(-2)/res" /><xsl:eval no-entities="1">ESC</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
															<xsl:value-of select="./@disp_n" />
														</xsl:otherwise>
													</xsl:choose>
												<!--
												</xsl:otherwise>
												</xsl:choose>
												-->
											</OPTION>
											</xsl:for-each>												
										</xsl:when>
										</xsl:choose>									
									</xsl:when>
									<xsl:otherwise>
										<xsl:choose>
											<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
												<OPTION SELECTED="1" VALUE="-none-">--- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ---</OPTION>
											</xsl:when>
											<xsl:otherwise>
												<OPTION VALUE="-none-">--- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ---</OPTION>
											</xsl:otherwise>	
										</xsl:choose>
									</xsl:otherwise>
									</xsl:choose>
							</SELECT>
						</TD>

						<!-- AND/OR all subexpressions -->
						<xsl:choose>
						<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
							<TD ALIGN="LEFT" VALIGN="TOP" NOWRAP="1">
								<DIV>
									<xsl:attribute name="ID">ANDOR_<xsl:value-of select='@pin' /></xsl:attribute>
									<xsl:attribute name="NAME">ANDOR_<xsl:value-of select='@pin' /></xsl:attribute>
									<xsl:choose>
									<xsl:when test="./pa[@ia='1']/exp[./nd/nd[1] $and$ $not$ unknowndef]">
										<xsl:attribute name="STYLE">display:block;</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<xsl:attribute name="STYLE">display:none;</xsl:attribute>
									</xsl:otherwise>
									</xsl:choose>				
									<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
										<CENTER><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_533").text</xsl:eval>:<BR /></CENTER>
										<INPUT TYPE="RADIO" VALUE="AND">
										<xsl:attribute name="NAME">FilterOperator_<xsl:value-of select='@pin' /></xsl:attribute>
										<xsl:if test="./pa[@ia='1']/exp/nd/op[@fnt='19']">
											<xsl:attribute name="CHECKED">1</xsl:attribute>
										</xsl:if>
										<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_534").text</xsl:eval><BR />
										</INPUT>
										<INPUT TYPE="RADIO" VALUE="OR">
										<xsl:attribute name="NAME">FilterOperator_<xsl:value-of select='@pin' /></xsl:attribute>
										<xsl:if test="./pa[@ia='1']/exp/nd/op[@fnt='20']">
											<xsl:attribute name="CHECKED">1</xsl:attribute>
										</xsl:if>
										<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_535").text</xsl:eval><BR />
										</INPUT>
									</FONT>
									<IMG SRC="Images/1ptrans.gif" WIDTH="130" HEIGHT="1" ALT="" BORDER="0" />
								</DIV>
							</TD>
						</xsl:when>
						<xsl:otherwise>					
							<xsl:if test="./pa[@ia='1']/exp[./nd/nd[1] $and$ $not$ unknowndef]">
								<TD ALIGN="LEFT" VALIGN="TOP" NOWRAP="1">
									<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
										<CENTER><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_533").text</xsl:eval>:<BR /></CENTER>
										<INPUT TYPE="RADIO" VALUE="AND">
										<xsl:attribute name="NAME">FilterOperator_<xsl:value-of select='@pin' /></xsl:attribute>
										<xsl:if test="./pa[@ia='1']/exp/nd/op[@fnt='19']">
											<xsl:attribute name="CHECKED">1</xsl:attribute>
										</xsl:if>
										<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_534").text</xsl:eval><BR />
										</INPUT>
										<INPUT TYPE="RADIO" VALUE="OR">
										<xsl:attribute name="NAME">FilterOperator_<xsl:value-of select='@pin' /></xsl:attribute>
										<xsl:if test="./pa[@ia='1']/exp/nd/op[@fnt='20']">
											<xsl:attribute name="CHECKED">1</xsl:attribute>
										</xsl:if>
										<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_535").text</xsl:eval><BR />
										</INPUT>
									</FONT>
									<IMG SRC="Images/1ptrans.gif" WIDTH="130" HEIGHT="1" ALT="" BORDER="0" />
								</TD>
							</xsl:if>
						</xsl:otherwise>
						</xsl:choose>
					  </TR>
					</TABLE>

					<INPUT TYPE="HIDDEN">
						<xsl:attribute name="NAME">BBcurr_<xsl:value-of select="./@pin" /></xsl:attribute> 
						<xsl:attribute name="VALUE"><xsl:value-of select="./increfetch/curr/@start" /></xsl:attribute>
					</INPUT>
					<!-- End Cart -->

				</TD>
				</TR>
			</TABLE>


			</TD>
		</TR>
	</TABLE>
   </xsl:template>
   
   <xsl:template match="increfetch">
	<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
		<!-- previous -->	
		<xsl:if test="./prev[@count $ne$ '']" >
			<INPUT TYPE="IMAGE" SRC="Images/arrow_left_inc_fetch.gif" WIDTH="5" HEIGHT="10" BORDER="0">
				<xsl:attribute name="ALT"><xsl:value-of select="./prev/@title" /></xsl:attribute>
				<xsl:attribute name="NAME">prev_<xsl:value-of select="./@pin" /></xsl:attribute>
			</INPUT> 
			<INPUT TYPE="HIDDEN">
				<xsl:attribute name="NAME">BBprev_<xsl:value-of select="./@pin" /></xsl:attribute>
				<xsl:attribute name="VALUE"><xsl:value-of select="./prev/@link" /></xsl:attribute>
			</INPUT> 
		</xsl:if>

		<!-- current -->
		<xsl:value-of select="./curr/@title" />

		<!-- next -->
		<xsl:if test="./next[@count $ne$ '']" >
			<INPUT TYPE="IMAGE" SRC="Images/arrow_right_inc_fetch.gif" WIDTH="5" HEIGHT="10" BORDER="0">
				<xsl:attribute name="ALT"><xsl:value-of select="./next/@title" /></xsl:attribute>
				<xsl:attribute name="NAME">next_<xsl:value-of select="./@pin" /></xsl:attribute>
			</INPUT>
			<INPUT TYPE="HIDDEN">
				<xsl:attribute name="NAME">BBnext_<xsl:value-of select="./@pin" /></xsl:attribute>
				<xsl:attribute name="VALUE"><xsl:value-of select="./next/@link" /></xsl:attribute>
			</INPUT>
		</xsl:if>
	</FONT>
  </xsl:template>
  
<xsl:script><![CDATA[
 var ESC="&#027;";
 var RS="&#030;";
]]></xsl:script>

</xsl:stylesheet>
