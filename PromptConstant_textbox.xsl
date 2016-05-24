<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<!-- Copyright 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. -->
  <xsl:template match=".">
	<xsl:apply-templates select="./pif" />
  </xsl:template>

  <xsl:template match='pif'>
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
					
						<xsl:choose>
						 <!-- BEGIN:  Display calendar component if prompt type is Date and DHTML is active -->
						 <xsl:when test=".[@pt='5' and /mi/inputs/DHTML[text()='1']]">
						
							<SCRIPT language="JavaScript">
								sDateFormat = "<xsl:eval>this.selectSingleNode("/mi/inputs/DATE_FORMAT").text;</xsl:eval>"
							</SCRIPT>
						
							<!-- INSERT DHTML CODE COMPONENT HERE -->
							<!-- BEGIN: constant-->
						
							<INPUT TYPE="TEXT" SIZE="15" >
								<xsl:attribute name="NAME">Available_<xsl:value-of select='@pin' /></xsl:attribute>
								<xsl:attribute name="ID">Available_<xsl:value-of select='@pin' /></xsl:attribute>
								<xsl:choose>
									<xsl:when test="./pa[@ia='1']/tp/stt">
										<xsl:attribute name="VALUE"><xsl:value-of select="./pa[@ia='1']/tp/stt/@stv" /></xsl:attribute>										
									</xsl:when>	
									<xsl:when test="./pa[@ia='1']/tp/dt">
										<xsl:attribute name="VALUE"><xsl:value-of select="./pa[@ia='1']/tp/dt" /></xsl:attribute>										
									</xsl:when>	
									<xsl:otherwise>
										<xsl:attribute name="VALUE"><xsl:value-of select="./pa[@ia='1']" /></xsl:attribute>										
									</xsl:otherwise>
								</xsl:choose>
							</INPUT>
							<IMG ALIGN="top" SRC="Images/calendar.gif" STYLE="cursor: hand;" >
								<xsl:attribute name="ID">Calendar_button_<xsl:value-of select='@pin' /></xsl:attribute>
								<xsl:attribute name="onClick">showCalendar(getMonth('Available_<xsl:value-of select='@pin' />'),getYear('Available_<xsl:value-of select='@pin' />'),'Available_<xsl:value-of select='@pin' />',getObjSumLeft('Calendar_button_<xsl:value-of select='@pin' />'), (getObjSumTop('Calendar_button_<xsl:value-of select='@pin' />')+getObjHeight('Calendar_button_<xsl:value-of select='@pin' />')));</xsl:attribute>
							</IMG>
								<!--For the Calendar Control-->
							<SCRIPT language="JavaScript">
								createDivForCalendar();
							</SCRIPT>
							<!-- END: constant-->
						 </xsl:when>
						 <xsl:otherwise>				
							<!-- BEGIN: constant-->
							<INPUT TYPE="TEXT" SIZE="15" >
								<xsl:attribute name="NAME">Available_<xsl:value-of select='@pin' /></xsl:attribute>
								<xsl:choose>
									<xsl:when test="./pa[@ia='1']/tp/stt">
										<xsl:attribute name="VALUE"><xsl:value-of select="./pa[@ia='1']/tp/stt/@stv" /></xsl:attribute>										
									</xsl:when>	
									<xsl:when test="./pa[@ia='1']/tp/dt">
										<xsl:attribute name="VALUE"><xsl:value-of select="./pa[@ia='1']/tp/dt" /></xsl:attribute>										
									</xsl:when>	
									<xsl:otherwise>
										<xsl:attribute name="VALUE"><xsl:value-of select="./pa[@ia='1']" /></xsl:attribute>										
									</xsl:otherwise>
								</xsl:choose>
							</INPUT>
							<!-- END: constant-->
						 </xsl:otherwise>
						</xsl:choose>	
					</TD>			
				</TR>
			</TABLE>


			</TD>
		</TR>
	</TABLE>

  </xsl:template>
  
</xsl:stylesheet>
