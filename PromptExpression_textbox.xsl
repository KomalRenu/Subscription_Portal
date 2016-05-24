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

					<!-- BEGIN: expression-->
					<xsl:if test="./res[. = '17' $or$ . = '18']" >
						<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_518").text</xsl:eval>:</FONT><BR />
					</xsl:if>
					<!-- Operator -->
					<INPUT TYPE="hidden">
						<xsl:attribute name="NAME">operator_<xsl:value-of select='@pin' /></xsl:attribute>
						<xsl:attribute name="VALUE"><xsl:value-of select='current/@op' /></xsl:attribute>
					</INPUT>
					<!-- available -->
					<SELECT>
						<xsl:attribute name="NAME">Available_<xsl:value-of select='@pin' /></xsl:attribute>
						<xsl:choose>
							<xsl:when test="./pa[@il='1' $or$ @idl='1']/mi[@pcc!='0']" >
						     	<xsl:choose>
						 		<xsl:when test="./res[. = '17' $or$ . = '18' $or$ . = '10']" >
									<xsl:for-each select="./pa[@il='1' $or$ @idl='1']/mi/oi" >
										<OPTION>
										<xsl:attribute name="VALUE">
										<xsl:value-of select="./@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_n" />\</xsl:attribute>
										<xsl:if test=".[@highlight='1']">
											<xsl:attribute name="SELECTED">1</xsl:attribute>
										</xsl:if>
										<xsl:value-of select="context()/@disp_n" />
										</OPTION>
									</xsl:for-each>
								</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<OPTION VALUE="--none--">--- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ---</OPTION>
							</xsl:otherwise>
						</xsl:choose>
					</SELECT>
					<P />
					
					<INPUT TYPE="text" SIZE="23" STYLE="font-family: courier">
					<xsl:attribute name="NAME">Input_<xsl:value-of select='@pin' /></xsl:attribute>					
					<xsl:if test="pa[@ia='1']/exp[not(./unknowndef) and not(./unknowndef)]">						
						<xsl:choose>
						<xsl:when test="./res[. = '10']" >
							<xsl:attribute name="VALUE"><xsl:value-of select="./pa[@ia='1']/exp/nd/nd/nd[1]" /></xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="VALUE"><xsl:value-of select="current/@disp_n" /></xsl:attribute>
						</xsl:otherwise>
						</xsl:choose>
					</xsl:if>
					<xsl:if test="pa[@ia='1']/exp/unknowndef">						
						<xsl:choose>
						<xsl:when test="./res[. = '17']" >							
							<xsl:attribute name="VALUE">
							<xsl:for-each select="./pa[@ia='1']/exp/nd/nd/nd[@nt='3' $and$ @et='1'$and$ @ddt='9']" >
								<xsl:value-of select="./cst" />
								<xsl:if expr="this.selectSingleNode('.') != this.parentNode.lastChild">
									<xsl:eval>";"</xsl:eval>
								</xsl:if>
							</xsl:for-each>
							</xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="VALUE"><xsl:value-of select="current/@disp_n" /></xsl:attribute>
						</xsl:otherwise>
						</xsl:choose>
					</xsl:if>
					</INPUT>
					
				</TD>
					<xsl:if test="./res[. = '17' $or$ . = '18']" >
						<xsl:for-each select="./pa[@il='1' $or$ @idl='1']/mi/oi" >
							<xsl:for-each select="./oi[@tp='21']" >
								<INPUT TYPE="hidden">
									<xsl:attribute name="NAME">form_<xsl:value-of select='../../../../@pin' />_<xsl:value-of select='../@did' /></xsl:attribute>
									<xsl:attribute name="VALUE">
									<xsl:value-of select="./@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@ddt" />
									</xsl:attribute>
								</INPUT>
							</xsl:for-each>
						</xsl:for-each>
					</xsl:if>
				</TR>
			</TABLE>

			</TD>
		</TR>
	</TABLE>
  </xsl:template>

<xsl:script><![CDATA[
 var ESC="&#027;";
 var RS="&#030;"; 
]]></xsl:script>
  
</xsl:stylesheet>
