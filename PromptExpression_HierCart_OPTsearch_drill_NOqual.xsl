<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<!-- Copyright 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. -->

<xsl:template language="JAVASCRIPT" match=".">
	<xsl:apply-templates select="pif" />
</xsl:template>

<xsl:template match="pif" >
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
			
	
				<!-- Begin Cart -->
				<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="300">
	
				<TR>

				<!-- pick hierachy -->
				<xsl:apply-templates select="./pickhier" />
	
				<!-- list -->
				<xsl:apply-templates select="./pa[@idl='1']/mi[@flag = 'ELEM']" />
				<xsl:apply-templates select="./pa[@idl='1']/mi[@flag = 'PICK_ELEM']" />
	
				<xsl:choose>
				<xsl:when test="./pa[@idl='1']/mi[@flag = 'PICK_ELEM']">
					<TD>
						<IMG SRC="Images/btn_add.gif" WIDTH="25" HEIGHT="25" BORDER="0">
							<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_537").text</xsl:eval></xsl:attribute>
							<xsl:attribute name="NAME">Add_<xsl:value-of select="./@pin" /></xsl:attribute> 
						</IMG>
						<BR />
						<xsl:choose>
						<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
							<A>
							<xsl:attribute name="HREF">javascript:RemoveItemsbyListObjectForHIInList(document.PromptForm.Selected_<xsl:value-of select="./@pin" />, document.PromptForm.Available_<xsl:value-of select="./@pin" />, document.PromptForm.Attribute_<xsl:value-of select="./@pin" />, '<xsl:value-of select="./@pin" />' )</xsl:attribute>
								<IMG SRC="Images/btn_remove.gif" WIDTH="25" HEIGHT="25" BORDER="0">					
									<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_875").text</xsl:eval></xsl:attribute> 
									<xsl:attribute name="NAME">Remove_<xsl:value-of select="./@pin" /></xsl:attribute> 
								</IMG>
							</A>				
						</xsl:when>
						<xsl:otherwise>	 			
							<INPUT TYPE="IMAGE" SRC="Images/btn_remove.gif" WIDTH="25" HEIGHT="25" BORDER="0">
								<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_875").text</xsl:eval></xsl:attribute>
								<xsl:attribute name="NAME">Remove_<xsl:value-of select="./@pin" /></xsl:attribute> 
							</INPUT>
						</xsl:otherwise>
						</xsl:choose>	
					</TD>	
				</xsl:when>
				<xsl:when test="./pa[@idl='1']/mi[@flag = 'ELEM']">
					<TD>
						<xsl:choose>
						<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
							<A>
							<xsl:attribute name="HREF">javascript:AddItemsbyListObjectForHI(document.PromptForm.Available_<xsl:value-of select="./@pin" />, document.PromptForm.Selected_<xsl:value-of select="./@pin" />, document.PromptForm.Attribute_<xsl:value-of select="./@pin" />, '<xsl:value-of select="./@pin" />' )</xsl:attribute>
								<IMG SRC="Images/btn_add.gif" WIDTH="25" HEIGHT="25" BORDER="0" >
									<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_537").text</xsl:eval></xsl:attribute> 
									<xsl:attribute name="NAME">Add_<xsl:value-of select="./@pin" /></xsl:attribute> 
								</IMG>
							</A>				
							<BR />
							<A>
							<xsl:attribute name="HREF">javascript:RemoveItemsbyListObjectForHIInList(document.PromptForm.Selected_<xsl:value-of select="./@pin" />, document.PromptForm.Available_<xsl:value-of select="./@pin" />, document.PromptForm.Attribute_<xsl:value-of select="./@pin" />, '<xsl:value-of select="./@pin" />' )</xsl:attribute>
								<IMG SRC="Images/btn_remove.gif" WIDTH="25" HEIGHT="25" BORDER="0">					
									<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_875").text</xsl:eval></xsl:attribute> 
									<xsl:attribute name="NAME">Remove_<xsl:value-of select="./@pin" /></xsl:attribute> 
								</IMG>
							</A>				
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
				</xsl:when>
				</xsl:choose>
	
				<!-- selected -->
				<xsl:apply-templates select="./pa[@ia='1']" />

				</TR>
				</TABLE>
	
				<!-- End Cart -->
					
				<xsl:if test="./filterHier">
					<INPUT TYPE="HIDDEN">					
						<xsl:attribute name="NAME">nuXML_filterHier_<xsl:value-of select="./@pin" /></xsl:attribute> 
						<xsl:attribute name="VALUE"><xsl:value-of select="./filterHier" /></xsl:attribute>
					</INPUT>
				</xsl:if>
	
				</TD>
				</TR>
			</TABLE>


			</TD>
		</TR>
	</TABLE>
</xsl:template>

<xsl:template match="./pa[@idl='1']/mi[@flag = 'PICK_ELEM']" >

	<TD WIDTH="18" VALIGN="TOP"><IMG SRC="Images/step2_gray.gif" WIDTH="18" HEIGHT="18" ALT="" BORDER="0" /></TD>
	
	<TD WIDTH="3" VALIGN="TOP"><IMG SRC="Images/1ptrans.gif" WIDTH="3" HEIGHT="1" ALT="" BORDER="0" /></TD>
	<TD VALIGN="TOP" CLASS="PROMPTTABBODYDISABLED">
		<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="1" WIDTH="100%" BGCOLOR="#000000"><TR><TD>
			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%" BGCOLOR="#FFFFFF">
				<TR>
					<TD BGCOLOR="#AAAA77" NOWRAP="1"><IMG SRC="Images/1ptrans.gif" HEIGHT="13" WIDTH="2" ALT="" BORDER="0" /><B><FONT COLOR="#666666"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_547").text</xsl:eval><BR /></FONT></B></TD>
				</TR>
				<TR><TD BGCOLOR="#000000" COLSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
				<TR><TD COLSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="5" ALT="" BORDER="0" /></TD></TR>
			</TABLE>
			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%" BGCOLOR="#FFFFFF"><TR><TD>
				<FONT COLOR="#666666"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
				<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_579").text</xsl:eval>
				<BR />
				</FONT>
				<IMG SRC="Images/1ptrans.gif" ALT="" WIDTH="100" HEIGHT="122" BORDER="0" />
			</TD></TR></TABLE>
		</TD></TR></TABLE>
	</TD>
	<TD WIDTH="3" VALIGN="TOP" ><IMG SRC="Images/1ptrans.gif" WIDTH="3" HEIGHT="1" ALT="" BORDER="0" /></TD>
	
	<TD WIDTH="10"><BR /></TD>
	
</xsl:template>

<xsl:template match="./pa[@idl='1']/mi[@flag = 'ELEM']" >
	<xsl:if test=".[@onedim='no']">
		<TD WIDTH="18" VALIGN="TOP"><IMG SRC="Images/step2_red.gif" WIDTH="18" HEIGHT="18" ALT="" BORDER="0" /></TD>
	</xsl:if>

	<TD WIDTH="3" VALIGN="TOP"><IMG SRC="Images/1ptrans.gif" WIDTH="3" HEIGHT="1" ALT="" BORDER="0" /></TD>

	<TD VALIGN="TOP" NOWRAP="1">
		<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="1" WIDTH="100%" BGCOLOR="#000000"><TR><TD>
			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%" BGCOLOR="#FFFFFF">
				<TR>
					<TD BGCOLOR="#CCCC99" NOWRAP="1"><IMG SRC="Images/arrow_down.gif" WIDTH="13" HEIGHT="13" ALT="" BORDER="0" /><B><FONT COLOR="#CC0000"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_547").text</xsl:eval><BR /></FONT></B></TD>
				</TR>
				<TR><TD BGCOLOR="#000000" COLSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD></TR>
				<TR><TD COLSPAN="3"><IMG SRC="Images/1ptrans.gif" WIDTH="1" HEIGHT="5" ALT="" BORDER="0" /></TD></TR>
			</TABLE>
			
			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2" WIDTH="100%" BGCOLOR="#FFFFFF">
			<TR>
			<TD>
				<!-- cart -->	
				<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
				<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_518").text</xsl:eval>:
				<BR /></FONT>
				<NOBR>
				<SELECT>
					<xsl:attribute name="NAME">Attribute_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
					<xsl:choose>
						<xsl:when test=".//oi[@tp='12' $and$ ./ad[@iep='1']]" >
						<xsl:for-each select=".//oi[@tp='12' $and$ (./ad[@iep='1'] $or$ @filtered='1') ]" >
							<OPTION>
							<xsl:attribute name="VALUE"><xsl:value-of select="./@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@highlight" /><xsl:if test=".[@filtered='1']"><xsl:eval no-entities="1">RS</xsl:eval>filtered</xsl:if></xsl:attribute>
							<xsl:if test=".[@highlight = '1']">
								<xsl:attribute name="SELECTED">1</xsl:attribute>
							</xsl:if>
							<xsl:value-of select="./@disp_n" />
							<xsl:if test=".[@filtered='1']"> *(filtered)* </xsl:if>
							</OPTION>
						</xsl:for-each>		
						</xsl:when>
						<xsl:otherwise>
							<OPTION VALUE="-none-">---- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ----</OPTION>
						</xsl:otherwise>
					</xsl:choose>
				</SELECT>
				<INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON">
					<xsl:attribute name="VALUE"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_110").text</xsl:eval>!</xsl:attribute>
					<xsl:attribute name="NAME">AttributeGO_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
				</INPUT>	
				</NOBR>
				<HR SIZE="1" />
					
				<!-- Search Field -->
				<xsl:if test=".[$not$ .//oi[@tp='12' $and$ @highlight='1']/ad[@lt = '2']]">
					<xsl:choose>
						<xsl:when test="./search" >
							<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_538").text</xsl:eval><BR /></FONT>
								<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
								<TR>
									<TD VALIGN="TOP">
										<INPUT TYPE="TEXT" SIZE="16" CLASS="PromptSearch">
											<xsl:attribute name="VALUE">
												<xsl:value-of select="./search/@text" />
											</xsl:attribute>
											<xsl:attribute name="NAME">Search_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
										</INPUT>
									</TD>
									<TD VALIGN="TOP">
										<INPUT TYPE="IMAGE" SRC="Images/btn_find_lightgray.gif" WIDTH="23" HEIGHT="23" BORDER="0">
											<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_515").text</xsl:eval></xsl:attribute>
											<xsl:attribute name="NAME">Find_<xsl:value-of select="ancestor(pif)/@pin" /></xsl:attribute> 
										</INPUT>
									</TD>
								</TR>
								</TABLE>
						</xsl:when>			
						<xsl:otherwise>
							<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%">
							<TR>
								<TD>
								<FONT COLOR="#000000"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_57").text</xsl:eval><BR /></FONT>
								<BR />
								</TD>
							</TR>	
							</TABLE>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>

				<!-- available elements -->
				<xsl:choose>
					<xsl:when test=".[$not$ .//oi[@tp='12' $and$ ./ad[@iep='1']]]" >
						<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%">
						<TR><TD>
							<BR /><FONT COLOR="#000000"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_554").text</xsl:eval><BR /><BR /><BR /></FONT>
						</TD></TR>	
						</TABLE>
					</xsl:when>						
					<xsl:when test=".//oi[@tp='12' $and$ @highlight='1']/ad[@lt = '2']" >
						<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%">
						<TR>
							<TD>
							<BR />
							<FONT COLOR="#000000"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_960").text</xsl:eval>.<BR /></FONT>
							<FONT COLOR="#000000"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_961").text</xsl:eval>.<BR /></FONT>
							</TD>
						</TR>	
						</TABLE>
					</xsl:when>
					<xsl:otherwise>
						<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_548").text</xsl:eval>:<BR /></FONT>
						
						<SELECT SIZE="10" MULTIPLE="1" >
						<xsl:attribute name="NAME">Available_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
							<xsl:choose>
							<xsl:when test=".//oi[@tp='12' $and$ @highlight='1']/es/e" >
								<xsl:for-each select=".//oi[@tp='12' $and$ @highlight='1']/es/e" >
									<xsl:if test=".[$not$ @selected]">
									<OPTION>
										<xsl:if test=".[@highlight='1']">
											<xsl:attribute name="SELECTED">1</xsl:attribute>
										</xsl:if>
										<xsl:attribute name="VALUE"><xsl:value-of select="./@ei" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_n" /></xsl:attribute>
										<xsl:value-of select="./@disp_n" />
									</OPTION>
									</xsl:if>
								</xsl:for-each>
								<xsl:if test=".//oi[@tp='12' $and$ @highlight='1']/es[@acc='0']">
									<OPTION VALUE="-none-">--- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ---</OPTION>
								</xsl:if>
							</xsl:when>
							<xsl:otherwise>
								<OPTION VALUE="-none-">---- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ----</OPTION>
							</xsl:otherwise>
							</xsl:choose>
						</SELECT><BR />
					</xsl:otherwise>
				</xsl:choose>
									
				<xsl:if test=".[.//oi[@tp='12' $and$ ./ad[@iep='1']]]" >
					<!-- incremental fetch links -->
					<xsl:if test=".[$not$ .//oi[@tp='12' $and$ @highlight='1']/ad[@lt = '2']]">
					<xsl:if test="../../increfetch/curr[@start != '']" >
							<xsl:apply-templates select="../../increfetch" />
						</xsl:if>
					</xsl:if>
					<BR />

					<!-- Drill -->
					<xsl:if test=".//oi[@tp='14' $and$ @highlight='1']/mi/oi[1]">
						<xsl:apply-templates select=".//oi[@tp='12' $and$ @highlight='1']/ad" />
					</xsl:if>
				</xsl:if>

				</TD></TR></TABLE>
				<!-- End Cart -->
				
		</TD></TR></TABLE>
	</TD>
	<TD WIDTH="3" VALIGN="TOP"><IMG SRC="Images/1ptrans.gif" WIDTH="3" HEIGHT="1" ALT="" BORDER="0" /></TD>
	
	<TD WIDTH="10"><BR /></TD>

</xsl:template>

<xsl:template match="./pa[@ia='1']">
	<TD NOWRAP="1">
		<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_536").text</xsl:eval>:</FONT><BR />
		<SELECT SIZE="10" MULTIPLE="1">
		<xsl:attribute name="NAME">Selected_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
		<xsl:choose>
			<xsl:when test="./exp/unknowndef" >	
				<OPTION VALUE="-default-"><xsl:value-of select="./exp/unknowndef/@text" /></OPTION>
			</xsl:when>
			<xsl:when test="./exp/nd/nd" >	
				<xsl:for-each select="./exp/nd/nd" >	
					<xsl:choose>
					<xsl:when test=".[@et='5']">
						<OPTION>
						<xsl:attribute name="VALUE"><xsl:value-of select="./nd[0]/oi/@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./nd[0]/@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
						<xsl:value-of select="./nd[0]/@disp_n" />:
						</OPTION>
						<xsl:for-each select="./nd/oi/es/e" >	
							<OPTION>
								<xsl:attribute name="VALUE"><xsl:value-of select="./@ei" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="../../../@disp_id" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
								<xsl:eval no-entities="1">NBSP+NBSP+NBSP</xsl:eval><xsl:value-of select="./@disp_n" />
							</OPTION>
						</xsl:for-each>
					</xsl:when>
					<xsl:otherwise>
						<OPTION>
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
						</OPTION>
					</xsl:otherwise>
					</xsl:choose>	
				</xsl:for-each>
			</xsl:when>
			<xsl:otherwise>
				<OPTION VALUE="-none-">--- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ---</OPTION>
			</xsl:otherwise>
		</xsl:choose>
	</SELECT>
	</TD>
	
	<!-- AND/OR all subexpressions -->
	<xsl:choose>
	<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
		<TD NOWRAP="1">
			<DIV>
				<xsl:attribute name="ID">ANDOR_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
				<xsl:attribute name="NAME">ANDOR_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
			<xsl:choose>
			<xsl:when test="./exp[./nd/nd[1] $and$ $not$ unknowndef]">
				<xsl:attribute name="STYLE">display:block;</xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name="STYLE">display:none;</xsl:attribute>
			</xsl:otherwise>
			</xsl:choose>		
			<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
				<CENTER><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_533").text</xsl:eval>:<BR /></CENTER>
				<INPUT TYPE="RADIO" VALUE="AND">
				<xsl:attribute name="NAME">FilterOperator_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
				<xsl:if test="./exp/nd/op[@fnt='19']">
					<xsl:attribute name="CHECKED">1</xsl:attribute>
				</xsl:if>
				<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_534").text</xsl:eval><BR />
				</INPUT>
				<INPUT TYPE="RADIO" VALUE="OR">
				<xsl:attribute name="NAME">FilterOperator_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
				<xsl:if test="./exp/nd/op[@fnt='20']">
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
		<xsl:if test="./exp[./nd/nd[1] $and$ $not$ unknowndef]">
			<TD NOWRAP="1">
				<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
					<BR />
					<CENTER><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_533").text</xsl:eval>:<BR /></CENTER>
					<INPUT TYPE="RADIO" VALUE="AND">
						<xsl:attribute name="NAME">FilterOperator_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
						<xsl:if test="./exp/nd/op[@fnt='19']">
							<xsl:attribute name="CHECKED">1</xsl:attribute>
						</xsl:if>
						<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_534").text</xsl:eval><BR />
					</INPUT>
					<INPUT TYPE="RADIO" VALUE="OR">
						<xsl:attribute name="NAME">FilterOperator_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
						<xsl:if test="./exp/nd/op[@fnt='20']">
							<xsl:attribute name="CHECKED">1</xsl:attribute>
						</xsl:if>
						<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_535").text</xsl:eval><BR />
					</INPUT>
				</FONT>
			</TD>
		</xsl:if>
	</xsl:otherwise>
	</xsl:choose>		
</xsl:template>


<xsl:template match="increfetch">
	<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
	<!-- previous -->	
	<xsl:if test="./prev[@count $ne$ '']" >
		<INPUT TYPE="IMAGE" SRC="Images/arrow_left_inc_fetch.gif" WIDTH="5" HEIGHT="10" BORDER="0">
			<xsl:attribute name="ALT"><xsl:value-of select="./prev/@title" /></xsl:attribute>
			<xsl:attribute name="NAME">prev_<xsl:value-of select="ancestor(pif)/@pin" /></xsl:attribute> 
		</INPUT> 
		<INPUT TYPE="HIDDEN">
			<xsl:attribute name="NAME">BBprev_<xsl:value-of select="ancestor(pif)/@pin" /></xsl:attribute> 
			<xsl:attribute name="VALUE"><xsl:value-of select="./prev/@link" /></xsl:attribute>
		</INPUT> 
	</xsl:if>

	<!-- current -->
	<xsl:value-of select="./curr/@title" />

	<!-- next -->
	<xsl:if test="./next[@count $ne$ '']" >
		<INPUT TYPE="IMAGE" SRC="Images/arrow_right_inc_fetch.gif" WIDTH="5" HEIGHT="10" BORDER="0">
			<xsl:attribute name="ALT"><xsl:value-of select="./next/@title" /></xsl:attribute>
			<xsl:attribute name="NAME">next_<xsl:value-of select="ancestor(pif)/@pin" /></xsl:attribute> 
		</INPUT> 
		<INPUT TYPE="HIDDEN">
			<xsl:attribute name="NAME">BBnext_<xsl:value-of select="ancestor(pif)/@pin" /></xsl:attribute> 
			<xsl:attribute name="VALUE"><xsl:value-of select="./next/@link" /></xsl:attribute>
		</INPUT> 
	</xsl:if>

	<INPUT TYPE="HIDDEN">
	<xsl:attribute name="NAME">BBcurr_<xsl:value-of select="ancestor(pif)/@pin" /></xsl:attribute> 
	<xsl:attribute name="VALUE"><xsl:value-of select="./curr/@start" /></xsl:attribute>
	</INPUT>
	</FONT>
</xsl:template>

<xsl:template match=".//oi[@tp='12' $and$ @highlight='1']/ad">
	<xsl:if test="..[@lt != '2' or $not$ @lt]" >
	  <xsl:if test=".[./ar/rc[./oi] or .[@iep != '1'] or ./ar/rc/oi or ./ar/rp/oi]"> 
		<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_183").text</xsl:eval>:<BR /></FONT>
	
		<SELECT>
		<xsl:attribute name="NAME">Drill_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
			<xsl:if test=".[@iep != '1']" >
				<OPTION VALUE="-none-">-- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> --</OPTION>
			</xsl:if>

			<xsl:if test="./ar/rc[$not$ ./oi]">
				<xsl:if test="./ar/rp[$not$ ./oi]">
					<OPTION VALUE="-none-">-- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> --</OPTION>
				</xsl:if>
			</xsl:if>

			<xsl:if test="./ar/rc/oi">
				<OPTION>
					<xsl:attribute name="VALUE"><xsl:value-of select="./ar/rc/oi/@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./ar/rc/oi/@did" /><xsl:eval no-entities="1">RS</xsl:eval>down</xsl:attribute>
					<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_153").text</xsl:eval>: 
				</OPTION>
				<xsl:for-each select="./ar/rc/oi">
					<OPTION>
						<xsl:attribute name="VALUE"><xsl:value-of select="./@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@did" /><xsl:eval no-entities="1">RS</xsl:eval>down</xsl:attribute>
						<xsl:if test=".[@highlight = '1']"><xsl:attribute name="SELECTED">1</xsl:attribute></xsl:if>
						<xsl:value-of select="./@disp_n" />
					</OPTION>
				</xsl:for-each>	
			</xsl:if>

			<xsl:if test="./ar/rp/oi">
				<OPTION>
					<xsl:attribute name="VALUE"><xsl:value-of select="./ar/rp/oi/@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./ar/rp/oi/@did" /><xsl:eval no-entities="1">RS</xsl:eval>up</xsl:attribute>
					<xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_152").text</xsl:eval>: 
				</OPTION>
				<xsl:for-each select="./ar/rp/oi">
					<OPTION>
						<xsl:attribute name="VALUE"><xsl:value-of select="./@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@did" /><xsl:eval no-entities="1">RS</xsl:eval>up</xsl:attribute>
						<xsl:if test=".[@highlight = '1']"><xsl:attribute name="SELECTED">1</xsl:attribute></xsl:if>
						<xsl:value-of select="./@disp_n" />
					</OPTION>
				</xsl:for-each>	
			</xsl:if>
		</SELECT>

		<INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON">
			<xsl:attribute name="VALUE"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_145").text</xsl:eval></xsl:attribute>
			<xsl:attribute name="NAME">DrillGO_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
		</INPUT>
	  </xsl:if>
	</xsl:if>
</xsl:template>

<xsl:template match="pickhier">
	<xsl:choose>
	<xsl:when test="../available[@flag = 'ELEM' $or$ @flag = 'QUAL']">
		<TD WIDTH="18" VALIGN="TOP"><IMG SRC="Images/step1_black.gif" WIDTH="18" HEIGHT="18" ALT="" BORDER="0" /></TD>
	</xsl:when>
	<xsl:otherwise>
		<TD WIDTH="18" VALIGN="TOP"><IMG SRC="Images/step1_red.gif" WIDTH="18" HEIGHT="18" ALT="" BORDER="0" /></TD>
	</xsl:otherwise>
	</xsl:choose>
	
	<TD WIDTH="3" VALIGN="TOP"><IMG SRC="Images/1ptrans.gif" WIDTH="3" HEIGHT="1" ALT="" BORDER="0" /></TD>
	
	<TD VALIGN="TOP" NOWRAP="1">
		<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="1" WIDTH="100%" BGCOLOR="#000000"><TR><TD>
			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%" BGCOLOR="#FFFFFF">
				<TR>
					<TD BGCOLOR="#CCCC99">
						<IMG SRC="Images/1ptrans.gif" HEIGHT="13" WIDTH="1" ALT="" BORDER="0" />
						<xsl:choose>
						<xsl:when test="../available[@flag = 'ELEM' $or$ @flag = 'QUAL']">
							<B><FONT COLOR="#000000"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_543").text</xsl:eval>:<BR /></FONT></B>
						</xsl:when>
						<xsl:otherwise>
							<B><FONT COLOR="#CC0000"><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_543").text</xsl:eval>:<BR /></FONT></B>
						</xsl:otherwise>
						</xsl:choose>
					</TD>
				</TR>
				<TR><TD BGCOLOR="#000000"><IMG SRC="Images/1ptrans.gif" HEIGHT="1" WIDTH="1" ALT="" BORDER="0" /></TD></TR>
				<TR><TD><IMG SRC="Images/1ptrans.gif" HEIGHT="5" WIDTH="1" ALT="" BORDER="0" /></TD></TR>
			</TABLE>
			<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2" WIDTH="100%" BGCOLOR="#FFFFFF"><TR><TD>
				<!-- Subfolders -->
				<xsl:if test="./subfolders/link" >
					<NOBR>
					<SELECT>
					<xsl:attribute name="NAME">sf_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
						<xsl:for-each select="./subfolders/link">
							<OPTION>
								<xsl:if test=".[@selected='1']" >
									<xsl:attribute name="SELECTED">1</xsl:attribute>
								</xsl:if>
								<xsl:attribute name="VALUE"><xsl:value-of select="./@did" /></xsl:attribute>
								<xsl:value-of select="./@fd" />
							</OPTION>
						</xsl:for-each>
					</SELECT>
					<INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON">
						<xsl:attribute name="VALUE"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_110").text</xsl:eval>!</xsl:attribute>
						<xsl:attribute name="NAME">sfGO_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
					</INPUT>
					</NOBR>
					<BR />
				</xsl:if>
			
				<!-- hierachies -->
				<SELECT SIZE="10">
				<xsl:attribute name="NAME">hi_<xsl:value-of select='ancestor(pif)/@pin' /></xsl:attribute>
					<xsl:choose>
					<xsl:when test="./hierachies/hi" >
						<xsl:for-each select="./hierachies/hi">
							<OPTION>
							<xsl:if test=".[@selected='1']" >
								<xsl:attribute name="SELECTED">1</xsl:attribute>
							</xsl:if>
							<xsl:attribute name="VALUE"><xsl:value-of select="./@n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@did" /></xsl:attribute>
							<xsl:value-of select="./@n" />
							</OPTION>
						</xsl:for-each>
					</xsl:when>
					<xsl:otherwise>
						<OPTION VALUE="-none-">---- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ----</OPTION>
					</xsl:otherwise>
					</xsl:choose>
				</SELECT>
				<BR />
			
				<INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON">
				<xsl:attribute name="NAME">hiGO_<xsl:value-of select="ancestor(pif)/@pin" /></xsl:attribute>
				<xsl:attribute name="VALUE"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_584").text</xsl:eval></xsl:attribute>
				</INPUT>
			</TD></TR></TABLE>
		</TD></TR></TABLE>
	</TD>
	
	<TD WIDTH="3" VALIGN="TOP"><IMG SRC="Images/1ptrans.gif" WIDTH="3" HEIGHT="1" ALT="" BORDER="0" /></TD>
	
	<TD WIDTH="18" VALIGN="TOP"><IMG SRC="Images/1ptrans.gif" WIDTH="18" HEIGHT="1" ALT="" BORDER="0" /></TD>

</xsl:template>

<xsl:script><![CDATA[
 var NBSP = "&#160;"
 var ESC="&#027;";
 var RS="&#030;"; 
]]></xsl:script>
		
</xsl:stylesheet>
