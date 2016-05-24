<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<!-- Copyright 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. -->
   <xsl:template language="JAVASCRIPT" match=".">
	<xsl:apply-templates select="./pif[@pt='6']" />
   </xsl:template>

   <xsl:template match='pif[@pt="6"]'> 
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

					<!-- BEGIN: HI browse links -->
					<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
					  <TR>
						<TD COLSPAN="2">
							<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_549").text</xsl:eval>:</FONT>
						    <!-- HI path links -->
							<xsl:if test="./pa[@idl='1']/mi/oi/as" >
								<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>

									<xsl:for-each select="./pa[@idl='1']/mi/oi/as//a" order="">
										<xsl:value-of select="./oi/@disp_n" /> > 
									</xsl:for-each>
									
									<xsl:value-of select="./pa[@idl='1']/mi/oi/@disp_n" /> > 

								<xsl:if test="./pa[@idl='1']/mi/oi/as/a">
									<IMG SRC="Images/1ptrans.gif" WIDTH="10" HEIGHT="10" BORDER="0" />
									<INPUT type="IMAGE" SRC="Images/arrow_shiftleft.gif" WIDTH="16" HEIGHT="16" BORDER="0">
										<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_865").text</xsl:eval></xsl:attribute>
										<xsl:attribute name="NAME">HiParentGO_<xsl:value-of select="./@pin" /></xsl:attribute> 
									</INPUT> 
									<INPUT TYPE="HIDDEN">
										<xsl:attribute name="NAME">Hiparent_<xsl:value-of select="./@pin" /></xsl:attribute> 
										<xsl:attribute name="VALUE"><xsl:value-of select="./pa[@idl='1']/mi/oi/as//a[end()]/oi/@did" /></xsl:attribute>
									</INPUT> 
								</xsl:if>
							</FONT>
							</xsl:if>
						</TD>
					  </TR>
					</TABLE>
					<!-- END: HI browse links -->

					<!-- BEGIN: Subfolders -->
					<xsl:choose>
						<!-- <xsl:when test="./hilinks/link" > -->
						<xsl:when test=".//fct/oi[@tp='8']" >
							<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_550").text</xsl:eval>:</FONT><BR />
							<SELECT>
							<xsl:attribute name="NAME">Hilink_<xsl:value-of select='@pin' /></xsl:attribute>
								<xsl:for-each select=".//fct/oi[@tp='8']">
									<OPTION>
										<xsl:attribute name="VALUE"><xsl:value-of select="./@did" /></xsl:attribute>
										<xsl:value-of select="./@disp_n" />
									</OPTION>
								</xsl:for-each>
							</SELECT>
						</xsl:when>
						<xsl:otherwise>
							<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_550").text</xsl:eval>:</FONT><BR />
							<SELECT>
							<xsl:attribute name="NAME">Hilink_<xsl:value-of select='@pin' /></xsl:attribute>
								<OPTION VALUE="-none-">---- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ----</OPTION>
							</SELECT>
						</xsl:otherwise>
					</xsl:choose>
					<INPUT TYPE="SUBMIT" CLASS="GOLDBUTTON">
						<xsl:attribute name="VALUE"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_110").text</xsl:eval>!</xsl:attribute>
						<xsl:attribute name="NAME">HilinkGO_<xsl:value-of select='@pin' /></xsl:attribute>
					</INPUT>		
							
					<BR />
					<!-- END: Subfolders -->

					<!-- BEGIN: search field --> 
					<xsl:if test="./search" >	
						<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_538").text</xsl:eval></FONT>
						<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
						  <TR>
							<TD VALIGN="TOP">
								<INPUT TYPE="TEXT" SIZE="12" CLASS="PromptSearch">
									<xsl:attribute name="NAME">Search_<xsl:value-of select="@pin" /></xsl:attribute>
									<xsl:attribute name="VALUE"><xsl:value-of select="./search/@text" /></xsl:attribute>
								</INPUT>
							</TD>
							<TD VALIGN="TOP">
								<INPUT TYPE="IMAGE" SRC="Images/btn_find.gif" HEIGHT="23" WIDTH="23" BORDER="0">
									<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_515").text</xsl:eval></xsl:attribute>
									<xsl:attribute name="NAME">Find_<xsl:value-of select="./@pin" /></xsl:attribute> 
								</INPUT>
							</TD>
						</TR>
						</TABLE>
					</xsl:if>
					<!-- END: search field --> 

					<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
					  <TR>
						<TD COLSPAN="2"><FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_551").text</xsl:eval>:</FONT></TD>
						<TD><FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_514").text</xsl:eval>:</FONT></TD>
					  </TR>
					  <TR>
						<TD VALIGN="TOP">
							<SELECT SIZE="10" MULTIPLE="1">
							<xsl:attribute name="NAME">Available_<xsl:value-of select='@pin' /></xsl:attribute>
								<xsl:choose>
								<xsl:when test=".//fct/oi[@tp!='8']" >
							     	<xsl:for-each select=".//fct/oi[@tp!='8']" >	
					 				<xsl:if test=".[$not$ @selected]">
										<OPTION>
											<xsl:if test=".[@highlight='1']">
												<xsl:if test="/mi/inputs/accessibilityModeOff">
													<xsl:attribute name="SELECTED">1</xsl:attribute>
												</xsl:if>											
											</xsl:if>
											<xsl:attribute name="VALUE"><xsl:value-of select="./@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@tp" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_n" /></xsl:attribute>
											<xsl:value-of select="./@disp_n" />
										</OPTION>
									</xsl:if>
									</xsl:for-each>
									<xsl:if test=".//fct[@acc='0']">
										<OPTION VALUE="-none-">--- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ---</OPTION>
									</xsl:if>
								</xsl:when>
								<xsl:otherwise>
									<OPTION VALUE="-none-">--- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ---</OPTION>
								</xsl:otherwise>
								</xsl:choose>
							</SELECT>
						</TD>
			 			<TD ALIGN="CENTER">
			 				<xsl:choose>
			 				<xsl:when expr="this.selectSingleNode('/mi/inputs/DHTML').text=='1'">
								<A>
								<xsl:attribute name="HREF">javascript:AddItemsbyListObjectForObjectPrompt(document.PromptForm.Available_<xsl:value-of select="./@pin" />, this.document.PromptForm.Selected_<xsl:value-of select="./@pin" /> )</xsl:attribute>
									<IMG SRC="Images/btn_add.gif" WIDTH="25" HEIGHT="25" BORDER="0">
										<xsl:attribute name="ALT"><xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_537").text</xsl:eval></xsl:attribute> 
										<xsl:attribute name="NAME">Add_<xsl:value-of select="./@pin" /></xsl:attribute> 
									</IMG
								></A>
								<BR />
								<A>
								<xsl:attribute name="HREF">javascript:MoveItemsbyListObject(document.PromptForm.Selected_<xsl:value-of select="./@pin" />, this.document.PromptForm.Available_<xsl:value-of select="./@pin" /> )</xsl:attribute>
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
						<TD VALIGN="TOP">
							<SELECT SIZE="10" MULTIPLE="1">
							<xsl:attribute name="NAME">Selected_<xsl:value-of select='@pin' /></xsl:attribute>
								<xsl:choose>
								<xsl:when test="./pa[@ia='1']/oi" >	
									<xsl:for-each select="./pa[@ia='1']/oi" >	
									<OPTION>
										<xsl:attribute name="VALUE"><xsl:value-of select="./@did" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@tp" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_n" /><xsl:eval no-entities="1">RS</xsl:eval><xsl:value-of select="./@disp_id" /></xsl:attribute>
										<xsl:value-of select="./@disp_n" />
									</OPTION>
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<OPTION VALUE="-none-">--- <xsl:eval no-entities="1">this.selectSingleNode("/mi/inputs/Desc_512").text</xsl:eval> ---</OPTION>
								</xsl:otherwise>
								</xsl:choose>
							</SELECT>
						</TD>
					  </TR>
			 		  
			 		  <!-- BEGIN: incremental fetch links -->
			 		  <xsl:if test="./increfetch/*" >
			 			<xsl:apply-templates select="./increfetch" />
			 		  </xsl:if>
			 		  <!-- END: incremental fetch links -->
					</TABLE>
					
					<INPUT TYPE="HIDDEN">
						<xsl:attribute name="NAME">BBcurr_<xsl:value-of select="./@pin" /></xsl:attribute> 
						<xsl:attribute name="VALUE"><xsl:value-of select="./increfetch/curr/@start" /></xsl:attribute>
					</INPUT>
			
					</TD>
				</TR>
			</TABLE>

			</TD>
		</TR>
	</TABLE>
  </xsl:template>

  <xsl:template match="increfetch">
	<TR>
		<TD VALIGN="TOP" NOWRAP="1" COLSPAN="3">
			<FONT><xsl:attribute name="FACE"><xsl:eval>this.selectSingleNode("/mi/inputs/FontFamily").text</xsl:eval></xsl:attribute><xsl:attribute name="SIZE"><xsl:eval>this.selectSingleNode("/mi/inputs/smallFont").text</xsl:eval></xsl:attribute>
				<!-- previous -->	
				<xsl:if test="./prev[@count $ne$ '']" >
					<INPUT TYPE="IMAGE" SRC="Images/arrow_left_inc_fetch.gif" WIDTH="5" HEIGHT="10" BORDER="0">
						<xsl:attribute name="NAME">prev_<xsl:value-of select="./@pin" /></xsl:attribute>
						<xsl:attribute name="ALT"><xsl:value-of select="./prev/@title" /></xsl:attribute>
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
					<INPUT type="IMAGE" SRC="Images/arrow_right_inc_fetch.gif" WIDTH="5" HEIGHT="10" BORDER="0">
						<xsl:attribute name="NAME">next_<xsl:value-of select="./@pin" /></xsl:attribute>
						<xsl:attribute name="ALT"><xsl:value-of select="./next/@title" /></xsl:attribute>
					</INPUT> 
					<INPUT TYPE="HIDDEN">
						<xsl:attribute name="NAME">BBnext_<xsl:value-of select="./@pin" /></xsl:attribute> 
						<xsl:attribute name="VALUE"><xsl:value-of select="./next/@link" /></xsl:attribute>
					</INPUT> 
				</xsl:if>
			</FONT>
		</TD>
	</TR>
  </xsl:template>

<xsl:script><![CDATA[
 var RS="&#030;"; 
]]></xsl:script>

</xsl:stylesheet>
