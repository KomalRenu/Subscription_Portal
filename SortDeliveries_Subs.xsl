<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<!-- Copyright (c) 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. -->
	<xsl:template>
		<xsl:copy>
			<xsl:apply-templates select="@* | * | text()" /> 
		</xsl:copy></xsl:template>

	<xsl:template match="/subs">
		<subs><xsl:choose>
				<xsl:when test="/subs/inputs[./OrderBy $eq$ 'TIME']">
					<xsl:choose>
						<xsl:when test="/subs/inputs[./SortOrder $eq$ 'ASC']">
							<xsl:for-each select="sub" order-by="./@srtt">
									<xsl:apply-templates select="." /> 					
							</xsl:for-each>	
						</xsl:when>
						<xsl:otherwise>
							<xsl:for-each select="sub" order-by="-./@srtt">
									<xsl:apply-templates select="." /> 					
							</xsl:for-each>	
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="/subs/inputs[./OrderBy $eq$ 'SERVICE']">
					<xsl:choose>
						<xsl:when test="/subs/inputs[./SortOrder $eq$ 'ASC']">
							<xsl:for-each select="sub" order-by="./@svn">
									<xsl:apply-templates select="." /> 					
							</xsl:for-each>	
						</xsl:when>
						<xsl:otherwise>
							<xsl:for-each select="sub" order-by="-./@svn">
									<xsl:apply-templates select="." /> 					
							</xsl:for-each>	
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="/subs/inputs[./OrderBy $eq$ 'SCHEDULE']">
					<xsl:choose>
						<xsl:when test="/subs/inputs[./SortOrder $eq$ 'ASC']">
							<xsl:for-each select="sub" order-by="./@scn">
									<xsl:apply-templates select="." /> 					
							</xsl:for-each>	
						</xsl:when>
						<xsl:otherwise>
							<xsl:for-each select="sub" order-by="-./@scn">
									<xsl:apply-templates select="." /> 					
							</xsl:for-each>	
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="/subs/inputs[./OrderBy $eq$ 'ADDRESS']">
					<xsl:choose>
						<xsl:when test="/subs/inputs[./SortOrder $eq$ 'ASC']">
							<xsl:for-each select="sub" order-by="./@adid">
									<xsl:apply-templates select="." /> 					
							</xsl:for-each>	
						</xsl:when>
						<xsl:otherwise>
							<xsl:for-each select="sub" order-by="-./@adid">
									<xsl:apply-templates select="." /> 					
							</xsl:for-each>	
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
				</xsl:otherwise>
			</xsl:choose>
		</subs>
	</xsl:template>
</xsl:stylesheet> 