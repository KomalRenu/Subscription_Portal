<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<!-- Copyright (c) 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. -->
	<xsl:template>
		<xsl:copy>
			<xsl:apply-templates select="@* | * | text()" />
		</xsl:copy>
	</xsl:template>

	<xsl:template match="/">
		<ERRORS>
			<xsl:choose>
				<xsl:when test="/ERRORS/options[./OrderBy $eq$ 'Time']">
					<xsl:choose>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on' $and$ ./Warning $eq$ 'on' $and$ ./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg" order-by="sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg" order-by="-sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on' $and$ ./Warning $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $lt$ '4']" order-by="sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $lt$ '4']" order-by="-sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on' $and$ ./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4' $or$ ./errLevel $eq$ '1']" order-by="sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4' $or$ ./errLevel $eq$ '1']" order-by="-sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Warning $eq$ 'on' $and$ ./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $gt$ '1']" order-by="sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $gt$ '1']" order-by="-sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>

						<xsl:when test="/ERRORS/options[./Error $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '1']" order-by="sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '1']" order-by="-sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Warning $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '2']" order-by="sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '2']" order-by="-sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4']" order-by="sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4']" order-by="-sortTime">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="/ERRORS/options[./OrderBy $eq$ 'User']">
					<xsl:choose>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on' $and$ ./Warning $eq$ 'on' $and$ ./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg" order-by="user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg" order-by="-user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on' $and$ ./Warning $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $lt$ '4'	]" order-by="user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $lt$ '4'	]" order-by="-user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on' $and$ ./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4' $or$ ./errLevel $eq$ '1']" order-by="user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4' $or$ ./errLevel $eq$ '1']" order-by="-user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Warning $eq$ 'on' $and$ ./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $gt$ '1']" order-by="user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $gt$ '1']" order-by="-user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>

						<xsl:when test="/ERRORS/options[./Error $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '1']" order-by="user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '1']" order-by="-user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Warning $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '2']" order-by="user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '2']" order-by="-user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4']" order-by="user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4']" order-by="-user">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="/ERRORS/options[./OrderBy $eq$ 'Level']">
					<xsl:choose>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on' $and$ ./Warning $eq$ 'on' $and$ ./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg" order-by="errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg" order-by="-errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on' $and$ ./Warning $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $lt$ '4']" order-by="errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $lt$ '4']" order-by="-errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on' $and$ ./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4' $or$ ./errLevel $eq$ '1']" order-by="errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4' $or$ ./errLevel $eq$ '1']" order-by="-errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Warning $eq$ 'on' $and$ ./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $gt$ '1']" order-by="errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $gt$ '1']" order-by="-errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Error $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '1']" order-by="errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '1']" order-by="-errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Warning $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '2']" order-by="errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '2']" order-by="-errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="/ERRORS/options[./Message $eq$ 'on']">
							<xsl:choose>
								<xsl:when test="/ERRORS/options[./SortOrder $eq$ 'ASC']">
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4']" order-by="errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="ERRORS/errMsg[./errLevel $eq$ '4']" order-by="-errLevel">
											<xsl:apply-templates select="." />
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
				</xsl:otherwise>
			</xsl:choose>
		</ERRORS>
	</xsl:template>
</xsl:stylesheet>