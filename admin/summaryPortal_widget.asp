<P><%Call Response.Write(asDescriptors(853)) 'This concludes editing the portal configuration.  You have completed the following:%></P>
<UL>
 <LI><%Call Response.Write(asDescriptors(854) & " <B>" & Server.HTMLEncode(GerPortalName(GetVirtualDirectoryName())) & "</B>") 'portal virtual directory %></LI>
 <LI><%Call Response.Write(asDescriptors(855) & " <B>" & Server.HTMLEncode(Application.Value("SITE_NAME")) & "</B>") 'site definition%></LI>
</UL>
<P>
<%Call Response.Write(asDescriptors(856)) 'At this point your site is ready to be viewed through the portal.%>
<B><A HREF="../login.asp" TARGET="portal"><%Call Response.Write(asDescriptors(857)) 'Click here to access the subscription portal.%></A></B>
</P>