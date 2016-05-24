<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>


<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->

<%
Response.write "<HTML>"
Response.write "<HEAD>"
Response.Write(putMETATagWithCharSet())
Response.write "<TITLE>"
Response.Write "Debug Page - MicroStrategy Narrowcast Server"
Response.write "</TITLE>"
Response.write "</HEAD>"
Dim lStatus

    aPageInfo(S_NAME_PAGE) = "debug.asp"
    aPageInfo(S_TITLE_PAGE) = ""
    aPageInfo(N_CURRENT_OPTION_PAGE) = 0
	lStatus = checkSiteConfiguration()
%>

<!-- #include file="../NSStyleSheet.asp" -->
<BODY BGCOLOR="FFFFFF" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
  <TR>
    <TD VALIGN="TOP" COLSPAN="5">
		<!-- begin header -->
		<!-- #include file="admin_header.asp" -->
		<!-- end header -->
    </TD>
  </TR>
  <TR HEIGHT="100%">
	<TD VALIGN="TOP" WIDTH="157" HEIGHT="100%" BGCOLOR="#666666">
		      <!-- begin toolbar -->
		        <!-- #include file="_toolbar_engine_config.asp" -->
		      <!-- end toolbar -->
   	</TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>
    <TD WIDTH="96%" valign="TOP">

<%
Dim oSite
Dim oDoc
Dim sitesXML
Dim oSites
Dim siteID
Dim tempDoc
Dim oAdmin
Dim engineProps
Dim metadataProps

on error resume next

set oSite = Server.CreateObject("Bridge2API.SiteInfo")
If Err.Number <> 0  then
	Response.Write "Error " & Err.Number & " " & Err.Description & " " & "while creating the site object"
	Err.clear
	response.end
End If

set oAdmin = Server.CreateObject("Bridge2API.Admin")
If Err.Number <> 0  then
	Response.Write "Error " & Err.Number & " " & Err.Description & " " & "while creating the admin object"
	Err.clear
	response.end
End If

Response.write "<BR>"
Response.write "<B>Subscription Engine information</B>"
Response.write "<BR>"
engineProps = oAdmin.getSubscriptionEngineLocation()
call RenderObjectProperties("","",engineProps,"mi/oi/prs")
If Err.Number <> 0  then
	Response.Write "Error " & Err.Number & " " & Err.Description & " " & "retieving the subscription engine location"
	Err.clear
	response.end
End If

Response.write "<BR>"
Response.write "<B>Metadata Connection information</B>"
Response.write "<BR>"

metadataProps = oAdmin.getMetadataConnectionProperties()
call RenderObjectProperties("","",metadataProps,"mi/oi/prs")
If Err.Number <> 0  then
	Response.Write "Error " & Err.Number & " " & Err.Description & " " & "retrieving the metadata connection information"
	Err.clear
	response.end
End If
Response.write "<BR>"

Response.write "<BR>"
Response.write "<B>Information about the sites in the current metadata</B>"
Response.write "<BR>"

siteID = Request.QueryString.Item("siteID")
If Err.Number <> 0  then
	Response.Write "Error " & Err.Number & " " & Err.Description & " " & "reading the query string"
	Err.clear
	Response.end
End If

if siteID <> "" then
	call RenderSiteObjects(siteID,siteID,"1001","")
else
	sitesXML = oSite.getAllSites()
	If Err.Number <> 0  then
		Response.Write "Error " & Err.Number & " " & Err.Description & " " & "while retrieving the list of sites"
		Err.clear
		Response.end
	End If
	call RenderObjects(sitesXML)
end if

	Response.Write "</TD>"
 	Response.Write "</Table>"
 	Response.Write "</BODY>"
	Response.Write "</Table>"
	If Err.Number <> 0  then
		Response.Write "Error " & Err.Number & " " & Err.Description & " " & "saving the document"
		Response.end
		Err.clear
	End If

set oDoc = nothing
set oSite = nothing
set oAdmin = nothing

Function RenderObjects(sitesXML)
	Dim oSiteNode
	Dim oAttrNode
	Dim siteID
	Dim siteName
	Dim oDoc


	'Just slap a <sites> tag around the dom document
	'oDomDoc.appendChild()

	set oDoc = Server.CreateObject("Microsoft.XMLDOM")
	oDoc.loadXML(sitesXML)
	If Err.Number <> 0  then
		Response.Write "Error " & Err.Number & " " & Err.Description & " " & "while loading list of sites"
		Err.clear
	End If

	oDoc.async = False

	Set oSites = oDoc.selectNodes("//oi")
	For each oSiteNode in oSites
		Dim numAttr
		For each oAttrNode in oSiteNode.attributes
			if oAttrNode.NodeName = "id" then siteID = oAttrNode.text
			if oAttrNode.NodeName = "n" then siteName = oAttrNode.text
		next
		'get all the child properties for the site
		call RenderSiteObjects(siteID,siteID,"1001",siteName)
	next


	set oSites = nothing
	set oDoc = nothing

	If Err.Number <> 0  then
		Response.Write "Error " & Err.Number & " " & Err.Description & " " & "while listing site properties"
		Err.clear
	End If


End Function

Function RenderSiteObjects(siteID,ObjectID,objectType,objectName)
on error resume next
Dim sitePropsXML
Dim oPropNodes
Dim oObjNodes
Dim oDoc

	Response.write "<TABLE BORDER=1>"
   	Response.Write "<TR>"
   	Response.Write "<TD>"

	Response.Write objectname & " " & objectType & " " & objectID
	sitePropsXML = getObjectProperties(siteID,ObjectID)
	If Err.Number <> 0  then
		Response.Write "RenderSiteObjects::Error " & Err.Number & " " & Err.Description
		Err.clear
	End If

	set oDoc = Server.CreateObject("Microsoft.XMLDOM")
	oDoc.loadXML(sitePropsXML)
	Set oObjNodes = oDoc.selectNodes("mi/in/oi/mi/in/oi")

	'render the xml definition of the object before the properties
	if objectType="1005" or objectType="1012" then
		call RenderObjectDefinition(siteID,objectID,objectType)
	end if

	if oPropNodes.length <> 0 then
		call RenderObjectProperties(siteID,objectID,sitePropsXML,"mi/in/oi/prs")
	end if

	if oObjNodes.length <> 0 then
		call RenderSiteChildren(siteID,oObjNodes)
	end if

	'if siteID<>objectID then
	Response.Write "</TD>"
	Response.Write "</TR>"
	'end if

	set oDoc = nothing

	If Err.Number <> 0  then
		Response.Write "RenderSiteObjects::Error " & Err.Number & " " & Err.Description
		Response.Write "<BR>"
		Err.clear
	End If

End Function

Function RenderSiteChildren(siteID,oObjNodes)
On Error Resume Next
Dim oObjNode
Dim oAttrNode
Dim objectID
Dim objectType
Dim objectName

	Response.write "<TR>"
   	Response.Write "<TD COLSPAN=1>"
   	Response.Write "</TD>"
   	Response.Write "<TD>"
   	Response.Write "<Table border=1>"
   	Response.Write "<TR>"
   	Response.Write "<TD></TD>"
	Reponse.write "<TD>"
	For each oObjNode in oObjNodes
		objectID=""
		objectType=""
		objectName=""
		For each oAttrNode in oObjNode.attributes
			if oAttrNode.NodeName = "id" then objectID = oAttrNode.text
			if oAttrNode.NodeName = "tp" then objectType = oAttrNode.text
			if objectID <> "" and objectType <> "" then
				call RenderSiteObjects(siteID,objectID,objectType,objectName)
			end if
		Next
		Response.Write "<BR>"
	next
   	Response.Write "</TD>"
   	Response.Write "</TR>"
   	Response.Write "</TABLE>"
   	Response.Write "</TD>"
   	Response.Write "</TR>"
	If Err.Number <> 0  then
		Response.Write "RenderSiteChildren::Error " & Err.Number & " " & Err.Description
		Err.clear
	End If
End Function

Function RenderObjectProperties(siteID,objectID,propXML,propPath)
On Error Resume Next
Dim propNodes
Dim propsNode
Dim oAttrNode
Dim propNode
Dim id
Dim value
Dim found
Dim oDoc

	Set oDoc = Server.CreateObject("Microsoft.XMLDOM")
	oDoc.loadXML(propXML)
	Set propNodes = oDoc.selectNodes(propPath)

	Response.Write "<Table  border=1>"
	id=""
	For each propsNode in propNodes
		For each propNode in propsNode.childNodes
			Response.Write "<TR>"
			found=false
			For each oAttrNode in propNode.attributes
				Response.Write "<TD>"
				if oAttrNode.NodeName="id" and oAttrNode.Text="STORAGE_MAPPING" then found=true
				if found=true and oAttrNode.NodeName="v" then id = oAttrNode.Text
				Response.Write oAttrNode.Text
				Response.Write "</TD>"
			Next
			Response.Write "</TR>"
		Next
	next
	Response.Write "</Table>"

	if id <> "" then
		call RenderObjectdefinition(siteID,id,"1012")
	end if

	set oDoc=nothing
	If Err.Number <> 0  then
		Response.Write "RenderObjectProperties::Error " & Err.Number & " " & Err.Description
		Err.clear
	End If
End Function

Function getObjectProperties(siteID,ObjectID)
On Error Resume Next
	Dim oSiteInfo
	Dim sitePropsXML

	set oSiteInfo = Server.CreateObject("Bridge2API.SiteInfo")
	sitePropsXML = oSiteInfo.getObjectProperties(siteID,objectID)

	If Err.Number <> 0  then
		Response.Write "getObjectProperties::Error " & Err.Number & " " & Err.Description & " " & "while fetching site properties"
		Err.clear
	End If

	set oSiteInfo = nothing
	getObjectProperties = sitePropsXML
End Function


Function RenderObjectDefinition(siteID,objectID,objectType)
On Error Resume Next
Dim oSite
Dim objDefn
Dim deviceIDs(1)
Dim title
	deviceIDs(1) = objectID
	set oSite = Server.CreateObject("Bridge2API.SiteInfo")
	if objectType="1005" then
		objDefn = oSite.getDeviceTypeDefinitions(siteID,deviceIDs)
		title = "Device Type Definition"
	elseif objectType="1012" then
		objDefn = oSite.getMappingDefinition(siteID,objectID)
		title = "Mapping Definition"
	end if
	Response.Write "<BR>"
	Response.Write "<B>" & title & "</B>"
	Response.Write "<BR>"
	Response.Write Server.HTMLEncode(objDefn)
	If Err.Number <> 0  then
		Response.Write "getObjectProperties::Error " & Err.Number & " " & Err.Description & " " & "while fetching site properties"
		Err.clear
	End If

	set oSite = nothing
End Function

%>
