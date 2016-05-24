<!-- #include file="../CustomLib/CommonLib.asp" -->
<%

	On Error Resume Next
	
	Dim bDebug
	Dim szPath
	Dim oFs
	Dim oTextFile
	Dim iCount
	Dim sTemp
	Dim lErrNumber
	Dim sLanguage
	Dim asDesc
	
	lErrNumber = NO_ERR
	
	Redim asDesc(NUMBER_OF_DESCRIPTORS)
	sLanguage = "1033"
	bDebug = Len(Request.QueryString("debug")) = 0
	
	sTemp = CStr(Application.Contents(CStr(sLanguage))(0))
	Err.Clear
	
	'Get the Path
	szPath = Server.MapPath("Internationalization/MGNCS_" & sLanguage & ".txt")

	Set oFs = Server.CreateObject("Scripting.FileSystemObject")
	If(oFs.FileExists(szPath)) Then
		Set oTextFile = oFs.OpenTextFile(szPath)
	Else
		szPath = Server.MapPath("../Internationalization/MGNCS_" & sLanguage & ".txt")
		If (oFs.FileExists(szPath)) Then
			Set oTextFile = oFs.OpenTextFile(szPath)
		End If
	End If

	If(Not IsEmpty(oTextFile)) Then

		iCount = 0
		Do While oTextFile.AtEndOfStream <> True And Err.number = 0
		    If bDebug Then
		        asDesc(iCount) = "[" & iCount & "]" & oTextFile.ReadLine
		    Else
		        asDesc(iCount) = oTextFile.ReadLine
		    End If
		    iCount = iCount + 1
		Loop

		Application.Contents(sLanguage) = asDesc
    End If
    
	oTextFile.Close
	Application.Value("Languages_Loaded") = sLanguage & ";"
    Application.Contents(sLanguage) = asDesc

    Set oFs = Nothing
    Set oTextFile = Nothing
    Err.Clear
    

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY BGCOLOR="FFFFFF" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>

<FORM METHOD="GET">

    <% If bDebug Then %>
        The English strings were switched to debug mode. To turn it off, please press the button below:<p>
        <INPUT TYPE="HIDDEN" NAME="debug" VALUE="Turn Debug Off" >
        <INPUT TYPE="SUBMIT" VALUE="Turn Debug Off">
    <% Else %>
        The English strings were switched to normal mode. To debug them again, please press the button below:<P>
        <INPUT TYPE="SUBMIT" VALUE="Turn Debug On" >
    <% End If%>

    <p>
    <A HREF="default.asp">Back to admin</A> / <A HREF="../default.asp">Back to Portal</A>
    
</BODY>
</HTML>
