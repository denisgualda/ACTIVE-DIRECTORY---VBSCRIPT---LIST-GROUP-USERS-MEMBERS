'**************************************************
'**************************************************
'  LLISTA USUARIS DE GRUPS DE SEGURETAT
'
'**************************************************
'**************************************************

strGrup = inputbox("INTRODUEIX EL NOM DEL GRUP")

'--------------------------------------
''EXTREU A FITXER
'Const ForAppending = 8
'Dim strLogFile, strDate
'Dim objFSO
'
''strLogFile = left(wscript.scriptfullname,len(wscript.scriptfullname)-len(wscript.scriptname)) + "\usuaris_grup.txt"
'strLogFile = "C:\usuaris_grup.txt"
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'if objFSO.FileExists(strLogFile) Then
	'objFSO.DeleteFile(strLogFile)
'End if
'Set objLogFile = objFSO.OpenTextFile(strLogFile, ForAppending, True)
'--------------------------------------

'------------------------------------------------------------
'FINESTRA EMERGENT MOSATRADA PER PANTALLA MENTRE DURA L'ACCIÓ
'--------------------------------------
Set objExplorer = CreateObject ("InternetExplorer.Application")

objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Left = 200
objExplorer.Top = 200
objExplorer.Width = 200
objExplorer.Height = 800 
objExplorer.Visible = 1             

objExplorer.Document.Title = "LLISTAT USUARIS DE GRUP DE SEGURETAT " & strGrup

objExplorer.Document.Body.InnerHTML = log_pantalla 




'------------------------------------------------
'OBTENIR DISTINGUISHED NAME
set objSystemInfo = CreateObject("ADSystemInfo")
strDomain = objSystemInfo.DomainShortName
grup = GetUserDN(strGrup,strDomain)
dsgrup = "LDAP://" & grup
'wscript.echo dsgrup
'------------------------------------------------

wscript.echo "-------------------------------------------"
wscript.echo "Usuaris que pertanyen al grup: " & VBCrlf
wscript.echo ">> " & strGrup 
wscript.echo "-------------------------------------------"
wscript.echo vbnullstring
'Set objGroup = GetObject("LDAP://CN=GL_PR_C3_IM5_BN,OU=Impressores,OU=Departaments,DC=girona,DC=gencat,DC=cat")

Set objGroup = GetObject(dsgrup)
'LLISTA USUARIS DEL GRUP
For Each objMember In objGroup.Members
	Wscript.Echo objMember.sAMAccountName
	'objLogFile.WriteLine objMember.sAMAccountName
	log_pantalla = log_pantalla & objMember.sAMAccountName & "<br>"
	objExplorer.Document.Body.InnerHTML = log_pantalla

Next
wscript.echo "-------------------------------------------"

wscript.sleep 1000000


'*******************************************************************************************	
'FUNCIO OBTENIR DISTINGUESHED NAME
Function GetUserDN(byval strUserName,byval strDomain)

	Set objTrans = CreateObject("NameTranslate")
	objTrans.Init 1, strDomain
	objTrans.Set 3, strDomain & "\" & strUserName 
	strUserDN = objTrans.Get(1) 
	GetUserDN = strUserDN

End function
'*******************************************************************************************