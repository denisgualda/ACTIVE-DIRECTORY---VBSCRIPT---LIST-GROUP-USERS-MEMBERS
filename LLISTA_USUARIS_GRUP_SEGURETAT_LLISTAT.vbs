'**************************************************
'**************************************************
'  LLISTA USUARIS DE GRUPS DE SEGURETAT D'UN FITXER DE TEXT

'	--> Llegir fitxer de text del mateix directori de l'script GRUP_TO_READ_LIST.txt	
'	--> Mostra el resultat en finestra d'IE.
'	--> Guarda restaultat en fitxer de text al mateix directori --> usuaris_grup.txt
'
'**************************************************
'**************************************************
wscript.echo "INDICA EL FITXER QUE CONTÉ ELS GRUPS DE SEGURETAT A LLEGIR PER LLISTAR ELS SEUS USUARIS:"

'SELECCIONA EL FITXER A LLEGIR:-------------------------------------------------------------------
Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
ruta_fitxer = oExec.StdOut.ReadLine		'GUARDA RUTA DE FITXER

'-------------------------------------------------------------------------------------------------

Const ForAppending = 8

''*************************************************************
'LLEGEIG FITXER
'*************************************************************
Dim objFSORead: Set objFSORead = CreateObject("Scripting.FileSystemObject")
dim filetoread
'LLISTAT DE PC'S AL FITXER TXT GUARDAT AL MATEIX DIRECTORI QUE L'SCRIPTS ---> COMPUTER_TO_READ_LIST.txt
'filetoread = left(wscript.scriptfullname,len(wscript.scriptfullname)-len(wscript.scriptname)) + "\GRUPS_A_LLEGIR.txt"
filetoread = ruta_fitxer 
wscript.echo filetoread
set objReadFile = objFSORead.OpenTextFile (filetoread,1)
'*************************************************************
'*************************************************************

'*************************************************************
'FINESTRA EMERGENT MOSATRADA PER PANTALLA MENTRE DURA L'ACCIÓ
'*************************************************************
'--------------------------------------
Set objExplorer = CreateObject ("InternetExplorer.Application")
objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Left = 200
objExplorer.Top = 200
objExplorer.Width = 800
objExplorer.Height = 800 
objExplorer.Visible = 1             
objExplorer.Document.Title = "LLISTAT USUARIS DE GRUP DE SEGURETAT "
'*************************************************************
'*************************************************************

'*************************************************************
'ESCRIU A FITXER
'*************************************************************
Dim strLogFile, strDate
Dim objFSO

strLogFile = left(wscript.scriptfullname,len(wscript.scriptfullname)-len(wscript.scriptname)) + "\VEURE_USUARIS_GRUPS.txt"
'strLogFile = "C:\usuaris_grup.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")
if objFSO.FileExists(strLogFile) Then
	objFSO.DeleteFile(strLogFile)
End if
Set objLogFile = objFSO.OpenTextFile(strLogFile, ForAppending, True)
'*************************************************************
'*************************************************************

'RECORRE FITXER PER OBTENIR GRUPS DE SEGURETAT
'----------------------------------------------------
Do While Not objReadFile.AtEndOfStream
	strLine = objReadFile.readline
	strGrup = strLine
	'------------------------------------------------
	'OBTENIR DISTINGUISHED NAME
	set objSystemInfo = CreateObject("ADSystemInfo")
	strDomain = objSystemInfo.DomainShortName
	grup = GetUserDN(strGrup,strDomain)
	dsgrup = "LDAP://" & grup
	'------------------------------------------------

	'--------------------------------------------------------------------
	'ESCRIUP PER PANTALLA GRUPS TROBATS


	'ESCRIU PER PANTALLA (CMD)
	printsc =  strGrup & VBCrlf
	printsc = printsc & "-------------------------------------"
	'wscript.echo printsc
	'********************************

	'ESCRIU PER PANTALLA IE
	log_pantalla = log_pantalla & "<b>" & strGrup & "</b><br>"
	log_pantalla = log_pantalla & "<b>-------------------------------------</b>" & "<br>"
	'objExplorer.Document.Body.InnerHTML = log_pantalla
	'*******************************

	'ESCRIU EN FITXER
	objLogFile.WriteLine printsc
	'*******************************


	'--------------------------------------------------------------------
	'--------------------------------------------------------------------




	'Set objGroup = GetObject("LDAP://CN=GL_PR_C3_IM5_BN,OU=Impressores,OU=Departaments,DC=girona,DC=gencat,DC=cat")
	Set objGroup = GetObject(dsgrup)
	'LLISTA USUARIS DEL GRUP
	For Each objMember In objGroup.Members
		printusers = objMember.sAMAccountName
		'wscript.echo printusers
		objLogFile.WriteLine printusers
		log_pantalla = log_pantalla & printusers & "<br>"
		objExplorer.Document.Body.InnerHTML = log_pantalla


	Next

		'wscript.echo vbnullstring
		'printsc = printsc & vbnullstring
		'SALT DE LINIA AL FINAL DE LECTURA DE CADA GRUP
		log_pantalla = log_pantalla & "<br>"
		objExplorer.Document.Body.InnerHTML = log_pantalla
		objLogFile.WriteLine vbnullstring

Loop		'fi recorre fitxer
'----------------------------------------------------


log_pantalla = log_pantalla & "<br>"
log_pantalla = log_pantalla & "<b>************************************************************************************* </b><br>"
log_pantalla = log_pantalla & " <b>S'han guardat els resultats al fitxer --> VEURE_USUARIS_GRUPS.txt  </b>"
log_pantalla = log_pantalla & "<b>************************************************************************************* </b><br>"
objExplorer.Document.Body.InnerHTML = log_pantalla
objLogFile.WriteLine vbnullstring

objReadFile.Close
objLogFile.Close

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