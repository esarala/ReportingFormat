
strFolderPath=GetFrameworkPath() & mid(Environment.Value("TestName"),1,4) & "_" & GenerateStringOfCurrentTime()
strResultFolderPath=strFolderPath&"\TestResult"
'--------------------------------------------------------------------------------------------
Call CreateFolder(strFolderPath)
Call CreateFolder(strResultFolderPath)

Function CreateFolder(strFolderPath) 
	Set oFolder = DotNetFactory.CreateInstance("System.IO.Directory")
	oFolder.CreateDirectory strFolderPath
End Function 
'---------------------------------------------------------------------------------------------
'HTML Report Header
Call fn_StartReporting(strResultFolderPath)

Public Function fn_StartReporting(strResultFolderPath)
	'strResultFolderPath=path
	'msgbox strResultFolderPath
	Call fn_CreateHeader(strResultFolderPath)
End Function


Public Function fn_CreateHeader(strResultFolderPath)
	'++++++++ Variable declaration +++++++++++++
	Dim strProjectName
	Dim objFSO, objFile
	Dim strHeaderFontColor,strHeadColor,strTableBg,strContentBGColor,strExecutionstartTime
	Dim strTCName, strHTMLResultName
	Dim objMyFile
	'msgbox strResultFolderPath
'	
	'++++++++ Variable Initilize+++++++++++
	strProjectName="CPC"
	strTCName=Environment.Value("TestName")
	
	Environment.Value("intStepNo")=1
	Environment.Value("intPassCount")=0
	Environment.Value("intFailCount")=0
	Environment.Value("strExecutionStartTime")=Time()
	
	'++++++++++++++++++ Call fn_CreateScreenShotDoc to create Word Doc which stores all screen shots+++++++++
	Call fn_CreateScreenShotDoc(strResultFolderPath)
	
	'+++++++++++++ Create Screen shot folder+++++++++++++
	'ScreenshotFolderPath=strResultFolderPath&"\Screenshots"
	'call CreateFolder(ScreenshotFolderPath)
	Call fn_CreateScreenshotFolderPath(strResultFolderPath)
		
	'+++++++++ Call funcation to create folder++++++++++++
	ResultFolderPath=strResultFolderPath
	strHTMLResultName=fn_CreateResultFilePath(ResultFolderPath)	
	
	'++++++++++++ Open text file+++++++++++++++
	Set objFSO= CreateObject("Scripting.FileSystemObject")	
	Set objMyFile=objFSO.OpenTextFile(strHTMLResultName,8)
	
	'+++++++++++ Open File to write +++++++++++
	'############## Create Header#############
	strHeaderFontColor="#680E0E"
	strHeadColor = "#F6F3E4"
	strTableBg = "#006699"
	strContentBGColor = "#FFFFFF"
	
	objMyFile.Writeline("<html>")
	objMyFile.Writeline("<head>")
		objMyFile.Writeline("<meta http-equiv=" & "Content-Language" & "content=" & "en-us>")
		objMyFile.Writeline("<meta http-equiv="& "Content-Type" & "content=" & "text/html; charset=windows-1252" & ">")
		objMyFile.Writeline("<title> Test Case Automation Execution Results</title>")
		objMyFile.Writeline("<script>")
				objMyFile.Writeline("top.window.moveTo(0, 0);")
				objMyFile.Writeline("window.resizeTo(screen.availwidth, screen.availheight);")
		objMyFile.Writeline("</script>")

	objMyFile.Writeline("</head>")
			
	objMyFile.Writeline("<Table><tr></tr></table>")
			
	objMyFile.Writeline("<style>table, th, td { border: 2px solid black; border-collapse: collapse;}th, td {  padding: 3px;}</style>")
		
	objMyFile.Writeline("<body bgcolor = #FFFFFF>")
	objMyFile.Writeline("<blockquote>")
	objMyFile.Writeline("<p align = center><table border=1 bordercolor=" & "#000000 id=table1 width=1000 height=35 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
	'+++++++++++ Header Name++++++++++++++++++
	objMyFile.Writeline("<tr>")
		objMyFile.Writeline("<td COLSPAN = 50 bgcolor ="& strHeadColor & ">")
				objMyFile.WriteLine("<p align=center><font color=#680E0E size=6 face= Copperplate Gothic Bold >Automation Execution Report </font></b></p>")
				'objMyFile.Writeline("<p align=center><b><font color="& strHeaderFontColor &"size=6 face= "& chr(50)&"Copperplate Gothic Bold"&chr(50) & "><B>Automation Execution Results<b></font></p>")
		objMyFile.Writeline("</td>")
	objMyFile.Writeline("</tr>")
					
	'++++++++++++++++ Sub Header details Project Details++++++++++++++
	objMyFile.Writeline("<tr>")
		objMyFile.Writeline("<td COLSPAN = 6 bgcolor ="& strHeadColor & ">")
				objMyFile.Writeline("<p align=center><font color="&strHeaderFontColor&"size=3 face= "& chr(34)&"Calibri Bold"&chr(34) & ">&nbsp;<b>Project Name: </b> </font>")
				objMyFile.Writeline("<font color="&strHeaderFontColor& "size=3 face= "& chr(34)&"Calibri Bold"&chr(34) & ">"&strProjectName&" </font></p>")
		objMyFile.Writeline("</td>")
					
					
	'++++++++++++++++ Sub Header details TC Details++++++++++++++					
		objMyFile.Writeline("<td COLSPAN = 6 bgcolor ="& strHeadColor & ">")
				objMyFile.Writeline("<p align=center><font color="&strHeaderFontColor&"size=3 face= "& chr(34)&"Calibri Bold"&chr(34) & ">&nbsp;<b>Test Case Name:  </b></font>")
				objMyFile.Writeline("<font color="&strHeaderFontColor&" size=3 face= "& chr(34)&"Calibri Bold"&chr(34) & ">"&strTCName&" </font></p>")
		objMyFile.Writeline("</td>")						
					
	'++++++++++++++++ Sub Header details Date Details++++++++++++++		
	strExecutionstartTime=time()
		objMyFile.Writeline("<td COLSPAN = 6 bgcolor ="& strHeadColor & ">")
				objMyFile.Writeline("<p align=center><font color="&strHeaderFontColor&" size=3 face= "& chr(34)&"Calibri Bold"&chr(34) & ">&nbsp;<b>Date/Time:  </b></font>")
				objMyFile.Writeline("<font color="&strHeaderFontColor&" size=3 face= "& chr(34)&"Calibri Bold"&chr(34) & ">"&date() & "-" &strExecutionstartTime&" </font></p>")
		objMyFile.Writeline("</td>")
	objMyFile.Writeline("</tr>")
	objMyFile.Writeline("</table>")
	
	'++++++++++++++++ Execution header +++++++++++++++++++
	objMyFile.Writeline("<p align = center>")
	objMyFile.Writeline("<table border=1 bordercolor=" & "#000000 id=table1 width=1000 height=35 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
	objMyFile.Writeline("<tr bgcolor="&strTableBg&">")
	objMyFile.Writeline("<td width=100 <p align= center><b><font color = white face=Arial size=2><b>Step #</b> </td></p>")
	objMyFile.Writeline("<td width=400 <p align= center><b><font color = white face=Arial size=2><b>Step Name</b> </td></p>")
	objMyFile.Writeline("<td width=500 <p align= center><b><font color = white face=Arial size=2><b>Expected Result</b> </td></p>")
	objMyFile.Writeline("<td width=500 <p align= center><b><font color = white face=Arial size=2><b>Actual Result </b> </td></p>")
	objMyFile.Writeline("<td width=200 <p align= center><b><font color = white face=Arial size=2><b>Status</b> </td>")
'	objMyFile.Writeline("<td width=200 <p align= center><b><font color = white face=Arial size=2><b>Screenshot </b> </td></p>")
	objMyFile.Writeline("<td width=200 <p align= center><b><font color = white face=Arial size=2><b>Time </b> </td></p>")
	objMyFile.Writeline("</tr>")
	
	'+++++++++ Close File +++++++++++++++++				
	objMyFile.Close()
	
	'++++++++++ Release memory++++++++
	Set objMyFile=nothing
	Set objFSO=nothing
End Function


Function fn_CreateScreenShotDoc(strResultFolderPath)
	Dim objWord,objDoc,objFSO
	Dim strDate,strTimestamp
	Dim strDocFolder,strDocFile
	
	strDate=replace(Date(), "/", "")
	strTimestamp=strDate&"_"&Hour(time())&"_"&Minute(time())&"_"&Second(time()) 
	
	'strDocFolder=Environment.Value("ResultDir")&"_ScreeenShotDoc"
	'strDocFile=strDocFolder&"\"&"SS_"&Environment.Value("TestName")&strTimestamp&".doc"
	strDocFolder=strResultFolderPath&"\ScreeenShotDoc"
	strDocFile=strDocFolder&"\"&"SS_"&Environment.Value("TestName")&strTimestamp&".doc"
	
	Set objFSO= CreateObject("Scripting.FileSystemObject")	
	
	If objFSO.FolderExists(strDocFolder) Then
		'++++++++ create Doc file +++++++++++++++++++
		Set objWord=CreateObject("Word.Application")
		objWord.Documents.Add
		objWord.ActiveDocument.SaveAs strDocFile
		objWord.Quit
	else
		'++++++ Create Doc folder++++++++++++++
		objFSO.CreateFolder(strDocFolder)
		'++++++++ create Doc file +++++++++++++++++++
		Set objWord=CreateObject("Word.Application")
		objWord.Documents.Add
		objWord.ActiveDocument.SaveAs strDocFile
		objWord.Quit
	end	If
	
	'+++++++++++++save file name in Environment variable +++++++++++++++
	Environment.Value("strDocFileName")=strDocFile
	
	'++++++++++ Release memory+++++++++++++++
	Set objFSO=nothing
	Set objWord= nothing
End Function


Public Function fn_CreateScreenshotFolderPath(strResultFolderPath)
	'+++++++++++ Decalare variables ++++++++++++++
	Dim objFSO,objFile
	Dim  strScreenshotFolderPath
	Dim strDate, strTimestamp, strScreenshotFolder
		
	'++++++++++++ Initalize variables ++++++++++	
	'strDate=replace(Date(), "/", "")
	'strTimestamp=strDate&"_"&Hour(time())&"_"&Minute(time())&"_"&Second(time()) 	
	
	strScreenshotFolderPath=strResultFolderPath&"\ScreenShots"
		'strScreenshotFolderPath=ScreenshotFolderPath	
	
	Set objFSO= CreateObject("Scripting.FileSystemObject")	
	'++++++++++++++ Check folder is exist or not +++++++
	If not objFSO.FolderExists(strScreenshotFolderPath) Then
		'++++++++++ Create Test Script Specific Folder +++++++
		objFSO.CreateFolder(strScreenshotFolderPath)	
	End if
	
	'
	
	'+++++++++++++ store in environement variable +++++++++
	Environment.Value("strScreenshotFolderPath") =strScreenshotFolderPath
	'++++++++++++++ Relese Memory++++++++
	Set objFSO=nothing
	'Set objFile= nothing
End Function

Public Function fn_CreateResultFilePath(ResultFolderPath)
	'+++++++++++ Decalare variables ++++++++++++++
	Dim strDate, strTimestamp
	Dim objFSO, objFile
	Dim strResultFolderPath, strTCName, strHTMLResultName
	'++++++++++++ Initalize variables ++++++++++	
	strDate=Date()
	
	strResultFolderPath=ResultFolderPath
	'strResultFolderPath=Environment.Value("ResultDir")&"_HTML_Result"

	'++++++++++++ Create Postix of HTML result file ++++++++++++
	strDate=replace(Date(), "/", "")
	strTimestamp=strDate&"_"&Hour(time())&"_"&Minute(time())&"_"&Second(time()) 
	strTCName=Environment.Value("TestName")
	
	strHTMLResultName=strResultFolderPath&"\"&strTCName&"_"&strTimestamp&"."&"html"
	
	'+++++++++++++++ Save File Name in Environment variable ++++++++
	Environment.Value("strHTMLResultName")=strHTMLResultName
	
		
	Set objFSO= CreateObject("Scripting.FileSystemObject")	

	'+++++++ /if folder already exist++++++++++++++++
	if objFSO.FolderExists(strResultFolderPath) Then		
		'++++++++ create html file +++++++++++++++++++
		Set objFile=objFSO.CreateTextFile(strHTMLResultName,true)
	else
		'++++++ Create HTML_result folder++++++++++++++
		objFSO.CreateFolder(strResultFolderPath)
		'+++++++++ create html file +++++++++++++++++++
		Set objFile=objFSO.CreateTextFile(strHTMLResultName,true)
	end	If

	'+++++++++++++Retrun Value+++++++++++++++
	fn_CreateResultFilePath=strHTMLResultName
	
	'++++++++++++++ Relese Memory++++++++
	Set objFSO=nothing
	Set objFile= nothing
End Function


Public Function fn_WriteReport(strStepName,strExpectedResult,strActualResult,strStatus)
	On error resume next
	Dim strDate,strResultFolderPath,strScreenshotFolderPath,strTimestamp
	Dim objFSO, objMyFile
	Dim strTCName,strResultName, strHTMLResultName
	Dim strHeaderFontColor, strHeadColor, strTableBg, strContentBGColor
	Dim intPassCount,strStepName1, strScreenShot, intFailCount,intStepNo
	
	
	'+++++++++ Get HTML result folder path ++++++++++
'	strResultFolderPath=Environment.Value("strResultFolderPath")
	
	'++++++++++++++++ Get Screen shot folder path +++++++++++++++
	strScreenshotFolderPath=Environment.Value("strScreenshotFolderPath")

	strDate=replace(Date(), "/", "")
	strTimestamp=strDate&"_"&Hour(time())&"_"&Minute(time())&"_"&Second(time())

	'+++++++++++++ Get Current HTML result file path +++++++++++
	strHTMLResultName= Environment.Value("strHTMLResultName")
	intStepNo=Environment.Value("intStepNo")
	
	If len(intStepNo)=1 Then
		intStepNo="0"&intStepNo
	End If
	
	Set objFSO= CreateObject("Scripting.FileSystemObject")	
	'++++++++++++ Open text file+++++++++++++++
	Set objMyFile=objFSO.OpenTextFile(strHTMLResultName,8)
	
	
	'+++++++++++ Open File to write +++++++++++
	'############## Create Header#############
	
	strHeaderFontColor="#680E0E"
	strHeadColor = "#F6F3E4"
	strTableBg = "#006699"
	strContentBGColor = "#FFFFFF"
	
	Select Case strStatus
		Case "Pass"
			Reporter.ReportEvent micPass, strStepName, strActualResult
		Case "Fail"
			Reporter.ReportEvent micFail, strStepName, strActualResult
		Case "Done"
			Reporter.ReportEvent micDone, strStepName, strActualResult
		Case "Warning"
			Reporter.ReportEvent micWarning, strStepName, strActualResult
	End select 
	'++++++++++++ Write values of reporting				
	objMyFile.Writeline("<tr bgcolor= #FFFFFF >")
	objMyFile.Writeline("<td width=100 <p align= center><font color = Black face=Arial size=2> "&intStepNo&"</p></td>")
	objMyFile.Writeline("<td width=400 <p align= center><font color = Black face=Arial size=2>"&strStepName &"</p></td>")
	objMyFile.Writeline("<td width=500 <p align= center><font color = Black face=Arial size=2>"&strExpectedResult &"</p></td>")
	objMyFile.Writeline("<td width=500 <p align= center><font color = Black face=Arial size=2>"&strActualResult&" </p></td>")
	If lcase(trim(strStatus))=lcase("Pass") Then
			objMyFile.Writeline("<td width=200 <p align= center><b><font color= #008000 face=Arial size=2><b> "&strStatus&"</b> </p></td>")
			intPassCount=cint(intPassCount)+cint(1)
			Environment.Value("intPassCount")=Environment.Value("intPassCount")+1			
	ElseIf lcase(trim(strStatus))=lcase("Fail") Then
			objMyFile.Writeline("<td width=200 <p align= center><b><font color=#FF0000 face=Arial size=2><b>"&strStatus&"</b> </p></td>")
			intFailCount=cint(intFailCount)+1
			Environment.Value("intFailCount")=Environment.Value("intFailCount")+1				
	ElseIf lcase(trim(strStatus))=lcase("Done")  Then
			objMyFile.Writeline("<td width=200 <p align= center><b><font color=#9900cc face=Arial size=2><b>"&strStatus&"</b> </p></td>")			
	End If
	'++++++++ Create unquie screen shot Name++++++++
	'strStepName1=strStepName&"_"&strTimestamp		
	'+++++++++++++ Screenshot file path
	strScreenShot=strScreenshotFolderPath&"\"&"Step_"&intStepNo&"_"&strStepName&".png" 
	'+++++++ Capture screen shot++++++++++
	Desktop.CaptureBitmap strScreenShot, True

	'objMyFile.Writeline("<td width=200 <p align= center><font color = Black face=Arial size=2 ><a target=""_blank"" class=""anibutton"" href=""..\TempResults\"&strStepName &".png"& """><img class=""screen"">"&strStepName&"</a></p> </td>")
	
	'objMyFile.Writeline("<td width=200 <p align= center><font color = Black face=Arial size=2 ><a target=""_blank"" class=""anibutton"" href=""..\TempResultsScreenShot\"&strStepName1 &".png"& """>"&strStepName&"</a></p> </td>")
	'objMyFile.Writeline("<td width=200 <p align= center><font color = Black face=Arial size=2 ><a target=""_blank"" class=""anibutton"" href=""..\Screenshot\"&strStepName1 &".png"& """>"&strStepName&"</a></p> </td>")
	'objMyFile.Writeline("<td width=200 <p align= center><font color = Black face=Arial size=2 ><a target=""_blank"" class=""anibutton"" href=""..\"&strStepName1 &".png"& """>"&strStepName&"</a></p> </td>")
	
	objMyFile.Writeline("<td width=200 <p align= center><font color = Black face=Arial size=2>"&time()&"</p> </td>")
	objMyFile.Writeline("</tr>")
	
	'+++++++++++Copy same screenshot file at QC run folder as well to make available for QC+++++++++
	'call fn_CreateQCAttachementFolder(strScreenShot)
	
	'++++ Increase step counter+++++++++++++++++
	intStepNo=cint(Environment.Value("intStepNo"))+cint("1")
	Environment.Value("intStepNo")=intStepNo
	'+++++++++ Close File +++++++++++++++++				
	objMyFile.Close()
	
	'++++++++++ Release memory++++++++
	Set objMyFile=nothing
	Set objFSO=nothing

End Function


'---------------------------------------------------
End Reporting

Public Function fn_EndReporting()
	'++++++++++++ Write Footer of Result +++++++
	Call fn_WriteFooter()
	'+++++++++++++++ Create Screen shot document file ++++++
	call fn_AddScreenshotInDoc()
	' Upload results to ALM
	Call fn_UploadToALM(Environment.Value("strHTMLResultName"))	
	
End Function

Function fn_WriteFooter()
	'++++++++++ Variable Declaration +++++++++===
 	Dim objFSO, objMyFile
 	Dim strHTMLResultName
 	Dim strHeaderFontColor,strHeadColor,strTableBg,strContentBGColor
 	Dim intPassCount, intFailCount, intTotalStep
	Dim intStartHr, intStartMin, intStartSec, intEndHr, intEndMin, intEndSec, strExecutionEndTime, intTotalHr
	Dim intTotalMin, intTotalSec, strTotalTime
	Dim strStart,intH, intM, intS,strExecutionStartTime
	
	Set objFSO= CreateObject("Scripting.FileSystemObject")	
	'+++++++++++++ Get Current HTML result file path +++++++++++
	strHTMLResultName= Environment.Value("strHTMLResultName")
	
	'++++++++++++ Open text file+++++++++++++++
	Set objMyFile=objFSO.OpenTextFile(strHTMLResultName,8)
	
	'+++++++++++ Open File to write +++++++++++
	'############## Create Header#############
	
	strHeaderFontColor="#680E0E"
	strHeadColor = "#F6F3E4"
	strTableBg = "#006699"
	strContentBGColor = "#FFFFFF"
	
	'############### Footer section #########################	
	intPassCount=Environment.Value("intPassCount")
	intFailCount=Environment.Value("intFailCount")
	
	intTotalStep=intPassCount+intFailCount	
	
	'################################
	intStartHr=hour(Environment.Value("strExecutionStartTime"))
	intStartMin=Minute(Environment.Value("strExecutionStartTime"))
	intStartSec=Second(Environment.Value("strExecutionStartTime"))
	
	
	strExecutionEndTime=time()
	strExecutionStartTime=Environment.Value("strExecutionStartTime")
	
'	intEndHr=hour(time())
'	intEndMin=Minute(time())
'	intEndSec=Second(time())
'	'+++++++++++++Total Hour++++++++++++++++
'	intTotalHr=intEndHr-intStartHr
'	
'	'+++++++++++ Total Min+++++++++++
'	If intStartMin< intEndMin Then
'		intTotalMin=intEndMin-intStartMin
'	else
'		intTotalMin=intStartMin-intEndMin
'	End If
'	'++++++total sec+++++++++
'	If intStartSec< intEndSec Then
'		intTotalSec=intEndSec-intStartSec
'	else
'		intTotalSec=intStartSec-intEndSec
'	End If
'	
'	strTotalTime= intTotalHr&":"&intTotalMin&":"&intTotalSec
	
	'_____________________________ Time calculate ____________________________________
	
	 intTotalSec = Abs(DateDiff("S", strExecutionStartTime, strExecutionEndTime))         
        intTotalMin = intTotalSec \ 60 
	 intTotalHr = intTotalMin \ 60 	        
        intTotalMin = intTotalMin mod 60 
        intTotalSec = intTotalSec mod 60 
        intTotalHr   = intTotalHr   mod 24 

        if len(intTotalHr) = 1 then intTotalHr = "0" & hours 

        strTotalTime = RIGHT("00" & intTotalHr , 2) & ":" &RIGHT("00" & intTotalMin, 2) & ":" &RIGHT("00" & intTotalSec, 2) 
	'___________________________________________
	
	objMyFile.Writeline("</table>")
	objMyFile.Writeline("<table>")
	objMyFile.Writeline("<p align = center><table  border=1 bordercolor=" & "#000000 id=table1 width=1000 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
	objMyFile.Writeline("<tr bgcolor =#F6F3E4>")
		objMyFile.Writeline("<td colspan =2 class=pfhead width=250px>")
			objMyFile.Writeline("<p align=justify><b><font color=#680E0E size=2 face= Arial>&nbsp; Verification Points :</b>&nbsp;"&intTotalStep &" &nbsp;</td>")
		objMyFile.Writeline("<td colspan =2 class=pfhead width=250px>")
			objMyFile.Writeline("<p align=justify><b><font color=#680E0E size=2 face= Arial>&nbsp; Passed : </b>&nbsp;"& intPassCount&" &nbsp;</td>")
		objMyFile.Writeline("<td colspan =2 class=pfhead width=250px>")
			objMyFile.Writeline("<p align=justify><b><font color=#680E0E size=2 face= Arial>&nbsp; Failed :</b>&nbsp;"& intFailCount &"&nbsp;</td>")
		objMyFile.Writeline("<td colspan =2 class=pfhead width=250px>")
			objMyFile.Writeline("<p align=justify><b><font color=#680E0E size=2 face= Arial>&nbsp; Total Time :</b>&nbsp;"& strTotalTime&" &nbsp;</td>")			
	objMyFile.Writeline("</table>")
	objMyFile.Writeline("</Body></html>")
					
	'+++++++++ Close File +++++++++++++++++				
	objMyFile.Close()
	
	'++++++++++ Release memory++++++++
	Set objMyFile=nothing
	Set objFSO=nothing
	
	'
End Function


Function fn_AddScreenshotInDoc()
	'+++++++++ Variable declaration+++++++
	Dim strDocFile,strScreenshotFolderPath, strFileName
	Dim file
	Dim objWord, objDoc
	
	'+++++++ Initailize Variable+++++++++
	strDocFile=Environment.Value("strDocFileName")
	strScreenshotFolderPath=Environment.Value("strScreenshotFolderPath") 
	
	'++++++ Create object++++++++++
	Set objWord = CreateObject("Word.Application")
	set objDoc=objWord.Documents.Open(strDocFile)
	
	set objFSO = createobject("scripting.filesystemobject") 
	set objFolder = objFSO.getfolder(strScreenshotFolderPath)  
	
	'+++++++++++ Uplocad All Screen shot in document ++++++++
	For each file in objFolder.Files
		strFileName=file.Name
		objWord.Selection.TypeText strFileName ' Insert Name of screen shot
		'objWord.Selection.insertbreak  'Insert Break
		strFileName=strScreenshotFolderPath&"\"&strFileName
		'objWord.Selection.TypeText strFileName
		objWord.Selection.InlineShapes.AddPicture (strFileName)	
		objWord.Selection.insertbreak  'Insert Break
	Next
	
	'++++++++++++++++++ Save and close file +++++++++++++++
	objDoc.Save
	objWord.Quit
	'++++++++++ release memory +++++++
	Set objWord = Nothing
	Call fn_UploadToALM(strDocFile)
End Function


Function fn_UploadToALM(strFilePath) '
'++++++++ Variable Declaration +++++++++++++
Dim strResultFolderPath,strResultName,strHTMLResultName
Dim objFoldAttachments, objFoldAttachment, objFso

' '+++++++++++ Create ALM folder path +++++++++++++=
Set objFoldAttachments = QCUtil.CurrentRun.Attachments
'
Set objFoldAttachment = objFoldAttachments.AddItem(Null)
'
' '+++++++++++++++ Attach File ++++++++++++
objFoldAttachment.FileName = strFilePath
'
objFoldAttachment.Type = 1
objFoldAttachment.Post
''
'+++++++++ Release Memory ++++++++++
Set objFoldAttachments=nothing

End Function


Function fn_UploadToALM(strFilePath) '
'++++++++ Variable Declaration +++++++++++++
Dim strResultFolderPath,strResultName,strHTMLResultName
Dim objFoldAttachments, objFoldAttachment, objFso

' '+++++++++++ Create ALM folder path +++++++++++++=
Set objFoldAttachments = QCUtil.CurrentRun.Attachments
'
Set objFoldAttachment = objFoldAttachments.AddItem(Null)
'
' '+++++++++++++++ Attach File ++++++++++++
objFoldAttachment.FileName = strFilePath
'
objFoldAttachment.Type = 1
objFoldAttachment.Post
''
'+++++++++ Release Memory ++++++++++
Set objFoldAttachments=nothing

End Function


