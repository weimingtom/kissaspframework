<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: Generate Html file
'	File Name	: GenHtml.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** dim myhtml
'*** set myhtml= new Kiss_GenHtml
'*** 'myhtml.foldename = "test" 
'*** 'myhtml.Filename = "ok.shtml"
'*** myhtml.Htmlstr = "<html><head><title>HTML file</title></head><body>Testing</body></html>"
'*** myhtml.Htmlmake
'*** set myhtml=nothing

'*************************************************************
'	Initialize the class
'*************************************************************


Class Kiss_GenHtml

	'************************************************************* 
	' Param: foldername as a folder name default [year month day]
	' Param: Filename as a folder name default [hour minute second]
	' Param: Htmlstr as a file content
	'************************************************************* 

	Private HtmlFolder,HtmlFilename,HtmlContent

	Public property let foldename(str)
		HtmlFolder=str
	End property

	Public property let Filename(str)
		HtmlFilename=str
	End property

	Public property let Htmlstr(str)
		HtmlContent=str
	End property

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Constructor
	'************************************************************* 
	Private Sub class_initialize()
		HtmlFolder=Datename1(now)
		HtmlFilename=Datename2(now)&".html"
		HtmlContent=""
		response.write filepath
	End Sub

	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	Private Sub class_terminate()
	End Sub

	'************************************************************* 
	' Name: Datename1
	' Purpose: File name convert to time
	'************************************************************* 
	Private Function Datename1(timestr)
		dim s_year,s_month,s_day
		s_year=year(timestr)
		if len(s_year)=2 then s_year="20"&s_year
		s_month=month(timestr)
		if s_month<10 then s_month="0"&s_month
		s_day=day(timestr)
		if s_day<10 then s_day="0"&s_day
		Datename1=s_year & s_month & s_day
	End Function

	Private Function Datename2(timestr)
		dim s_hour,s_minute,s_ss
		s_hour=hour(timestr)
		if s_hour<10 then s_hour="0"&s_hour
		s_minute=minute(timestr)
		if s_minute<10 then s_minute="0"&s_minute
		s_ss=second(timestr)
		if s_ss<10 then s_ss="0"&s_ss
		Datename2 = s_hour & s_minute & s_ss
	End Function

	'************************************************************* 
	' Name: Htmlmake
	' Purpose: Generate html
	'************************************************************* 
	Public Sub Htmlmake()
		'On Error Resume Next
		dim filepath,fso,fout
		filepath = HtmlFolder&"/"&HtmlFilename
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if not (fso.FolderExists(Server.Mappath(HtmlFolder))) Then
			fso.CreateFolder(Server.MapPath(HtmlFolder)) 'Create a folder
		End If
		Set fout = fso.CreateTextFile(Server.MapPath(filepath),true)
		fout.WriteLine HtmlContent 'Generate a html file
		fout.close
	End Sub

	'************************************************************* 
	' Name: getTemplate
	' Purpose: Get the template
	'************************************************************* 
	public Function getTemplate(template)
		Dim fso,f
		set fso=CreateObject("Scripting.FileSystemObject")
		set f = fso.OpenTextFile(template)
		getTemplate=f.ReadAll
		f.close
		set f=nothing
		set fso=Nothing
	End Function 

	'************************************************************* 
	' Name: Htmldel
	' Purpose: Delete the html file
	'************************************************************* 
	Public Sub Htmldel()
		dim filepath,fso
		filepath = HtmlFolder&"/"&HtmlFilename
		Set fso = CreateObject("Scripting.FileSystemObject")
		fso.DeleteFile(Server.mappath(filepath))
		Set fso = nothing
	End Sub

End class
%>
