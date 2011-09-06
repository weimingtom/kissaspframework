<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: Generate Html file
'	File Name	: GenHtml2.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** sub genHtml()
'*** 	dim i
'*** 	'i = code_int(request.querystring("i"))
'*** 	dim objsave : set objsave = new Kiss_GenHtml
'*** 	objsave.foldername = webdir & "web.intro"
'*** 	objsave.fileurl = "http://127.0.0.1/Kiss/?a=view"
'*** 	objsave.filename = "12345.htm"
'*** 	objsave.save
'*** 	if i = 0 then
'*** 	objsave.filename = "index.htm"
'*** 	objsave.save
'*** 	end if
'*** 	thismessage = "已生成静态文件！<a href=""" & webdir & "web.intro/" & i & ".htm"" target=""_blank"">浏览</a>"
'*** end sub
'*** 
'*** call genHtml

'*************************************************************
'	Initialize the class
'*************************************************************

class Kiss_GenHtml2
	public foldername, filename, fileurl
	public fso, stream, xmlhttp
	public thiserror
	private i, j

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Constructor
	'************************************************************* 
	private sub class_initialize()
		foldername = ""
		filename = "html.htm"
		set fso = server.createobject("scripting.filesystemobject")
		set stream = server.createobject("adodb.stream")
		dim xmldocarr(3)
		xmldocarr(0) = "msxml2.serverxmlhttp.6.0"
		xmldocarr(1) = "msxml2.serverxmlhttp.5.0"
		xmldocarr(2) = "msxml2.serverxmlhttp.4.0"
		xmldocarr(3) = "msxml2.serverxmlhttp.3.0"
		for i = 0 to ubound(xmldocarr)
			xmldocstr = xmldocarr(i)
			if isobj(xmldocstr) then exit for
		next
		erase xmldocarr
		if objerr(xmldocstr) then
			response.clear
			rwrite thiserror
			response.end
		end if
		set xmlhttp = server.createobject (xmldocstr)
	end sub

	private sub class_terminate()
		if isobject(fso) then set fso = nothing
		if isobject(stream) then set stream = nothing
		if isobject(xmlhttp) then set xmlhttp = nothing
	end sub

	private function sco(byval objstr)
		set sco = server.createobject (objstr)
	end function

	private function isobj(byval objstr)
		dim testobj
		on error resume next
		set testobj = server.createobject (objstr)
		if -2147221005 <> err then
			isobj = true
		else
			isobj = false
		end if
		set testobj = nothing
		err.clear
	end function

	private function objerr(byval objstr)
	objerr = false
	if not isobj(objstr) then
		thiserror = objstr & "Registering component fails."
		objerr = true
	end if
	end function

	' Generating html file
	public sub save()
		dim filepath
		dim binFileData
		filepath = server.mappath(foldername & "/" & filename)
		if not fso.folderexists(server.mappath(foldername)) and foldername <> "" then
			fso.createfolder server.mappath(foldername)
		end if
		xmlhttp.open "GET",fileurl,false
		xmlhttp.send()
		binFileData = xmlhttp.responseBody
		Stream.Type = 1
		Stream.Open()
		Stream.Write(binFileData)
		Stream.SaveToFile FilePath, 2
		Stream.Close()
	end sub
end class

%>