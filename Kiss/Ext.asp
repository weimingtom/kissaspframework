<%
if APP_ = "" or isNull(APP_) then 
	'show404("No direct script access allowed")
end if
'*************************************************************
'	class		: A Extensive function Library
'	File Name	: Ext.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'************************************************************* 
' Extensive String
'************************************************************* 

'************************************************************* 
' Name: print
' Param: str as a output string
' Purpose: Print the value of a variable
'************************************************************* 
function print(str)
	response.write str
end function

'************************************************************* 
' Name: echo
' Param: str as a output string
' Purpose: Print the value of a variable
'************************************************************* 
function echo(str)
	print str
end function

'************************************************************* 
' Name: die
' Param: str as a output string
' Purpose: Print the value of a variable and exit the procedure
'************************************************************* 
function die(str)
	print str
	response.end
end function

'************************************************************* 
' Name: show404
' Param: message as a output string
' Purpose: Show the 404 error page
'************************************************************* 
sub show404(message)
	die "<html><head><title>Exception page</title><meta http-equiv=""Content-Type"" content=""text/html; charset=gbk"" /><style type=""text/css""><!--" & vbNewLine & "* { margin:0; padding:0 }" & vbNewLine & "body { background:#333; color:#0f0; font:14px/1.6em ""宋体"", Verdana, Arial, Helvetica, sans-serif; }" & vbNewLine & "dl { margin:20px 40px; padding:20px; border:3px solid #f63; }" & vbNewLine & "dt { margin:0 0 0.8em 0; font-weight:bold; font-size:1.6em; }" & vbNewLine & "dd { margin-left:2em; margin-top:0.2em; }" & vbNewLine & "--></style></head><body><div id=""container""><dl><dt>Description:</dt><dd><span style=""color:#ff0;font-weight:bold;font-size:1.2em;"">Position:</span> " & message & "</dd></dl></div></body></html>"
end sub

'************************************************************* 
' Name: throwException
' Param: message as a output string
' Purpose: Throw a exception
'************************************************************* 
sub throwException(message)
	dim tmp
	tmp = message & "<p><span style=""color:#ff0;font-weight:bold;font-size:1.2em;"">Error:</span> " & err.number & " " & err.description & "</p>"
	show404(tmp)
	err.clear
end sub

'************************************************************* 
' Name: substr
' Param: str as the input string
' Purpose: substr — Return part of a string
'************************************************************* 
function substr(str, start, length)
	substr = mid(str, start, length)
end function

'************************************************************* 
' Name: strtolower
' Param: str as the input string
' Purpose: strtolower — Make a string lowercase
'************************************************************* 
function strtolower(str)
	strtolower = lcase(str)
end function

'************************************************************* 
' Name: strtoupper
' Param: str as the input string
' Purpose: strtoupper — Make a string uppercase
'************************************************************* 
function strtoupper(str)
	strtoupper = ucase(str)
end function

'************************************************************* 
' Name: ucfirst
' Param: str as the input string
' Purpose: ucfirst — Make a string's first character uppercase
'************************************************************* 
function ucfirst(str)
	dim tmp : tmp = ""
	dim tmp2 : tmp2 = ""
	tmp = substr(str, 1, 1)
	tmp = strtoupper(tmp)
	tmp2 = substr(str, 2, len(str))
	ucfirst = tmp & tmp2
end function

'************************************************************* 
' Name: randStr
' Purpose: Generate a specific length random string
'************************************************************* 
public function randStr(StrLen)
	dim tmpStr : tmpStr = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
	dim i, ti
	randStr = ""
	for i = 1 To StrLen
		ti = 0
		while ti = 0
			ti = CInt(rand(1,62))
			randStr = randStr & subStr(tmpStr, ti, 1)
		wend
		randStr = randStr & Mid(r, ti, 1)
	next
end function

'************************************************************* 
' Name: rand
' Param: str as the input string
' Purpose: rand — Generate a random integer
'************************************************************* 
function rand(min, max)
	randomize
	rand = int((max - min + 1) * rnd + min)
end function

'************************************************************* 
' Name: eval
' Param: str as the input string
' Purpose: eval — execute asp source code
'************************************************************* 
sub eval(str)
	execute str
end sub

'************************************************************* 
' Name: evalglobal
' Param: str as the input string
' Purpose: evalglobal — execute asp source code entirely
'************************************************************* 
sub evalglobal(str)
	executeglobal str
end sub

'************************************************************* 
' Name: htmlEncode
' Param: str as the input string
' Purpose: htmlEncode — filter html code
'************************************************************* 
function htmlEncode(str)
	if Trim(Str)="" Or IsNull(str) then
		htmlEncode=""
	else
		str=Replace(str,">",">")
		str=Replace(str,"<","<")
		str=Replace(str,Chr(32)," ")
		str=Replace(str,Chr(9)," ")
		str=Replace(str,Chr(34),"""")
		str=Replace(str,Chr(39),"'")
		str=Replace(str,Chr(13),"")
		str=Replace(str,Chr(10) & Chr(10), "</p><p>")
		str=Replace(str,Chr(10),"<br> ")
		htmlEncode=str
	end if
end function

'************************************************************* 
' Name: strLen
' Param: str as the input string
' Purpose: strLen — Get the string length
'************************************************************* 
function strLen(Str)
 if Trim(Str)="" Or IsNull(str) then
  strlen=0
 else
  dim P_len,x
  P_len=0
  StrLen=0
  P_len=Len(Trim(Str))
  for x=1 To P_len
   if Asc(Mid(Str,x,1))<0 then
	StrLen=Int(StrLen) + 2
   else
	StrLen=Int(StrLen) + 1
   end if
  next
 end if
end function

'************************************************************* 
' Name: getLenStr
' Param: str as the input string
' Purpose: getLenStr — Get the string of specific length
'************************************************************* 
function getLenStr(text,length) 
	dim i, text_length, word_len, tmp
	tmp = trim(text)  
	text_length = len(tmp) 
	word_len = 0
	if text_length >= 1 then  
		for i = 1 to text_length 
			if asc(mid(tmp,i,1)) < 0 or asc(mid(tmp,i,1)) >255 then 'if chinese word  
				word_len = word_len + 2  
			else  
				word_len = word_len + 1  
			end if  
			if word_len >= length then  
				tmp = left(trim(tmp),i) 'length of string  
				exit for  
			end if  
		next  
		getLenStr = tmp  
	else  
		getLenStr = ""  
	end if  
end function

'************************************************************* 
' Name: import
' Param: pPath as the class path such as "app.model.news"
' Purpose: import — import a asp class
'************************************************************* 
sub import(pPath)
	dim str_path, require, tmp, tmpA
	set require = new Kiss_require
	if pPath <> "" then
		if inStr(pPath, ".") = 0 then
			str_path = "Kiss/" & ucfirst(pPath) & ".asp"
		else
			tmpA = split(pPath, ".")
			tmp = ucfirst(tmpA(uBound(tmpA)))
			str_path = join(array_pop(tmpA), "/") & "/" & tmp & ".asp" 
			'str_path = join(tmpA, "/") & "/" & tmp & ".asp" 
		end if
	end if
	require(str_path)
	set require = nothing
end sub	

'************************************************************* 
' Name: function_exists
' Param: pName as a function name
' Purpose: function_exists — Check if the function of a class exists
'************************************************************* 
function function_exists(className, pName)
	dim file, content, tmpPath
	tmpPath = APP_CONTROLLER & className & ".asp"
	set file = new Kiss_File
	content = file.readFile(tmpPath)
	'die content
	if inStr(content, "function action" & ucfirst(pName)) = 0 and inStr(content, "sub action" & ucfirst(pName)) = 0 then
		function_exists = false
	else
		function_exists = true
	end if
	set file = nothing
end function

'************************************************************* 
' Name: fileExists
' Param: inFile as file name
' Purpose: Check existence of the file
'************************************************************* 
function fileExists(inFile)
	dim objFSO
	set objFSO = createobject("Scripting.FileSystemObject")
	if objFSO.FileExists(Server.Mappath(inFile)) then FileExists = true
	set objFSO = nothing
end function

'************************************************************* 
' Name: isNothing
' Param: Obj as a object
' Purpose: isNothing — Check if the object is nothing
'************************************************************* 
function isNothing(Obj)
 if not isObject(Obj) then
	isNothing = true
	exit function
 end if  
 if Obj is Nothing then
	isNothing=true
	exit function
 end if
 if isNull(Obj) then
	IsNothing=true
	exit function
 end if    
 isNothing=false
end function


'************************************************************* 
' Extensive Array
'************************************************************* 

'************************************************************* 
' Name: print_r
' Param: arr as a output array
' Purpose: Print all the element from index of an array
'************************************************************* 
function print_r(arr)
	if isarray(arr) then
		dim tmp : tmp = Join(arr, ",")
		response.write tmp
	else
		die("The input is not a array.")
	end if
end function

'************************************************************* 
' Name: toString
' Param: arr as an array converted to a string
' Purpose: Convert to a String
' Remarks: dim a : a = array(1,3,5)
'		   echo(toString(a))
'************************************************************* 
function toString(arr)
	if isArray(arr) then
		dim tmp : tmp = Join(arr, ",")
		toString = tmp
	else
		toString = arr
	end if
end function

'************************************************************* 
' Name: toArray
' Param: str as a string converted to an array
' Purpose: Convert to an array
' Remarks: dim a : a = "a, b, c"
'		   print_r(toArray(a))
'************************************************************* 
function toArray(str)
	dim tmp : tmp = split(str, ",")
	toArray = tmp
end function

'************************************************************* 
' Name: array_push
' Param: arr as an Array
' Param: str as a variable added to an array
' Purpose: array_push — Push one element onto the end of array
'************************************************************* 
function array_push(arr, var)
	dim tmp : tmp = toString(arr)
	tmp = tmp & "," & var
	tmp = toArray(tmp)
	array_push = tmp
end function

'************************************************************* 
' Name: array_strip
' Param: str as a string such as "1,2,3,"
' Purpose: Strip "," of string
'************************************************************* 
function array_strip(str)
	dim tmp : tmp = ""
	if left(str, 1) = "," then tmp = right(str,Len(str)-1)
	if right(str, 1) = "," then tmp = Left(str,Len(str)-1)
	tmp = toArray(tmp)
	array_strip = tmp
end function

'************************************************************* 
' Name: array_filter
' Param: arr as an Array
' Param: callback as callback function
' Purpose: array_filter — Filters elements of an array using a callback function
'************************************************************* 
function array_filter(arr, callback)
	dim e : e = ""
	dim tmp : tmp = ""
	for each e in arr
		execute ("tmp=tmp&" & callback & "(" & e & ")" & "&"",""")
	next
	tmp = array_strip(tmp)
	array_filter = tmp
end function

'************************************************************* 
' Name: array_shift
' Param: arr as an Array
' Purpose: array_shift — Shift an element off the beginning of array
'************************************************************* 
function array_shift(arr)
	dim i, tmp : tmp = ""
	for i=0 to Ubound(arr)
		 if i<>0 then tmp=tmp & arr(i) & ","
	next
	tmp = array_strip(tmp)
	array_shift = tmp
end function

'************************************************************* 
' Name: array_unshift
' Param: arr as an Array
' Purpose: array_unshift — Prepend one element to the beginning of an array
'************************************************************* 
function array_unshift(arr, var)
	dim i, tmp : tmp = ""
	tmp = toString(arr)
	tmp = var & "," & tmp
	tmp = toArray(tmp)
	array_unshift = tmp
end function

'************************************************************* 
' Name: array_pop
' Param: arr as an array
' Purpose: array_pop — Pop the element off the end of array
'************************************************************* 
function array_pop(arr)
	dim i, tmp : tmp = ""
	for i=0 to Ubound(arr)
		 if i<>Ubound(arr) then tmp=tmp & arr(i) & ","
	next
	tmp = array_strip(tmp)
	array_pop = tmp
end function

'************************************************************* 
' Name: array_splice
' Param: arr as an array
' Purpose: array_splice — Remove a portion of the array and replace it with something else
'************************************************************* 
function array_splice(arr, start, final)
	dim i, temp, tmp : tmp = ""
	if start > final then
		temp = start
		start = final
		final = temp
	end if
	for i=0 to Ubound(arr)
		 if i < (start-1) or i > (final-1) then tmp=tmp & arr(i) & ","
	next
	tmp = array_strip(tmp)
	array_splice = tmp
end function

'************************************************************* 
' Name: array_fill
' Param: arr as a Array
' Param: index as index to insert into an array
' Param: value as element to insert into an array
' Purpose: array_fill — Fill an array with value
'************************************************************* 
function array_fill(arr, index, value)
	dim i, tmp : tmp = ""
	for i=0 to Ubound(arr)
		if i <> (index-1) then
			tmp = tmp & arr(i) & ","
		else
			tmp = tmp & value & "," & arr(i) & ","
		end if
	next
	tmp = array_strip(tmp)
	array_fill = tmp
end function

'************************************************************* 
' Name: array_unique
' Param: arr as a Array
' Purpose: array_unique — Removes duplicate values from an array
'************************************************************* 
function array_unique(arr)
	dim tmp : tmp = ""
	for each e in arr
		if inStr(1, tmp, e) = 0 then
			tmp = tmp & e & ","
		end if
	next
	tmp = array_strip(tmp)
	array_unique = tmp
end function

'************************************************************* 
' Name: array_reverse
' Param: arr as a Array
' Purpose: array_reverse — Return an array with elements in reverse order
'************************************************************* 
function array_reverse(arr)
	dim tmp : tmp = ""
	for each e in arr
		tmp = tmp & e & ","
	next
	tmp = strReverse(tmp)
	tmp = array_strip(tmp)
	array_reverse = tmp
end function

'************************************************************* 
' Name: array_search
' Param: arr as a Array
' Param: value as searching value
' Purpose: array_search — searchs the array of a given value
'		and return a corresponding key if successful
'************************************************************* 
function array_search(arr, value)
	dim i
	for i=0 to Ubound(arr)
		if arr(i) = value then
			array_search = i
			exit function
		end if
	next
	array_search = false
end function

'************************************************************* 
' Name: array_rand
' Param: arr as a Array
' Param: num as specifies how many entries you want to pick 
' Purpose: array_rand — pick one or more random entries out of an array
'************************************************************* 
function array_rand(arr, num)
	dim tmpi, tmp : tmp = ""
	for i = 0 to num-1
		tmpi = rand(0, uBound(arr))
		tmp = tmp & arr(tmpi) & ","
	next
	tmp = array_strip(tmp)
	array_rand = tmp
end function

'************************************************************* 
' Name: sort
' Param: arr as a Array
' Purpose: sort — This function sorts an array. Elements will
'		   be arranged from lowest to highest when this 
'		   function has completed.
'************************************************************* 
function sort(arr)
	dim tmp, i, j
	redim tmpA(Ubound(arr))
	for i=0 to Ubound(tmpA)
		 tmpA(i)=CDBL(arr(i))
	next
	for i = 0 to Ubound(tmpA)
		for j = i+1 to Ubound(tmpA)
			if tmpA(i) > tmpA(j) then
				tmp      = tmpA(i)
				tmpA(i)  = tmpA(j)
				tmpA(j)  = tmp
			end if
		next
	next
	sort = tmpA
end function

'************************************************************* 
' Name: max
' Param: arr as a Array
' Purpose: max — Find highest value
'************************************************************* 
function max(arr)
	dim tmp : tmp =""
	tmp = array_reverse(sort(arr))
	max = tmp(0)
end function

'************************************************************* 
' Name: min
' Param: arr as a Array
' Purpose: min — Find lowest value
'************************************************************* 
function min(arr)
	dim tmp : tmp =""
	tmp = sort(arr)
	min = tmp(0)
end function


'************************************************************* 
' URL Access 2007-10-20
'************************************************************* 

'************************************************************* 
' Name: rGet
' Param: qStr as a http get name
' Param: filterType as a filter type
' Purpose: get the http get value and filter all the value
'************************************************************* 
function rGet(qStr, filterType)
	qStr = trim(qStr)

	if request.queryString(qStr) <> "" then
		rGet = request.queryString(qStr)
	else
		exit function
	end if

	rGet = rFilter(rGet, filterType)
end function

'************************************************************* 
' Name: rPost
' Param: qStr as a post form name
' Param: filterType as a filter type
' Purpose: get the post form value and filter all the value
'************************************************************* 
function rPost(qStr, filterType)
	qStr = trim(qStr)

	if request.form(qStr) <> "" then
		rPost = request.form(qStr)
	else
		exit function
	end if

	rPost = rFilter(rPost, filterType)
end function

'************************************************************* 
' Name: redirect
' Param: url as a URL
' Param: delay as a time delay
' Purpose: redirect the URL
'************************************************************* 
sub redirect(url,delay)
	delay = Int(delay)
	if delay = 0 then response.redirect(url)
	die("<html><head><meta http-equiv=""refresh"" content=""" &delay& ";URL=" &url& """ /></head></html>")
end sub

'************************************************************* 
' Name: getIP
' Purpose: Get the current IP
'************************************************************* 
function getIP()
	getIP = ""
	getIP = Request.ServerVariables("REMOTE_HOST")
end function

'************************************************************* 
' Name: getScriptName
' Purpose: Get the current URL path
'************************************************************* 
function getScriptName()
	getScriptName = Request.ServerVariables("SCRIPT_NAME")
end function

'************************************************************* 
' Name: getUrl
' Purpose: Get the current file name
'************************************************************* 
function getSelfName()
	getSelfName = Request.ServerVariables("PATH_TRANSLATED")
	getSelfName = Mid(getSelfName, InstrRev(getSelfName,"\") + 1, Len(getSelfName))
end function

'************************************************************* 
' Name: getExt
' Param: filename as a file name
' Purpose: Get the suffix of a file
'************************************************************* 
function getExt(filename)
    getExt = Mid(filename,InstrRev(filename,".")+1)
end function

'************************************************************* 
' Name: getUrl
' Purpose: Get the current URL with port
'************************************************************* 
function getUrl()
	if request.Servervariables("Server_PORT") = "80" then
		getUrl = "http://" & request.Servervariables("Server_name")&Replace(LCase(request.Servervariables("script_name")),getScriptName,"")
	else
		getUrl = "http://" & request.Servervariables("Server_name") & ":" & request.Servervariables("Server_PORT")&Replace(LCase(request.Servervariables("script_name")),getScriptName,"")
	end if
end function


'************************************************************* 
' Name: filterWords
' Param: str as the input string
' Purpose: filterWords — Filter the bad words
'************************************************************* 
function filterWords(fString)
    dim BadWords,bwords,i
    BadWords = "我操|操你|操他|你妈的|他妈的|狗|杂种|屄|屌|王八|强奸|做爱|处女|泽民|法轮|法伦|洪志|法輪"
    if not(IsNull(BadWords) or IsNull(fString)) Then
    bwords = Split(BadWords, "|")
    for i = 0 to UBound(bwords)
        fString = Replace(fString, bwords(i), string(Len(bwords(i)),"*"))
    next
    filterWords = fString
    end if
end function

'************************************************************* 
' Name: unAllowPost
' Purpose: unAllowPost — unAllow the outside post
'************************************************************* 
function unAllowPost()
    dim URL1,URL2
    unAllowPost = False
    URL1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
    URL2 = Cstr(Request.ServerVariables("SERVER_NAME"))
    if Mid(URL1,8,Len(URL2))<>URL2 then
        unAllowPost = false
    else
        unAllowPost = true
    end if
end function

'************************************************************* 
' Name: HTMLEncode
' Param: str as the input string
' Purpose: HTMLEncode — Encode the html tag
'************************************************************* 
function HTMLEncode(fString)
	if Not IsNull(fString) And fString <> "" Then
		fString = Replace(fString, "&", "&amp;")
		fString = Replace(fString, ">", "&gt;")
		fString = Replace(fString, "<", "&lt;")
		fString = Replace(fString, Chr(32), "&nbsp;")
		fString = Replace(fString, Chr(9), "&nbsp;&nbsp;")
		fString = Replace(fString, Chr(34), "&quot;")
		fString = Replace(fString, Chr(39), "&#39;")
		fString = Replace(fString, Chr(13), "")
		fString = Replace(fString, Chr(10) & Chr(10), "</p><p>")
		fString = Replace(fString, Chr(10), "<br />")
		fString = Replace(fString, Chr(255), "&nbsp;")
		HTMLEncode = fString
	end if
end function

'************************************************************* 
' Name: HTMLDecode
' Param: str as the input string
' Purpose: HTMLDecode — Decode the html tag
'************************************************************* 
function HTMLDecode(fString)
	if Not IsNull(fString) And fString <> "" Then
		fString = Replace(fString, "&amp;", "&")
		fString = Replace(fString, "&gt;", ">")
		fString = Replace(fString, "&lt;", "<")
		fString = Replace(fString, "&nbsp;", Chr(32))
		fString = Replace(fString, "&nbsp;&nbsp;", Chr(9))
		fString = Replace(fString, "&quot;", Chr(34))
		fString = Replace(fString, "&#39;", Chr(39))
		fString = Replace(fString, "", Chr(13))
		fString = Replace(fString, "</p><p>", Chr(10) & Chr(10))
		fString = Replace(fString, "<br />", Chr(10))
		fString = Replace(fString, "&nbsp;", Chr(255))
		HTMLDecode = fString
	end if
end function

'************************************************************* 
' Name: stripHTML
' Param: str as the input string
' Purpose: stripHTML — Strip the html tag
'************************************************************* 
function stripHTML(strHTML)
    Dim objRegExp,strOutput
    Set objRegExp = New Regexp
    objRegExp.IgnoreCase = true
    objRegExp.Global = true
    objRegExp.Pattern = "<.+?>"
    strOutput = objRegExp.Replace(strHTML,"")
    strOutput = Replace(strOutput, "<","&lt;")
    strOutput = Replace(strOutput, ">","&gt;")
    stripHTML = strOutput
    Set objRegExp = Nothing
end function

'************************************************************* 
' Name: filterSql
' Purpose: filterSql — Prevent sql injecting
'************************************************************* 
sub filterSql()
	dim inputWords, inputWordsArr, i, tpost, tget
	inputWords = "'|;|and|(|)|exec|insert|select|delete|update|count|chr|mid|master|truncate|char|declare"
	inputWordsArr = split(inputWords,"|")
	
	'post
	if Request.Form<>"" Then
		for each tpost in Request.Form
			for i = 0 to Ubound(inputWordsArr)
				if Instr(LCase(Request.Form(tpost)),inputWordsArr(i))<>0 then
					die("<script language=""javaScript"">alert(""Submit a unlawful argument. Forbid injecting sql."");</script>")
				end if
			next
		next
	end if
	
	'get
	if Request.QueryString<>"" then
		for each tget in Request.QueryString
			for i = 0 to Ubound(inputWordsArr)
				if Instr(LCase(Request.QueryString(tget)),inputWordsArr(i))<>0 then
					die("<script language=""javaScript"">alert(""Submit a unlawful argument. Forbid injecting sql."");</script>")
				end if
			next
		next
	end if
end sub

'************************************************************* 
' Name: CUrl
' Purpose: CUrl — Return current URL
'************************************************************* 
function CUrl()
	Domain_Name = LCase(Request.ServerVariables("Server_Name"))
	Page_Name = LCase(Request.ServerVariables("Script_Name"))
	Quary_Name = LCase(Request.ServerVariables("Quary_String"))
	if Quary_Name ="" then
		CUrl = "http://"&Domain_Name&Page_Name
	else
		CUrl = "http://"&Domain_Name&Page_Name&"?"&Quary_Name
	end if
end function

'************************************************************* 
' Name: sendMail
' Purpose: sendMail — Return current URL
'************************************************************* 
function sendMail(MailServerAddress,AddRecipient,subject,Body,Sender,MailFrom)
	'on error resume next
	dim JMail
	set JMail=Server.CreateObject("JMail.SMTPMail")
	if err.number<>0 then
		throwException("No install JMail companent.")
	end if
	JMail.Logging=true
	JMail.Charset="gbk"
	JMail.ContentType = "text/html"
	JMail.ServerAddress=MailServerAddress
	JMail.AddRecipient=AddRecipient
	JMail.subject=subject
	JMail.Body=MailBody
	JMail.Sender=Sender
	JMail.From = MailFrom
	JMail.Priority=1
	JMail.Execute 
	set JMail=nothing 
	if err.number<>0 then 
		throwException("Sending mail fails.")
	else
		sendMail="OK"
	end if
end function

'************************************************************* 
' Name: formatTime
' Param: DateTime as the input time
' Param: format as the formating type
' Purpose: formatTime — Format time
'************************************************************* 
function formatTime(DateTime,format) 
	select case format
	case "1"
		 formatTime=""&year(DateTime)&"年"&month(DateTime)&"月"&day(DateTime)&"日"
	case "2"
		 formatTime=""&month(DateTime)&"月"&day(DateTime)&"日"
	case "3" 
		 formatTime=""&year(DateTime)&"/"&month(DateTime)&"/"&day(DateTime)&""
	case "4"
		 formatTime=""&month(DateTime)&"/"&day(DateTime)&""
	case "5"
		 formatTime=""&month(DateTime)&"月"&day(DateTime)&"日"&formatDateTime(DateTime,4)&""
	case "6"
	   temp="周日,周一,周二,周三,周四,周五,周六"
	   temp=split(temp,",") 
	   formatTime=temp(Weekday(DateTime)-1)
	case else
	formatTime=DateTime
	end select
end function

'************************************************************* 
' Name: zodiac
' Param: birthday as a birthday
' Purpose: zodiac — Format time
'************************************************************* 
function zodiac(birthday)
	if IsDate(birthday) then
		birthyear=year(birthday)
		zodiacList=array("猴","鸡","狗","猪","鼠","牛","虎","兔","龙","蛇","马","羊")        
		zodiac=zodiacList(birthyear mod 12)
	end if
end function

'************************************************************* 
' Name: constellation
' Param: birthday as a birthday
' Purpose: constellation — Format time
'************************************************************* 
function constellation(birthday)
	if IsDate(birthday) then
		ConstellationMon=month(birthday)
		ConstellationDay=day(birthday)
		if Len(ConstellationMon)<2 then ConstellationMon="0"&ConstellationMon
		if Len(ConstellationDay)<2 then ConstellationDay="0"&ConstellationDay
		MyConstellation=ConstellationMon&ConstellationDay
		if MyConstellation < 0120 then
			constellation="<img src=images/Constellation/g.gif title='魔羯座 Capricorn'>"
		elseif MyConstellation < 0219 then
			constellation="<img src=images/Constellation/h.gif title='水瓶座 Aquarius'>"
		elseif MyConstellation < 0321 then
			constellation="<img src=images/Constellation/i.gif title='双鱼座 Pisces'>"
		elseif MyConstellation < 0420 then
			constellation="<img src=images/Constellation/^.gif title='白羊座 Aries'>"
		elseif MyConstellation < 0521 then
			constellation="<img src=images/Constellation/_.gif title='金牛座 Taurus'>"
		elseif MyConstellation < 0622 then
			constellation="<img src=images/Constellation/`.gif title='双子座 Gemini'>"
		elseif MyConstellation < 0723 then
			constellation="<img src=images/Constellation/a.gif title='巨蟹座 Cancer'>"
		elseif MyConstellation < 0823 then
			constellation="<img src=images/Constellation/b.gif title='狮子座 Leo'>"
		elseif MyConstellation < 0923 then
			constellation="<img src=images/Constellation/c.gif title='处女座 Virgo'>"
		elseif MyConstellation < 1024 then
			constellation="<img src=images/Constellation/d.gif title='天秤座 Libra'>"
		elseif MyConstellation < 1122 then
			constellation="<img src=images/Constellation/e.gif title='天蝎座 Scorpio'>"
		elseif MyConstellation < 1222 then
			constellation="<img src=images/Constellation/f.gif title='射手座 Sagittarius'>"
		elseif MyConstellation > 1221 then
			constellation="<img src=images/Constellation/g.gif title='魔羯座 Capricorn'>"
		end if
	end if
end function


'************************************************************* 
' Security Filter 2007-10-20
'************************************************************* 

'Filter Constant
Const fNo = 0
Const fNumber = 1
Const fInt = 1
Const fString = 2
Const fLCase = 3
Const fUCase = 4
Const fNumberAndString = 5
Const fAll = 6
Const fforSQL = 7
Const fReal = 8

function rFilter(qStr, filterType)
	dim vReg
	set vReg = New RegExp
	vReg.Global = true
	vReg.IgnoreCase = true

	Select Case filterType
		Case fforSQL	'0	' Prevent SQL inject
			rFilter = FL2(qStr,"['#%;]", vReg)
		Case fInt	'1	'Number
			rFilter = FL1(qStr, "^-?\d+$", vReg)
			if IsNull(rFilter) Or IsEmpty(rFilter) Or rFilter = "" then rFilter = 0
			rFilter = CLng(rFilter)
		Case fString	'2	'Letter
			rFilter = FL2(qStr, "[^a-zA-Z]", vReg)
		Case fLCase		'3	'Lowercase
			rFilter = FL2(qStr, "[^a-z]", vReg)
		Case fUCase		'4	'Uppercase
			rFilter = FL2(qStr, "[^A-Z]", vReg)
		Case fNumberAndString	'5	'number/Letter/underline
			rFilter = FL1(qStr, "\w+", vReg)
		Case fAll	'6	'number/Letter/all denotation
			rFilter = qStr
		Case fReal	'8	'float
			rFilter = FL1(qStr, "^-?\d+(\.\d*)?(e-?\d*)?$", vReg)
		Case else
			rFilter = FL1(qStr, FilterType, vReg)
	end Select
end function

function FL1(str, AllowStr, vReg)
	dim i,j,t,mStr, Rex

	vReg.Pattern = AllowStr
	if vReg.Test(str) then FL1 = str
end function

function FL2(str, NotAllowStr, vReg)
	dim i,mStr

	vReg.Pattern = NotAllowStr
	FL2 = vReg.Replace(str, "")
end function

%>