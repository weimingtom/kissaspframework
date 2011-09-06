<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: Kiss Debugging class
'	File Name	: Debug.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** dim debug, output : output="Just output for debugging" 
'*** set debug = new Kiss_Debug 
'*** debug.Enabled = true 
'*** debug.Print "Debugging", output 
'*** 
'*** debug.draw 
'*** set debug = nothing 

dim debug
set debug = New Kiss_Debug 
debug.Enabled = true 

class Kiss_Debug
	public className	'Class name
	private dbg_Enabled 
	private dbg_Show 
	private dbg_RequestTime 
	private dbg_FinishTime 
	private dbg_Data 
	private dbg_DB_Data 
	private dbg_AllVars 
	private dbg_Show_default 
	private Divsets(2) 

	public property let Enabled(bNewValue) ''[bool] sets "enabled" to true or false 
		dbg_Enabled = bNewValue 
	end property 
	public property get Enabled ''[bool] gets the "enabled" value 
		Enabled = dbg_Enabled 
	end property 
	public property let Show(bNewValue) ''[string] sets the debugging panel. Where each digit in the string represents a debug information pane in order (11 of them). 1=open, 0=closed 
		dbg_Show = bNewValue 
	end property 
	public property get Show ''[string] gets the debugging panel. 
		Show = dbg_Show 
	end property 
	public property let AllVars(bNewValue) ''[bool] sets wheather all variables will be displayed or not. true/false 
		dbg_AllVars = bNewValue 
	end property 
	public property get AllVars ''[bool] gets if all variables will be displayed. 
		AllVars = dbg_AllVars 
	end property 

	'Construktor => set the default values 
	private sub class_Initialize() 
		classname = "Kiss_Debug"
		dbg_RequestTime = Now() 
		dbg_AllVars = false 
		set dbg_Data = Server.CreateObject("Scripting.Dictionary") 
		Divsets(0) = "<tr><td style='cursor:hand;' onclick=""javascript:if (document.getElementById('data#sectname#').style.display=='none'){document.getElementById('data#sectname#').style.display='block';}else{document.getElementById('data#sectname#').style.display='none';}""><div id=sect#sectname# style=""font-weight:bold;cursor:hand;background:#7EA5D7;color:white;padding-left:4;padding-right:4;padding-bottom:2;"">|#title#| <div id=data#sectname# style=""cursor:text;display:none;background:#FFFFFF;padding-left:8;"" onclick=""window.event.cancelBubble = true;"">|#data#| </div>|</div>|" 
		Divsets(1) = "<tr><td><div id=sect#sectname# style=""font-weight:bold;cursor:hand;background:#7EA5D7;color:white;padding-left:4;padding-right:4;padding-bottom:2;"" onclick=""javascript:if (document.getElementById('data#sectname#').style.display=='none'){document.getElementById('data#sectname#').style.display='block';}else{document.getElementById('data#sectname#').style.display='none';}"">|#title#| <div id=data#sectname# style=""cursor:text;display:block;background:#FFFFFF;padding-left:8;"" onclick=""window.event.cancelBubble = true;"">|#data#| </div>|</div>|" 
		Divsets(2) = "<tr><td><div id=sect#sectname# style=""background:#7EA5D7;color:lightsteelblue;padding-left:4;padding-right:4;padding-bottom:2;"">|#title#| <div id=data#sectname# style=""display:none;background:lightsteelblue;padding-left:8"">|#data#| </div>|</div>|" 
		dbg_Show_default = "0,0,0,0,0,0,0,0,0,0,0" 
	end sub 
 
	'Destructor 
	private sub class_Terminate() 
		set dbg_Data = Nothing 
	end sub 

	'****************************************************************************************************************** 
	''@SDESCRIPTION: Adds a variable to the debug-informations. 
	''@PARAM: - label [string]: Description of the variable 
	''@PARAM: - output [variable]: The variable itself 
	'****************************************************************************************************************** 
	public sub Print(label, output) 
		if dbg_Enabled then 
			if err.number > 0 then 
				call dbg_Data.Add(ValidLabel(label), "!!! Error: " & err.number & " " & err.Description) 
				err.Clear 
			else 
				uniqueID = ValidLabel(label) 
				response.write uniqueID 
				call dbg_Data.Add(uniqueID, output) 
			end if 
		end if 
	end sub

	'****************************************************************************************************************** 
	'* ValidLabel 
	'****************************************************************************************************************** 
	private function ValidLabel(byval label) 
		dim i, lbl 
		i = 0 
		lbl = label 
		do 
			if not dbg_Data.Exists(lbl) then exit do 
			i = i + 1 
			lbl = label & "(" & i & ")" 
		loop until i = i 
		ValidLabel = lbl 
	end function

	'****************************************************************************************************************** 
	'* PrintCookiesInfo 
	'****************************************************************************************************************** 
	private sub PrintCookiesInfo(byval DivsetNo) 
		dim tbl, cookie, key, tmp 
		for Each cookie in Request.Cookies 
			if Not Request.Cookies(cookie).HasKeys then 
				tbl = AddRow(tbl, cookie, Request.Cookies(cookie)) 
			else 
				for Each key in Request.Cookies(cookie) 
					tbl = AddRow(tbl, cookie & "(" & key & ")", Request.Cookies(cookie)(key)) 
				next 
			end if 
		next 
		tbl = Maketable(tbl) 
		if Request.Cookies.count <= 0 then DivsetNo = 2 
		tmp = replace(replace(replace(Divsets(DivsetNo),"#sectname#","COOKIES"),"#title#","COOKIES"),"#data#",tbl) 
		Response.Write replace(tmp,"|", vbcrlf) 
	end sub

	'****************************************************************************************************************** 
	'* PrintSummaryInfo 
	'****************************************************************************************************************** 
	private sub PrintSummaryInfo(byval DivsetNo) 
		dim tmp, tbl 
		tbl = AddRow(tbl, "Time of Request",dbg_RequestTime) 
		tbl = AddRow(tbl, "Elapsed Time",DateDiff("s", dbg_RequestTime, dbg_FinishTime) & " seconds") 
		tbl = AddRow(tbl, "Request Type",Request.ServerVariables("REQUEST_METHOD")) 
		tbl = AddRow(tbl, "Status Code",Response.Status) 
		tbl = AddRow(tbl, "Script Engine",ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion) 
		tbl = Maketable(tbl) 
		tmp = replace(replace(replace(Divsets(DivsetNo),"#sectname#","SUMMARY"),"#title#","SUMMARY INFO"),"#data#",tbl) 
		Response.Write replace(tmp,"|", vbcrlf) 
	end sub

	'****************************************************************************************************************** 
	''@SDESCRIPTION: Adds the Database-connection object to the debug-instance. To display Database-information 
	''@PARAM: - oSQLDB [object]: connection-object 
	'****************************************************************************************************************** 
	public sub GrabDatabaseInfo(byval oSQLDB) 
		dbg_DB_Data = AddRow(dbg_DB_Data, "ADO Ver",oSQLDB.Version) 
		dbg_DB_Data = AddRow(dbg_DB_Data, "OLEDB Ver",oSQLDB.Properties("OLE DB Version")) 
		dbg_DB_Data = AddRow(dbg_DB_Data, "DBMS",oSQLDB.Properties("DBMS Name") & " Ver: " & oSQLDB.Properties("DBMS Version")) 
		dbg_DB_Data = AddRow(dbg_DB_Data, "Provider",oSQLDB.Properties("Provider Name") & " Ver: " & oSQLDB.Properties("Provider Version")) 
	end sub 

	'****************************************************************************************************************** 
	'* PrintdatabaseInfo 
	'****************************************************************************************************************** 
	private sub PrintdatabaseInfo(byval DivsetNo) 
		dim tbl 
		tbl = Maketable(dbg_DB_Data) 
		tbl = replace(replace(replace(Divsets(DivsetNo),"#sectname#","DATABASE"),"#title#","DATABASE INFO"),"#data#",tbl) 
		Response.Write replace(tbl,"|", vbcrlf) 
	end sub

	'****************************************************************************************************************** 
	'* PrintCollection 
	'****************************************************************************************************************** 
	private sub PrintCollection(Byval Name, ByVal Collection, ByVal DivsetNo, ByVal ExtraInfo) 
		dim vItem, tbl, Temp 
		for Each vItem In Collection 
			if isobject(Collection(vItem)) and Name <> "SERVER VARIABLES" and Name <> "QUERYStrING" and Name <> "FORM" then 
				tbl = AddRow(tbl, vItem, "{object}") 
			elseif isnull(Collection(vItem)) then 
				tbl = AddRow(tbl, vItem, "{null}") 
			elseif isarray(Collection(vItem)) then 
				tbl = AddRow(tbl, vItem, "{array}") 
			else 
				if dbg_AllVars then 
					tbl = AddRow(tbl, "<nobr>" & vItem & "</nobr>", server.HTMLEncode(Collection(vItem))) 
				elseif (Name = "SERVER VARIABLES" and vItem <> "ALL_HTTP" and vItem <> "ALL_RAW") or Name <> "SERVER VARIABLES" then 
					if Collection(vItem) <> "" then 
						tbl = AddRow(tbl, vItem, server.HTMLEncode(Collection(vItem))) ' & " {" & TypeName(Collection(vItem)) & "}") 
					else 
						tbl = AddRow(tbl, vItem, "...") 
					end if 
				end if 
			end if 
		next 
		if ExtraInfo <> "" then tbl = tbl & "<tr><td COLSPAN=2><HR></tr>" & ExtraInfo 
		tbl = Maketable(tbl) 
		if Collection.count <= 0 then DivsetNo =2 
		tbl = replace(replace(Divsets(DivsetNo),"#title#",Name),"#data#",tbl) 
		tbl = replace(tbl,"#sectname#",replace(Name," ","")) 
		Response.Write replace(tbl,"|", vbcrlf) 
	end sub

	'****************************************************************************************************************** 
	'* AddRow 
	'****************************************************************************************************************** 
	private function AddRow(byval t, byval var, byval val) 
		t = t & "|<tr valign=""top"">|<td>|" & var & "|<td>= " & val & "|</tr>" 
		AddRow = t 
	end function

	'****************************************************************************************************************** 
	'* Maketable 
	'****************************************************************************************************************** 
	private function Maketable(byval tdata) 
		tdata = "|<table border=""0"" style=""font-size:10pt;font-weight:normal;"">" + tdata + "</table>|" 
		Maketable = tdata 
	end function

	'****************************************************************************************************************** 
	''@SDESCRIPTION: Draws the Debug-panel 
	'****************************************************************************************************************** 
	public sub draw() 
		if dbg_Enabled then 
		dbg_FinishTime = Now() 
		dim Divset, x 
		Divset = split(dbg_Show_default,",") 
		dbg_Show = split(dbg_Show,",") 
		for x = 0 to ubound(dbg_Show) 
		divset(x) = dbg_Show(x) 
		next 
		Response.Write "<html><head><title>Debugging page</title><meta http-equiv=""Content-Type"" content=""text/html; charset=gbk"" /><body><table width=100% cellspacing=0 border=0 style=""font-family:arial;font-size:9pt;font-weight:normal;""><tr><td><div style=""background:#005A9E;color:white;padding:4;font-size:12pt;font-weight:bold;"">Debugging-console:</div>" 
		Call PrintSummaryInfo(divset(0)) 
		Call PrintCollection("VARIABLES", dbg_Data,divset(1),"") 
		Call PrintCollection("QUERYStrING", Request.QueryString(), divset(2),"") 
		Call PrintCollection("FORM", Request.form(),divset(3),"") 
		Call PrintCookiesInfo(divset(4)) 
		Call PrintCollection("SESSION", Session.Contents(),divset(5),AddRow(AddRow(AddRow("","Locale ID",Session.LCID & " (&H" & Hex(Session.LCID) & ")"),"Code Page",Session.CodePage),"Session ID",Session.SessionID)) 
		Call PrintCollection("APPLICATION", Application.Contents(),divset(6),"") 
		Call PrintCollection("SERVER VARIABLES", Request.ServerVariables(),divset(7),AddRow("","Timeout",Server.ScriptTimeout)) 
		Call PrintdatabaseInfo(divset(8)) 
		Call PrintCollection("SESSION STATIC OBJECTS", Session.StaticObjects(),divset(9),"") 
		Call PrintCollection("APPLICATION STATIC OBJECTS", Application.StaticObjects(),divset(10),"") 
		Response.Write "</table></body></html>" 
		end if 
	end sub 
end class 
%>  