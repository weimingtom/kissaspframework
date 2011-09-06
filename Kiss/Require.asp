<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: Dynamic Include Asp File
'	File Name	: Require.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** require("kiss/Db.asp")

'*************************************************************
'	Initialize the class
'*************************************************************
dim require
set require = new Kiss_Require
dim include
set include = require

class Kiss_Require
    
	public className	'Class name
	private loadFile_vars

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Construtor
	'************************************************************* 
	private sub class_Initialize()
      classname = "Kiss_Require"
	  set loadFile_vars = server.createobject("scripting.dictionary")
    end sub
    
 	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
    'private sub class_deactivate()
    private sub class_Terminate()
      'arr_variables.removeall
      set loadFile_vars = nothing
      'set loadFile = nothing
    end sub
	
	'************************************************************* 
	' Name: Require
	' Param: str_path as file path
	' Purpose: Require the file
	'************************************************************* 
	public default function require(byval str_path)
      dim str_source
      if str_path <> "" then
		
		' Checks if the file exists
		on error resume next 
		if not fileExists(str_path) then
		  show404("The file <span style=""color:#f60; font-weight:bold;"">" & str_path & "</span> doesn't exist.")
		  'call showMsg(NO_PAGE, "Warning")
		end if
		str_source = readfile(str_path)
        if str_source <> "" then
          processloadFiles str_source
          convert2code str_source
          formatcode str_source
          if str_source <> "" then
			  executeglobal str_source
			  if err.number<>0 then
				throwException("Loading <span style=""color:#f60; font-weight:bold;"">" & str_path & "</span> on error. Please check the file <span style=""color:#f60; font-weight:bold;"">" & str_path & "</span>")
				err.clear
			  end if
              loadFile_vars.removeall
          end if
        end if
      end if
    end function
	
	'************************************************************* 
	' Name: processloadFiles
	' Param: str_source as source code of a file
	' Purpose: Convert asp keyword to another code
	'************************************************************* 
	private sub convert2code(str_source)
      dim i, str_temp, arr_temp, int_len
      if str_source <> "" then
        if instr(str_source,"%" & ">") > instr(str_source,"<" & "%") then
          str_temp = replace(str_source,"<" & "%","|#@#%")
          str_temp = replace(str_temp,"%" & ">","|#@#")
          if left(str_temp,1) = "|#@#" then str_temp = right(str_temp,len(str_temp) - 1)
          if right(str_temp,1) = "|#@#" then str_temp = left(str_temp,len(str_temp) - 1)
          arr_temp = split(str_temp,"|#@#")
          int_len = ubound(arr_temp)
          if (int_len + 1) > 0 then
            for i = 0 to int_len
              str_temp = trim(arr_temp(i))
              str_temp = replace(str_temp,vbcrlf & vbcrlf,vbcrlf)
              if left(str_temp,2) = vbcrlf then str_temp = right(str_temp,len(str_temp) - 2)
              if right(str_temp,2) = vbcrlf then str_temp = left(str_temp,len(str_temp) - 2)
              if left(str_temp,1) = "%" then
                str_temp = right(str_temp,len(str_temp) - 1)
                if left(str_temp,1) = "=" then
                  str_temp = right(str_temp,len(str_temp) - 1)
                  str_temp = "response.write " & str_temp
                end if
              else
                if str_temp <> "" then
                  loadFile_vars.add i, str_temp
                  str_temp = "response.write loadFile_vars.item(" & i & ")" 
                end if
              end if
              str_temp = replace(str_temp,chr(34) & chr(34) & " & ","")
              str_temp = replace(str_temp," & " & chr(34) & chr(34),"")
              if right(str_temp,2) <> vbcrlf then str_temp = str_temp
              arr_temp(i) = str_temp
            next
            str_source = join(arr_temp,vbcrlf)
          end if
        else
          if str_source <> "" then
            loadFile_vars.add "var", str_source
            str_source = "response.write loadFile_vars.item(""var"")"
          end if
        end if
      end if
    end sub

	'************************************************************* 
	' Name: processloadFiles
	' Param: str_source as source code of a file
	' Purpose: read the code of the include file such as <!--#include file="file.asp"-->
	'************************************************************* 
   private sub processloadFiles(str_source)
      dim int_start, str_path, str_mid, str_temp
      str_source = replace(str_source,"<!-- #","<!--#")
      int_start = instr(str_source,"<!--#include")
      str_mid = lcase(getbetween(str_source,"<!--#include","-->"))
      do until int_start = 0
        str_mid = lcase(getbetween(str_source,"<!--","-->"))
        int_start = instr(str_mid,"#include")
        if int_start >  0 then
          str_temp = lcase(getbetween(str_mid,chr(34),chr(34)))
          str_temp = trim(str_temp)
          str_path = readfile(str_temp)
          str_source = replace(str_source,"<!--" & str_mid & "-->",str_path & vbcrlf)
        end if
        int_start = instr(str_source,"#include")
      loop
    end sub

	'************************************************************* 
	' Name: formatcode
	' Param: str_code as source code of a file
	' Purpose: Format the code
	'************************************************************* 
    private sub formatcode(str_code)
      dim i, arr_temp, int_len
      str_code = replace(str_code,vbcrlf & vbcrlf,vbcrlf)
      if left(str_code,2) = vbcrlf then str_code = right(str_code,len(str_code) - 2)
      str_code = trim(str_code)
      if instr(str_code,vbcrlf) > 0 then
        arr_temp = split(str_code,vbcrlf)
        for i = 0 to ubound(arr_temp)
          arr_temp(i) = ltrim(arr_temp(i))
          if arr_temp(i) <> "" then arr_temp(i) = arr_temp(i) & vbcrlf
        next
        str_code = join(arr_temp,"")
        arr_temp = vbnull
      end if
    end sub

	'************************************************************* 
	' Name: readfile
	' Param: File as file path
	' Purpose: Read contents of the file
	'************************************************************* 
	private function readfile(ByVal File)
		dim objStream
		set objStream = Server.CreateObject("ADODB.Stream")
		With objStream
			.Type = 2
			.Mode = 3
			.Open
			.LoadFromFile Server.MapPath(File)
			.Charset = "gbk"  'set the encoding of the page charset
			.Position = 2
			readfile = .ReadText()
			.Close
		end With
		set objStream = Nothing
	end function

    'private function readfile(str_path)
      'dim objfso, objfile
      'if str_path <> "" then
        'if instr(str_path,":") = 0 then str_path = server.mappath(str_path)
        'set objfso = server.createobject("scripting.filesystemobject")
        'if objfso.fileexists(str_path) then
          'set objfile = objfso.opentextfile(str_path, 1, false)
          'if err.number = 0 then
            'readfile = objfile.readall
            'objfile.close
          'end if
          'set objfile = nothing
        'end if
        'set objfso = nothing
      'end if
	  'readfile = LoadFile(str_path)
    'end function

    private function getbetween(strdata, strstart, strend)
      dim lngstart, lngend
      lngstart = instr(strdata, strstart) + len(strstart)
      if (lngstart <> 0) then
        lngend = instr(lngstart, strdata, strend)
        if (lngend <> 0) then
          getbetween = mid(strdata, lngstart, lngend - lngstart)
        end if
      end if
    end function
	
end class

%>