<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: File Manipulating class
'	File Name	: File.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

class Kiss_File      
	public className	'Class name
    private adSaveCreateOverWrite   
    private adSaveCreateNotExist    
	private objStream

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Construtor
	' Remarks: set the initial value to the private variable
	'************************************************************* 
    private sub class_Initialize() 
		classname = "Kiss_File"
		set objStream =server.createobject("ADODB.Stream")
		adSaveCreateOverWrite =2 
        adSaveCreateNotExist = 1
    end sub 

 	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
		 set objStream = nothing
    end sub

	'************************************************************* 
	' Name: readFile
	' Param: filepath as file directory
	' Purpose: Read contents of the file
	'************************************************************* 
	function readFile(filepath) 
		on error resume next 
		With objStream
			.Type = 2
			.Mode = 3
			.Open
			.LoadFromFile Server.MapPath(filepath)
			.Charset = "gbk"
			.Position = 2
			readfile = .ReadText()
			.Close
		end with
		if err.number<>0 then
			throwException("Reading the file  <span style=""color:#f60;font-weight:bold;"">" & filepath & "</span> fails. Please check if the file exists.")
		end if
	end function  

	'************************************************************* 
	' Name: writeFile
	' Param: filepath as target directory
	'        str as content writing to the file
	' Purpose: Write the contents to the file
	'************************************************************* 
    function writeFile(filepath,str) 
        on error resume next 
		with objStream
			.Charset = "gbk" 
			.Open 
			.WriteText str 
			.SaveToFile Server.MapPath(filepath), adSaveCreateOverWrite 
		end with
		if err.number<>0 then
			throwException("Write the file  <span style=""color:#f60;font-weight:bold;"">" & filepath & "</span> fails. Please check if the file exists.")
		end if
    end function 

	'************************************************************* 
	' Name: copyFile
	' Param: filepath_s as source directory
	'        filepath_d as target directory
	' Purpose: Copy the file to another path
	'************************************************************* 
	function copyFile(filepath_s,filepath_d)     
		on error resume next     
		with objStream
			.Charset = "gbk" 
			.Open 
			.LoadFromFile Server.MapPath(filepath_s) 
			.SaveToFile Server.MapPath(filepath_d), adSaveCreateOverWrite 
		end with
		if err.number<>0 then
			throwException("Copy the file <span style=""color:#f60;font-weight:bold;"">" & filepath_s & "</span> to <span style=""color:#f00;font-weight:bold;"">" & filepath_d & "</span> fails. Please check if the file exists.")
		end if
	end function

end class 

%> 