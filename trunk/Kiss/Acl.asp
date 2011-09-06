<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A access control list class
'	File Name	: Acl.asp
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
'	Initialize the class
'*************************************************************
dim acl
set acl = new Kiss_Acl

class Kiss_Acl
	public classname ' Class name

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Contructor
	'************************************************************* 
	private sub class_initialize()
		classname = "Kiss_Acl"
	end sub

	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
	end sub

end class
%>