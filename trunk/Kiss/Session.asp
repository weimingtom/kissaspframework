<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A Sesssion class
'	File Name	: Session.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** code

'*************************************************************
'	Initialize the class
'*************************************************************
dim sn
set sn = new Kiss_Session

'Note: The main problem with sessions is WHEN they should end. We do not know if the user's last request was the final one or not. So we do not know how long we should keep the session "alive". Waiting too long for an idle session uses up resources on the server, but if the session is deleted too soon the user has to start all over again because the server has deleted all the information. Finding the right timeout interval can be difficult!
'Tip: If you are using session variables, store SMALL amounts of data in them. 

class Kiss_Session
	public className	'Class name
	private sysName

	public Default property get contents(ByVal Key)
		contents = getSession(sysName & "_" & Key)
	end property

	public property let TimeOut(ByVal Value)
		Session.TimeOut = Value
	end property

	public property get TimeOut()
		TimeOut = Session.TimeOut
	end property

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Constructor
	'************************************************************* 
	private sub class_initialize()
		classname = "Kiss_Session"
		dim scriptName : scriptName =  Request.ServerVariables("SCRIPT_NAME")
		sysName = Server.MapPath(scriptName)
		sysName = Replace(sysName,"/","")
		sysName = Replace(sysName,"\","")
		sysName = Replace(sysName,":","")
		sysName = Replace(sysName,".","")
		sysName = Replace(sysName,"@","")
	end sub

	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
	end sub

	'************************************************************* 
	' Name: getSession
	' Purpose: Get a cookie
	'************************************************************* 
	public function getSession(ByVal Key)
		getSession = Session(sysName & "_" & Key)
	end function

	'************************************************************* 
	' Name: setSession
	' Purpose: Set a session
	'************************************************************* 
	public sub setSession(ByVal Key, ByVal Value)
		Session(sysName & "_" & Key) = Value
	end sub

	'************************************************************* 
	' Name: remove
	' Purpose: Remove a session
	'************************************************************* 
	public sub remove(ByVal Key)
		Session.contents.Remove(sysName & "_" & Key)
	end sub

	'************************************************************* 
	' Name: removeAll
	' Purpose: Remove all sessions
	'************************************************************* 
	public sub removeAll()
		Session.Contents.RemoveAll()
	end sub

	'************************************************************* 
	' Name: clear
	' Purpose: To end a session immediately, you may use the Abandon method
	'************************************************************* 
	private sub clear()
		Session.Abandon()
	end sub

end class

%>