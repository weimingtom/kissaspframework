<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A Caching class
'	File Name	: Cache.asp
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
dim ca
set ca = new Kiss_Cache

class Kiss_Cache

	public className	'Class name
	private IExpires
	private sysName, reloadTime
	private localName, localValue

	public Default property get Contents(ByVal Value)
		Contents = getCache(Value)
	end property

	public property let Expires(ByVal Value)
		IExpires = DateAdd("d", Value, Now)
	end property
	public property get Expires()
		Expires = IExpires
	end property

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Constructor
	'************************************************************* 
	private sub class_initialize()
		classname = "Kiss_Cache"
		sysName = Request.ServerVariables("PATH_TRANSLATED")
		sysName = Mid(sysName, InstrRev(sysName,"\") + 1, Len(sysName))
		sysName = Server.MapPath(sysName)
		sysName = Replace(sysName,"/","")
		sysName = Replace(sysName,"\","")
		sysName = Replace(sysName,":","")
		sysName = Replace(sysName,".","")
		sysName = Replace(sysName,"@","")
		reloadTime = 1	'cache 1 minutes
	end sub

	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
	end sub

	'************************************************************* 
	' Name: lock
	' Purpose: lock the applaction
	'************************************************************* 
	public sub lock()
		Application.lock()
	end sub

	'************************************************************* 
	' Name: unlock
	' Purpose: unLock the applaction
	'************************************************************* 
	public sub unlock()
		Application.unLock()
	end sub

	'************************************************************* 
	' Name: SetCache
	' Purpose: Set a cache
	'************************************************************* 
	public sub setCache(ByVal Key, ByVal Value, ByVal Expire)
		Expires = Expire
		lock
		Application(sysName & "_" & Key) = Value
		Application(sysName & "_" & Key & "_Expires") = Expires
		unLock
	end sub

	'************************************************************* 
	' Name: remove
	' Purpose: remove a cache
	'************************************************************* 
	public sub remove(ByVal Key)
		lock
		Application.Contents.remove(sysName & "_" & Key)
		Application.Contents.remove(sysName & "_" & Key & "_Expires")
		unLock
	end sub

	'************************************************************* 
	' Name: removeAll
	' Purpose: remove all cache
	'************************************************************* 
	public sub removeAll()
		clear()
	end sub

	'************************************************************* 
	' Name: clear
	' Purpose: Get a cookie
	'************************************************************* 
	private sub clear()
		Application.Contents.removeAll()
	end sub

	'************************************************************* 
	' Name: getCache
	' Purpose: Get a cache
	'************************************************************* 
	public function getCache(ByVal Key)
		dim Expire : Expire = Application(sysName & "_" & Key & "_Expires")
		if IsNull(Expire) Or IsEmpty(Expire) then
			getCache = ""
		else
			if IsDate(Expire) And CDate(Expire) > Now then
				getCache = Application(sysName & "_" & Key)
			else
				Call remove(sysName & "_" & Key)
				Value = ""
			end if
		end if
	end function

	'************************************************************* 
	' Name: compare
	' Purpose: Compare two caches
	'************************************************************* 
	public function compare(ByVal Key1, ByVal Key2)
		dim Cache1 : Cache1 = getCache(sysName & "_" & Key1)
		dim Cache2 : Cache2 = getCache(sysName & "_" & Key2)
		if TypeName(Cache1) <> TypeName(Cache2) then
			Compare = true
		else
			if TypeName(Cache1)="Object" then
				Compare = (Cache1 Is Cache2)
			else
				if TypeName(Cache1) = "Variant()" then
				Compare = (Join(Cache1, "^") = Join(Cache2, "^"))
				else
				Compare = (Cache1 = Cache2)
				end if
			end if
		end if
	end function

end class
%>