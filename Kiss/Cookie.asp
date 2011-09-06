<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A Cookie class
'	File Name	: Cookie.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** 'create the cookie
'*** Response.Cookies("brownies") = 13
'*** 'get the cookie
'*** myBrownie = Request.Cookies("brownies")
'*** Response.Cookies("name").Expires = #January 1,2009#
'*** 'create a big cookie -- Array
'*** Response.Cookies("brownies")("numberEaten") = 13 
'*** Response.Cookies("brownies")("eater") = "George" 
'*** Response.Cookies("brownies")("weight") = 400
'*** For Each key In Request.Cookies("Brownies")
'*** 	Response.Write("<br />" & key & " = " & _
'*** 	 Request.Cookies("Brownies")(key))	
'*** Next

'*************************************************************
'	Initialize the class
'*************************************************************
dim ck
set ck = new Kiss_Cookie

class Kiss_Cookie

	public className	'Class name
	private CurrentKey

	public Default property get contents(ByVal Value)
		contents = getCookie(Value)
	end property

	public property let Expires(ByVal Value)
		Response.Cookies(CurrentKey).Expires = DateAdd("d", Value, Now)
	end property
	public property get Expires()
		Expires = Request.Cookies(CurrentKey).Expires
	end property

	public property let Path(ByVal Value)
		Response.Cookies(CurrentKey).Path = Value
	end property
	public property get Path()
		Path = Request.Cookies(CurrentKey).Path
	end property

	public property let Domain(ByVal Value)
		Response.Cookies(CurrentKey).Domain = Value
	end property
	public property get Domain()
		Domain = Request.Cookies(CurrentKey).Domain
	end property

	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Save the session
	'************************************************************* 
	private sub class_initialize()
		classname = "Kiss_Cookie"
	end sub

	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
	private sub class_Terminate()
	end sub

	'************************************************************* 
	' Name: setCookie
	' Purpose: Add a cookie
	'************************************************************* 
	public sub setCookie(ByVal Key, ByVal Value, ByVal Options)
		Response.Cookies(Key) = Value
		CurrentKey = Key
		if Not (IsNull(Options) Or IsEmpty(Options) Or Options = "") then
			if IsArray(Options) then
				dim l : l = UBound(Options)
				Expire = Options(0)
				if l = 1 then Path = Options(1)
				if l = 2 then Domain = Options(2)
			else
				Expire = Options
			end if
		end if
	end sub

	'************************************************************* 
	' Name: remove
	' Purpose: Remove a cookie
	'************************************************************* 
	public sub remove(ByVal Key)
		CurrentKey = Key
		Expires = -1000
	end sub

	'************************************************************* 
	' Name: removeAll
	' Purpose: Remove all cookies
	'************************************************************* 
	public sub removeAll()
		Clear()
	end sub

	'************************************************************* 
	' Name: Clear
	' Purpose: Remove all cookies
	'************************************************************* 
	private sub Clear()
		dim iCookie
		for Each iCookie In Request.Cookies
			Response.Cookies(iCookie).Expires = formatDateTime(Now)
		next
	end sub

	'************************************************************* 
	' Name: removeAll
	' Purpose: remove all cookies
	'************************************************************* 
	public function getCookie(ByVal Key)
		getCookie = Request.Cookies(Key)
	end function

end class
%>