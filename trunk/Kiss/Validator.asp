<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A Validating class
'	File Name	: Validator.asp
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
dim va
set va = new Kiss_Validator

class Kiss_Validator
	public className	'Class name
	private Re
	private ICodeName
	private ICodeSessionName

	public property let CodeName(ByVal PCodeName)
		ICodeName = PCodeName
	end property

	public property get CodeName()
		CodeName = ICodeName
	end property

	public property let CodeSessionName(ByVal PCodeSessionName)
		ICodeSessionName = PCodeSessionName
	end property

	public property get CodeSessionName()
		CodeSessionName = ICodeSessionName
	end property

	private sub class_Initialize()
		classname = "Kiss_Validator"
		set Re = New RegExp
		Re.IgnoreCase = true
		Re.Global = true
		Me.CodeName = "vCode"
		Me.CodeSessionName = "vCode"
	end sub

	private sub class_Terminate()
		set Re = Nothing
	end sub

	public function IsEmail(ByVal Str)
		IsEmail = Test("^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$", Str)
	end function

	public function IsUrl(ByVal Str)
		IsUrl = Test("^http:\/\/[A-Za-z0-9]+\.[A-Za-z0-9]+[\/=\?%\-&_~`@[\]\'':+!]*([^<>""])*$", Str)
	end function

	public function IsNum(ByVal Str)
		IsNum= Test("^\d+$", Str)
	end function

	public function IsQQ(ByVal Str)
		IsQQ = Test("^[1-9]\d{4,8}$", Str)
	end function

	public function IsZip(ByVal Str)
		IsZip = Test("^[1-9]\d{5}$", Str)
	end function

	public function IsIdCard(ByVal Str)
		IsIdCard = Test("^\d{15}(\d{2}[A-Za-z0-9])?$", Str)
	end function

	public function IsChinese(ByVal Str)
		IsChinese = Test("^[\u0391-\uFFE5]+$", Str)
	end function

	public function IsEnglish(ByVal Str)
		IsEnglish = Test("^[A-Za-z]+$", Str)
	end function

	public function IsMobile(ByVal Str)
		IsMobile = Test("^((\(\d{3}\))|(\d{3}\-))?13\d{9}$", Str)
	end function

	public function IsPhone(ByVal Str)
		IsPhone = Test("^((\(\d{3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}$", Str)
	end function

	public function IsSafe(ByVal Str)
		IsSafe = (Test("^(([A-Z]*|[a-z]*|\d*|[-_\~!@#\$%\^&\*\.\(\)\[\]\{\}<>\?\\\/\''\""]*)|.{0,5})$|\s", Str) = false)
	end function

	public function IsNotEmpty(ByVal Str)
		IsNotEmpty = LenB(Str) > 0
	end function

	public function IsDateformat(ByVal Str, ByVal format)
		IF Not IsDate(Str) then
			IsDateformat = false
			exit function
		end IF

		IF format = "YMD" then
			IsDateformat = Test("^((\d{4})|(\d{2}))([-./])(\d{1,2})\4(\d{1,2})$", Str)
		else
			IsDateformat = Test("^(\d{1,2})([-./])(\d{1,2})\\2((\d{4})|(\d{2}))$", Str)
		end IF
	end function

	public function IsEqual(ByVal Src, ByVal Tar)
		IsEqual = (Src = Tar)
	end function

	public function Compare(ByVal Op1, ByVal Operator, ByVal Op2)
		Compare = false
		IF Dic.Exists(Operator) then
			Compare = Eval(Dic.Item(Operator))
		elseif IsNotEmpty(Op1) then
			Compare = Eval(Op1 & Operator & Op2 )
		end IF
	end function

	public function Range(ByVal Src, ByVal Min, ByVal Max)
		Min = CInt(Min) : Max = CInt(Max)
		Range = (Min < Src And Src < Max)
	end function

	public function Group(ByVal Src, ByVal Min, ByVal Max)
		Min = CInt(Min) : Max = CInt(Max)
		dim Num : Num = UBound(Split(Src, ",")) + 1
		Group = Range(Num, Min - 1, Max + 1)
	end function

	public function Custom(ByVal Str, ByVal Reg)
		Custom = Test(Reg, Str)
	end function

	public function Limit(ByVal Str, ByVal Min, ByVal Max)
		Min = CInt(Min) : Max = CInt(Max)
		dim L : L = Len(Str)
		Limit = (Min <= L And L <= Max)
	end function

	public function LimitB(ByVal Str, ByVal Min, ByVal Max)
		Min = CInt(Min) : Max = CInt(Max)
		dim L : L =bLen(Str)
		LimitB = (Min <= L And L <= Max)
	end function

	private function Test(ByVal Pattern, ByVal Str)
		if IsNull(Str) Or IsEmpty(Str) then
			Test = false
		else
			Re.Pattern = Pattern
			Test = Re.Test(CStr(Str))
		end if
	end function

	public function bLen(ByVal Str)
		bLen = Len(Replace(Str, "[^\x00-\xFF]", ".."))
	end function

	private function Replace(ByVal Str, ByVal Pattern, ByVal ReStr)
		Re.Pattern = Pattern
		Replace = Re.Replace(Str, ReStr)
	end function

	private function B2S(ByVal iStr)
		dim reVal : reVal= ""
		dim i, Code, nCode
		for i = 1 to LenB(iStr)
			Code = AscB(MidB(iStr, i, 1))
			IF Code < &h80 then
				reVal = reVal & Chr(Code)
			else
				nCode = AscB(MidB(iStr, i+1, 1))
				reVal = reVal & Chr(CLng(Code) * &h100 + CInt(nCode))
				i = i + 1
			end IF
		next
		B2S = reVal
	end function

	public function SafeStr(ByVal Name)
		if IsNull(Name) Or IsEmpty(Name) then
			SafeStr = false
		else
			SafeStr = Replace(Trim(Name), "(\s*and\s*\w*=\w*)|[''%&<>=]", "")
		end if
	end function

	public function SafeNo(ByVal Name)
		if IsNull(Name) Or IsEmpty(Name) then
			SafeNo = 0
		else
			SafeNo = (Replace(Trim(Name), "^[\D]*(\d+)[\D\d]*$", "$1"))
		end if
	end function

	public function IsValidCode()
		IsValidCode = ((Request.form(Me.CodeName) = Session(Me.CodeSessionName)) AND Session(Me.CodeSessionName) <> "")
	end function

	'************************************************************* 
	' Name: IsValidPost
	' Purpose: Check if a URL is outside
	'************************************************************* 
	public function IsValidPost()
		dim Url1 : Url1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
		dim Url2 : Url2 = Cstr(Request.ServerVariables("SERVER_NAME"))
		IsValidPost = (Mid(Url1, 8, Len(Url2)) = Url2)
	end function

end class
%>