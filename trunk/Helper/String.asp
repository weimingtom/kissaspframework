<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A String operating class
'	File Name	: String.asp
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

class Kiss_String

	''****************************************************************************
	'''' @功能说明: 把字符串换为char型数组
	'''' @参数说明:  - str [string]: 需要转换的字符串
	'''' @返回值:   - [Array] Char型数组
	''****************************************************************************
	public function toCharArray(byVal str)
	 redim charArray(len(str))
	 for i = 1 to len(str)
	  charArray(i-1) = Mid(str,i,1)
	 next
	 toCharArray = charArray
	end function

	''****************************************************************************
	'''' @功能说明: 把一个数组转换成一个字符串
	'''' @参数说明:  - arr [Array]: 需要转换的数据
	'''' @返回值:   - [string] 字符串
	''****************************************************************************
	public function arrayToString(byVal arr)
	 for i = 0 to UBound(arr)
	  strObj = strObj & arr(i)
	 next
	 arrayToString = strObj
	end function

	''****************************************************************************
	'''' @功能说明: 检查源字符串str是否以chars开头
	'''' @参数说明:  - str [string]: 源字符串
	'''' @参数说明:  - chars [string]: 比较的字符/字符串
	'''' @返回值:   - [bool]
	''****************************************************************************
	public function startsWith(byVal str, chars)
	 if Left(str,len(chars)) = chars then
	  startsWith = true
	 else
	  startsWith = false
	 end if
	end function

	''****************************************************************************
	'''' @功能说明: 检查源字符串str是否以chars结尾
	'''' @参数说明:  - str [string]: 源字符串
	'''' @参数说明:  - chars [string]: 比较的字符/字符串
	'''' @返回值:   - [bool]
	''****************************************************************************
	public function endsWith(byVal str, chars)
	 if Right(str,len(chars)) = chars then
	  endsWith = true
	 else
	  endsWith = false
	 end if
	end function

	''****************************************************************************
	'''' @功能说明: 复制N个字符串str
	'''' @参数说明:  - str [string]: 源字符串
	'''' @参数说明:  - n [int]: 复制次数
	'''' @返回值:   - [string] 复制后的字符串
	''****************************************************************************
	public function clone(byVal str, n)
	 for i = 1 to n
	  value = value & str
	 next
	 clone = value
	end function

	''****************************************************************************
	'''' @功能说明: 删除源字符串str的前N个字符
	'''' @参数说明:  - str [string]: 源字符串
	'''' @参数说明:  - n [int]: 删除的字符个数
	'''' @返回值:   - [string] 删除后的字符串
	''****************************************************************************
	public function trimStart(byVal str, n)
	 value = Mid(str, n+1)
	 trimStart = value
	end function

	''****************************************************************************
	'''' @功能说明: 删除源字符串str的最后N个字符串
	'''' @参数说明:  - str [string]: 源字符串
	'''' @参数说明:  - n [int]: 删除的字符个数
	'''' @返回值:   - [string] 删除后的字符串
	''****************************************************************************
	public function trimend(byVal str, n)
	 value = Left(str, len(str)-n)
	 trimend = value
	end function

	''****************************************************************************
	'''' @功能说明: 检查字符character是否是英文字符 A-Z or a-z
	'''' @参数说明:  - character [char]: 检查的字符
	'''' @返回值:   - [bool] 如果是英文字符,返回TRUE,反之为FALSE
	''****************************************************************************
	public function isAlphabetic(byVal character)
	 asciiValue = cint(asc(character))
	 if (65 <= asciiValue and asciiValue <= 90) or (97 <= asciiValue and asciiValue <= 122) then
	  isAlphabetic = true
	 else
	  isAlphabetic = false
	 end if
	end function

	''****************************************************************************
	'''' @功能说明: 对str字符串进行大小写转换
	'''' @参数说明:  - str [string]: 源字符串
	'''' @返回值:   - [string] 转换后的字符串
	''****************************************************************************
	public function swapCase(str)
	 for i = 1 to len(str)
	  current = mid(str, i, 1)
	  if isAlphabetic(current) then
	   high = asc(ucase(current))
	   low = asc(lcase(current))
	   sum = high + low
	   return = return & chr(sum-asc(current))
	  else
	   return = return & current
	  end if
	 next
	 swapCase = return
	end function

	''****************************************************************************
	'''' @功能说明: 将源字符串str中每个单词的第一个字母转换成大写
	'''' @参数说明:  - str [string]: 源字符串
	'''' @返回值:   - [string] 转换后的字符串
	''****************************************************************************
	public function capitalize(str)
	 words = split(str," ")
	 for i = 0 to ubound(words)
	  if not i = 0 then
	   tmp = " "
	  end if
	  tmp = tmp & ucase(left(words(i), 1)) & right(words(i), len(words(i))-1)
	  words(i) = tmp
	 next
	 capitalize = arrayToString(words)
	end function

	''****************************************************************************
	'''' @功能说明: 将源字符Str后中的''过滤为''''
	'''' @参数说明:  - str [string]: 源字符串
	'''' @返回值:   - [string] 转换后的字符串
	''****************************************************************************
	public function checkstr(Str)
	 if Trim(Str)="" Or IsNull(str) then
	  checkstr=""
	 else
	  checkstr=Replace(Trim(Str),"''","''''")
	 end if
	end function

	''****************************************************************************
	'''' @功能说明: 将字符串中的str中的HTML代码进行过滤
	'''' @参数说明:  - str [string]: 源字符串
	'''' @返回值:   - [string] 转换后的字符串
	''****************************************************************************
	public function HtmlEncode(str)
	 if Trim(Str)="" Or IsNull(str) then
	  HtmlEncode=""
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
	  HtmlEncode=str
	 end if
	end function

	''****************************************************************************
	'''' @功能说明: 计算源字符串Str的长度(一个中文字符为2个字节长)
	'''' @参数说明:  - str [string]: 源字符串
	'''' @返回值:   - [Int] 源字符串的长度
	''****************************************************************************
	public function strLen(Str)
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

	''****************************************************************************
	'''' @功能说明: 截取源字符串Str的前LenNum个字符(一个中文字符为2个字节长)
	'''' @参数说明:  - str [string]: 源字符串
	'''' @参数说明:  - LenNum [int]: 截取的长度
	'''' @返回值:   - [string]: 转换后的字符串
	''****************************************************************************
	public function CutStr(Str,LenNum)
	 dim P_num
	 dim I,X
	 if StrLen(Str)<=LenNum then
	  Cutstr=Str
	 else
	  P_num=0
	  X=0
	  Do While Not P_num > LenNum-2
	   X=X+1
	   if Asc(Mid(Str,X,1))<0 then
		P_num=Int(P_num) + 2
	   else
		P_num=Int(P_num) + 1
	   end if
	   Cutstr=Left(Trim(Str),X)&"..."
	  Loop
	 end if
	end function

end class
%>