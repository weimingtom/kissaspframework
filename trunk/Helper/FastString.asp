<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A fast connecting String class
'	File Name	: FastString.asp
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

class Kiss_FastString
	'************************************
	'变量定义
	'************************************
	'index --- 字符串数组的下标
	'ub ------ 用于调整数组度数的整数变量
	'ar() ---- 字符串数组
	private index, ub, ar()
	'************************************
	'实例 初始化/终止
	'************************************
	private sub class_Initialize()
		redim ar(50)
		index = 0
		ub = 49
	end sub
	private sub class_Terminate()
		Erase ar
	end sub
	'************************************
	'事件
	'************************************
	'默认事件，添加字符串
	public Default sub Add(value)
		ar(index) = value
		index = index+1
		if index>ub then
			ub = ub + 50
			redim Preserve ar(ub)
		end if
	end sub
	'************************************
	'方法
	'************************************
	'返回连接后的字符串
	public function Dump
		redim preserve ar(index-1)
		Dump = join(ar,"") '关键所在哦^_^
	end function
end class
%> 