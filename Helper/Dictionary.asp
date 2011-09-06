<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A dictionary class
'	File Name	: Dictionary.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

''/*用法:
''/*dim objDic,sKey,I,sValue
''/*set objDic=New Dictionaryclass
''/*Add方法:Add(字典的Key值,字典数据)    说明:如果"字典的Key值"已存在则Add方法失败
''/*objDic.Add "a","字母a"       ''Add方法
''/*objDic.Add "b","字母b"    
''/*objDic.Add "c","字母c"
''/*''insert方法:insert(被插入位置的Key值,新的字典Key值,新的字典数据,插入方式:b后面,f前面)
''/*objDic.insert "a","aa","字母aa","b"    
''/*objDic.insert "b","bb","字母bb","f"
''/*''Exists方法,返回是否存在以"b"为Key值的字典数据
''/*Response.Write objDic.Exists("b")
''/*sKey=objDic.Keys             ''获取Keys集合,(数组集合)
''/*sValue=objDic.Items          ''获取字典数据集合(数组集合)
''/*objDic.Item("a")="aaaaaa"    ''Item属性方法:返回或设置对应Key的字典数据
''/*for I=0 To objDic.Count-1    ''Count属性返回有多少条字典数据
''/*    ''Item属性方法:返回或设置对应Key的字典数据
''/*    Response.Write objDic.Item(sKey(I))&"<br>"
''/*next
''/*Remove方法:Remove(字典的Key值)
''/*objDic.Remove("a")           ''删除Key值为a的字典数据
''/*objDic.RemoveAll             ''清空字典数据
''/*objDic.ErrCode               ''返回操作字典时的一些错误代码(调试时用)
''/*objDic.ClearErr              ''清空错误代码(调试时用)
''/*set objDic=nothing

'*************************************************************
'	Initialize the class
'*************************************************************
dim va
set va = new Kiss_Dictionary

class Kiss_Dictionary
	dim ArryObj()     
	dim MaxIndex        
	dim CurIndex        
	dim C_ErrCode        

	private sub class_Initialize
		CurIndex=0          
		C_ErrCode=0        
		MaxIndex=50        
		redim ArryObj(1,MaxIndex)  
	end sub

	private sub class_Terminate  
		Erase ArryObj   
	end sub

	public property get ErrCode  
		ErrCode=C_ErrCode
	end property

	public property get Count  
		Count=CurIndex
	end property

	public property get Keys     
		dim KeyCount,ArryKey(),I
		KeyCount=CurIndex-1
		redim ArryKey(KeyCount)
		for I=0 To KeyCount
		  ArryKey(I)=ArryObj(0,I)
		next
		Keys=ArryKey
		Erase ArryKey
	end property

	public property get Items    
		dim KeyCount,ArryItem(),I
		KeyCount=CurIndex-1
		redim ArryItem(KeyCount)
		for I=0 To KeyCount
		 if isObject(ArryObj(1,I)) then
			set ArryItem(I)=ArryObj(1,I)
		else
		 ArryItem(I)=ArryObj(1,I)
		end if
		next
		Items=ArryItem
		Erase ArryItem
	end property


	public property let Item(sKey,sVal)  
		if sIsEmpty(sKey) then
		exit property
		end if
		dim i,iType
		iType=getType(sKey)
		if iType=1 then  
		if sKey>CurIndex Or sKey<1 then
		C_ErrCode=2
		 exit property
		end if
		end if
		if iType=0 then
		for i=0 to CurIndex-1
		  if ArryObj(0,i)=sKey then
		  if isObject(sVal) then
			 set ArryObj(1,i)=sVal
		else
		  ArryObj(1,i)=sVal
		end if
		exit property
		end if
		next
		elseif iType=1 then
			 sKey=sKey-1
		  if isObject(sVal) then
			 set ArryObj(1,sKey)=sVal
		else
		  ArryObj(1,sKey)=sVal
		end if
		exit property
		end if
		C_ErrCode=2 
	end property

	public property get Item(sKey)
		if sIsEmpty(sKey) then
		 Item=Null
		exit property
		end if
		dim i,iType
		iType=getType(sKey)
		if iType=1 then 
		if sKey>CurIndex Or sKey<1 then
		 Item=Null
		exit property
		end if
		end if
		if iType=0 then
		for i=0 to CurIndex-1
		  if ArryObj(0,i)=sKey then
		  if isObject(ArryObj(1,i)) then
			 set Item=ArryObj(1,i)
		else
		  Item=ArryObj(1,i)
		end if
		exit property
		end if
		next
		elseif iType=1 then
			 sKey=sKey-1
		  if isObject(ArryObj(1,sKey)) then
			 set Item=ArryObj(1,sKey)
		else
		  Item=ArryObj(1,sKey)
		end if
		exit property
		end if
		Item=Null
	end property

	'************************************************************* 
	' Name: Add
	' param: sKey as a dictionary key
	' param: sVal as a dictionary Value
	' Purpose: Add a element into a dictionary
	'************************************************************* 
	public sub Add(sKey,sVal)  
		''On Error Resume next
		if Exists(sKey) Or C_ErrCode=9 then
			C_ErrCode=1               
			exit sub
		end if
			if CurIndex>MaxIndex then
			MaxIndex=MaxIndex+1        
			redim Preserve ArryObj(1,MaxIndex)
		end if
		ArryObj(0,CurIndex)=Cstr(sKey)      
		if isObject(sVal) then
		   set ArryObj(1,CurIndex)=sVal       
		else
		   ArryObj(1,CurIndex)=sVal   
		end if
		CurIndex=CurIndex+1
	end sub

	'************************************************************* 
	' Name: insert
	' param: sKey as a dictionary key
	' Purpose: insert a element into a dictionary
	'************************************************************* 
	public sub insert(sKey,nKey,nVal,sMethod)
		if Not Exists(sKey) then
		   C_ErrCode=4
		   exit sub
		end if
		   if Exists(nKey) Or C_ErrCode=9 then
		   C_ErrCode=4 
		   exit sub
		end if
		sType=getType(sKey) 
		dim ArryResult(),I,sType,subIndex,sAdd
		redim ArryResult(1,CurIndex) 
		if sIsEmpty(sMethod) then sMethod="b" 
		sMethod=lcase(cstr(sMethod))
		subIndex=CurIndex-1
		sAdd=0
		if sType=0 then     
		  if sMethod="1" Or sMethod="b" Or sMethod="back" then 
			   for I=0 TO subIndex
				  ArryResult(0,sAdd)=ArryObj(0,I)
			if IsObject(ArryObj(1,I)) then
			  set ArryResult(1,sAdd)=ArryObj(1,I)
			else
			  ArryResult(1,sAdd)=ArryObj(1,I)
			end if
			if ArryObj(0,I)=sKey then  ' insert data
			  sAdd=sAdd+1
			  ArryResult(0,sAdd)=nKey
		   if IsObject(nVal) then
			 set ArryResult(1,sAdd)=nVal
		   else
			 ArryResult(1,sAdd)=nVal
		   end if
			end if
			sAdd=sAdd+1
			next
		  else
			   for I=0 TO subIndex
			if ArryObj(0,I)=sKey then  ' insert data
			  ArryResult(0,sAdd)=nKey
		   if IsObject(nVal) then
			 set ArryResult(1,sAdd)=nVal
		   else
			 ArryResult(1,sAdd)=nVal
		   end if
		   sAdd=sAdd+1
			end if
			ArryResult(0,sAdd)=ArryObj(0,I)
			if IsObject(ArryObj(1,I)) then
			  set ArryResult(1,sAdd)=ArryObj(1,I)
			else
			  ArryResult(1,sAdd)=ArryObj(1,I)
			end if
			sAdd=sAdd+1
			next
		  end if
		elseif sType=1 then
		  sKey=sKey-1                    
		  if sMethod="1" Or sMethod="b" Or sMethod="back" then  ' insert after sKey 
			   for I=0 TO sKey									' Get data before sKey
				  ArryResult(0,I)=ArryObj(0,I)
			if IsObject(ArryObj(1,I)) then
			  set ArryResult(1,I)=ArryObj(1,I)
			else
			  ArryResult(1,I)=ArryObj(1,I)
			end if
			next
		 ' insert new data
		 ArryResult(0,sKey+1)=nKey
		 if IsObject(nVal) then
			set ArryResult(1,sKey+1)=nVal
		 else
			ArryResult(1,sKey+1)=nVal
		 end if
		 ' Get data after sKey
			   for I=sKey+1 TO subIndex
				  ArryResult(0,I+1)=ArryObj(0,I)
			if IsObject(ArryObj(1,I)) then
			  set ArryResult(1,I+1)=ArryObj(1,I)
			else
			  ArryResult(1,I+1)=ArryObj(1,I)
			end if
			next
		  else
			   for I=0 TO sKey-1              ' Get data before sKey-1
				  ArryResult(0,I)=ArryObj(0,I)
			if IsObject(ArryObj(1,I)) then
			  set ArryResult(1,I)=ArryObj(1,I)
			else
			  ArryResult(1,I)=ArryObj(1,I)
			end if
			next
		 ' insert new data
		 ArryResult(0,sKey)=nKey
		 if IsObject(nVal) then
			set ArryResult(1,sKey)=nVal
		 else
			ArryResult(1,sKey)=nVal
		 end if
		 'Get data after sKey
			   for I=sKey TO subIndex
				  ArryResult(0,I+1)=ArryObj(0,I)
			if IsObject(ArryObj(1,I)) then
			  set ArryResult(1,I+1)=ArryObj(1,I)
			else
			  ArryResult(1,I+1)=ArryObj(1,I)
			end if
			next
		  end if
		else
		  C_ErrCode=3
		  exit sub
		end if
		redim ArryObj(1,CurIndex)  ' Reset data
		for I=0 To CurIndex
		 ArryObj(0,I)=ArryResult(0,I)
		 if isObject(ArryResult(1,I)) then
			set ArryObj(1,I)=ArryResult(1,I)
		 else
			ArryObj(1,I)=ArryResult(1,I)
		 end if
		next
		MaxIndex=CurIndex
		Erase ArryResult
		CurIndex=CurIndex+1      ' After inserting pointer plus 1
	end sub

	'************************************************************* 
	' Name: Exists
	' param: sKey as a dictionary key
	' Purpose: Check if a dictionary exists
	'************************************************************* 
	public function Exists(sKey)
		if sIsEmpty(sKey) then
		  Exists=false
		  exit function
		end if
		dim I,vType
		vType=getType(sKey)
		if vType=0 then
		  for I=0 To CurIndex-1
			if ArryObj(0,I)=sKey then
			Exists=true
			exit function
		 end if
		  next
		elseif vType=1 then
			 if sKey<=CurIndex And sKey>0 then
			 Exists=true
			 exit function
		  end if
		end if
		Exists=false
	end function

	'************************************************************* 
	' Name: Remove
	' param: sKey as a dictionary key
	' Purpose: Remove a dictionary element
	'************************************************************* 
	public sub Remove(sKey) 
		if Not Exists(sKey) then
		   C_ErrCode=3
		   exit sub
		end if
		sType=getType(sKey)
		dim ArryResult(),I,sType,sAdd
		redim ArryResult(1,CurIndex-2)
		sAdd=0
		if sType=0 then
			 for I=0 TO CurIndex-1
			if ArryObj(0,I)<>sKey then
				  ArryResult(0,sAdd)=ArryObj(0,I)
			if IsObject(ArryObj(1,I)) then
			  set ArryResult(1,sAdd)=ArryObj(1,I)
			else
			  ArryResult(1,sAdd)=ArryObj(1,I)
			end if
			sAdd=sAdd+1
		 end if
		  next
		elseif sType=1 then
		  sKey=sKey-1 
			 for I=0 TO CurIndex-1
			if I<>sKey then
				  ArryResult(0,sAdd)=ArryObj(0,I)
			if IsObject(ArryObj(1,I)) then
			  set ArryResult(1,sAdd)=ArryObj(1,I)
			else
			  ArryResult(1,sAdd)=ArryObj(1,I)
			end if
			sAdd=sAdd+1
		 end if
		  next
		else
		  C_ErrCode=3
		  exit sub
		end if
		MaxIndex=CurIndex-2
		redim ArryObj(1,MaxIndex)  'reset data
		for I=0 To MaxIndex
		 ArryObj(0,I)=ArryResult(0,I)
		 if isObject(ArryResult(1,I)) then
			set ArryObj(1,I)=ArryResult(1,I)
		 else
			ArryObj(1,I)=ArryResult(1,I)
		 end if
		next
		Erase ArryResult
		CurIndex=CurIndex-1
	end sub

	'************************************************************* 
	' Name: RemoveAll
	' Purpose: Clear all dictionary data
	'************************************************************* 
	public sub RemoveAll
		redim ArryObj(MaxIndex) 'redim the array
		CurIndex=0
	end sub

	'************************************************************* 
	' Name: ClearErr
	' Purpose: Reset error
	'************************************************************* 
	public sub ClearErr
		C_ErrCode=0
	end sub

	'************************************************************* 
	' Name: sIsEmpty
	' Param: sVal as variable name
	' Purpose: Check if a variable is empty
	'************************************************************* 
	private function sIsEmpty(sVal)  
		if IsEmpty(sVal) then
		   C_ErrCode=9  ' Error code
		   sIsEmpty=true
		   exit function
		end if
		if IsNull(sVal) then
		   C_ErrCode=9 
		   sIsEmpty=true
		   exit function
		end if
		if Trim(sVal)="" then
		   C_ErrCode=9 s
		   sIsEmpty=true
		   exit function
		end if
		sIsEmpty=false
	end function

	'************************************************************* 
	' Name: getType
	' Param: sVal as variable name
	' Purpose: Get a variable type
	'************************************************************* 
	private function getType(sVal)
		dim sType
		sType=TypeName(sVal)
		Select Case sType
			Case "String"
			  getType=0
			Case "Integer","Long","Single","Double"
			  getType=1
			Case else
			  getType=-1
		end Select
	end function

end class

%>