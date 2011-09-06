<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: Array extensive class
'	File Name	: Array.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

'*** '使用:
'*** '比如传入一个数组元素值和顺序是 17,23,55,99,37,比如存入TheArr变量中
'*** dim TheArr
'*** TheArr = array(17,55,23)
'*** set ca = New Kiss_Array
'*** ca.srcStr=TheArr
'*** '取最小值
'*** ca.cMin
'*** '最大值
'*** ca.cMax
'*** '取元素个数
'*** ca.Length
'*** '输出传入的数组
'*** ca.print_r
'*** '在未尾增加元素
'*** ca.print_rr(ca.AddItem("99"))
'*** '从大到小排序
'*** ca.print_rr(ca.Sort)
'*** '*** '从小到大
'*** ca.print_rr(ca.SortAsc)
'*** '移除值为xx的元素
'*** ca.print_rr(ca.RemoveItem(17))
'*** '移除索引值为index的元素
'*** ca.print_rr(ca.RemoveItemI(1))
'*** '在索引值为index的元素前插入一个值为str的元素
'*** ca.print_rr(ca.AddItemI(1, 88))

'*************************************************************
'	Initialize the class
'*************************************************************

class Kiss_Array
     private tMin		''//最小值 thisMin
     private tMax		''//最大值
     private tStr		''//数组串
     private tmpStr		''//暂存数组串
     private tSortStr	''//排序串，当对同一数组串进行操作时避免多次排序操作
     private tLength    ''//元素个数
    
	'************************************************************* 
	' Name: class_Initialize
	' Purpose: Construtor
	'************************************************************* 
    private sub class_Initialize()
           tMin			= 0
           tMax			= 0
           tStr			= "0"
           tLength      = 0
           tSortStr		= ""
    end sub
    
 	'************************************************************* 
	' Name: class_Terminate
	' Purpose: Deconstrutor
	'************************************************************* 
    private sub class_Terminate()
    end sub

 	'************************************************************* 
	' Name: cMin
	' Purpose: get the Minimal Value
	'************************************************************* 
     property get cMin
           if tMin=0 then
                 dim tempStr
                 if tSortStr="" then
                       'sStr=Split(Sort,",")
                       tempStr=Sort
                 else
                       tempStr=Split(tSortStr,",")
                 end if
                 tMin=tempStr(0)
           end if
           cMin=tMin
     end property

 	'************************************************************* 
	' Name: cMax
	' Purpose: get the Maximal Value
	'************************************************************* 
     property get cMax
           if tMax=0 then
                 dim tempStr
                 if tSortStr="" then
                       'sStr=Split(Sort,",")
                       tempStr=Sort
                 else
                       tempStr=Split(tSortStr,",")
                 end if
                 tMax=tempStr(Ubound(tempStr))
           end if
           cMax=tMax
     end property

 	'************************************************************* 
	' Name: Length
	' Purpose: get an Array Length
	'************************************************************* 
     property get Length
           if tStr="" Or tStr=0 then
                 Length      =0
           else
                 tLength      =Ubound(Split(tStr,","))+1
                 Length      =tLength
           end if
     end property

  	'************************************************************* 
	' Name: srcStr
	' Purpose: set value of an Array to value of the property
	'************************************************************* 
    property let srcStr(arr)
           tmpStr = arr
		   dim str : str = Join(arr, ",")
		   if str="" then
                 tStr="0"
           else
                 tStr=str
           end if
     end property

  	'************************************************************* 
	' Name: getArray
	' Purpose: get the Array
	'************************************************************* 
     public function getArray
           getArray=split(tStr, ",")
     end function

  	'************************************************************* 
	' Name: getArray
	' Purpose: get the Array
	'************************************************************* 
     public function getSortArray
           getSortArray=split(tSortStr, ",")
     end function

  	'************************************************************* 
	' Name: print_r
	' Purpose: Print all the elements of an Array 
	'************************************************************* 
     public function print_r
           response.write "<pre>" & tStr & "</pre>"
     end function

  	'************************************************************* 
	' Name: print_rr
	' Purpose: Print all the elements of the Array arr
	'************************************************************* 
     public function print_rr(arr)
		dim str : str = Join(arr, ",")
		response.write "<pre>" & str & "</pre>"
     end function

  	'************************************************************* 
	' Name: Sort
	' Purpose: Reverse order to an Array
	'************************************************************* 
     public function Sort
           dim Arr,i
           redim Arr(Ubound(Split(tStr,",")))
           dim tmp,j
           for i=0 to Ubound(Arr)
                 Arr(i)=CDBL(Split(tStr,",")(i))
           next
           for I= 0 to Ubound(Arr)
                 for  j=i+1 to Ubound(Arr)
                       if Arr(i)<Arr(j) then
                             tmp      =Arr(i)
                             Arr(i)      =Arr(j)
                             Arr(j)      =tmp
                       end if
                 next
           next
           tSortStr=Join(Arr,",")
           'sStr = tSortStr
           Sort = getSortArray
     end function
   
   	'************************************************************* 
	' Name: SortAsc
	' Purpose: Ascending order to an Array
	'************************************************************* 
    public function SortAsc
           dim tmp,i,Arr
           if tSortStr="" then
                 tmp=Sort
           else
                 tmp=tSortStr
           end if
           redim Arr(Ubound(Split(tmp,",")))
           Arr=Split(tmp,",")
           for i= Ubound(Arr) to 0 Step -1
                 SortAsc=SortAsc & Arr(i) & ","
           next
           tSortStr = Left(SortAsc,Len(SortAsc)-1)
           'sStr		= tSortStr
           SortAsc	= getSortArray
     end function

   	'************************************************************* 
	' Name: AddItem
	' Purpose: Append a element to an Array
	'************************************************************* 
     public function AddItem(str)
           if tStr="" OR tStr="0" then
                 tStr=str
           else
                 tStr=tStr & "," & str
           end if
           tMin      =0
           tMax      =0
           tSortStr=""
           tLength      =	tLength + 1
           'AddItem      =	tStr
           AddItem = getArray
     end function

   	'************************************************************* 
	' Name: AddItemBefore
	' Purpose: Add a element to first position of an Array
	'************************************************************* 
     public function AddItemBefore(str)
           if tStr="" OR tStr="0" then
                 tStr=str
           else
                 tStr=str & "," & tStr
           end if
           tMin      =0
           tMax      =0
           tSortStr=""
           tLength       = tLength + 1
           'AddItemBefore = tStr
           AddItemBefore = getArray
     end function

   	'************************************************************* 
	' Name: RemoveItem
	' Purpose: Remove a element equaling to str of an Array
	'************************************************************* 
     public function RemoveItem(str)
           tStr      =Replace("," & tStr & "," , "," & str & "," , ",")
           tStr      =Mid(tStr,2,Len(tStr)-2)
           tMin      =0
           tMax      =0
           tSortStr=Replace("," & tSortStr & "," , "," & str & "," , ",")
           tSortStr=Mid(tSortStr,2,Len(tSortStr)-2)
           tLength      =0
           'RemoveItem = tStr
           RemoveItem = getArray
     end function

   	'************************************************************* 
	' Name: AddItemI
	' Purpose: Add a element to index of an Array
	'************************************************************* 
     public function AddItemI(index,str)
           if index>=Length then
                 AddItem(str)
                 exit function
           end if
           if index<=0 then
                 AddItemBefore(str)
                 exit function
           end if
         
           dim Arr,i,tmps
           redim Arr(Ubound(Split(tStr,",")))
           Arr=Split(tStr,",")
           for i=0 to index-1
                 tmps=tmps & Arr(i) & ","
           next
           tmps=tmps & str & ","
           for i=index to Ubound(Arr)
                 tmps=tmps & Arr(i) & ","
           next
           tmps      =Left(tmps,Len(tmps)-1)
           tStr      =tmps
           tSortStr=""
           tLength      =tLength+1
           tMin      =0
           tMax      =0
           'AddItemI=tStr
           AddItemI = getArray
     end function

   	'************************************************************* 
	' Name: RemoveItemI
	' Purpose: Remove a element from index of an Array
	'************************************************************* 
     public function RemoveItemI(index)
           if index>=Length Or index<=0 then
                 exit function
           end if
         
           dim Arr,i,tmps
           redim Arr(Ubound(Split(tStr,",")))
           Arr=Split(tStr,",")
           for i=0 to Ubound(Arr)
                 if i<>index then tmps=tmps & Arr(i) & ","
           next
           tmps			= tmps & str & ","
           tmps			= Left(tmps,Len(tmps)-1)
           tStr			= tmps
           tSortStr		= ""
           tLength      = 0
           tMin			= 0
           tMax			= 0
           'RemoveItemI=tStr
           RemoveItemI = getArray
     end function

end class

%>