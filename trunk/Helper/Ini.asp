<%
if APP_ = "" or isNull(APP_) then 
	response.write ("No direct script access allowed")
	response.end
end if
'*************************************************************
'	class		: A ini class
'	File Name	: Ini.asp
'	Version		: 0.2.0
'	Author		: Andy Cai
'	License		: Dual licensed under the MIT (MIT-LICENSE.txt)
'				  and GPL (GPL-LICENSE.txt) licenses.
'	Date		: 2007-11-1
'*************************************************************


'*************************************************************
'	Sample of usage of the class
'*************************************************************

''============================= 属性说明 =========================
''=      ini.openFile = 文件路径（使用虚拟路径需在外部定义）  =
''=      ini.Codeset  = 编码设置，默认为 GB2312               =
''=      ini.Istrue   = 检测文件是否正常（存在）              =
''============================= 方法说明 =========================
''=      isGroup(组名)            检测组是否存在       =
''=      isNode(组名,节点名)            检测节点是否存在     =
''=      getGroup(组名)            读取组信息           =
''=      countGroup()            统计组数量           =
''=      readNode(组名,节点名)            读取节点数据         =
''=      writeGroup(组名)            创建组               =
''=      writeNode(组,节点,节点数据)      插入/更新节点数据    =
''=      deleteGroup(组名)            删除组               =
''=      deleteNode(组名,节点名)      删除节点             =
''=      save()                  保存文件             =
''=      close()                  清除内部数据（释放） =
''================================================================

dim ini
set ini = New Kiss_ini
ini.openFile = Server.MapPath("Config.ini")
'' Write data to .ini file
ini.writeNode("SiteConfig","SiteName","Leadbbs极速论坛")
ini.writeNode("SiteConfig","Mail","leadbbs@leadbbs.com")
ini.save()
'' Read data from .ini file
Response.Write("站点名称：" & ini.readNode("SiteConfig","SiteName"))

'*************************************************************
'	Initialize the class
'*************************************************************

class Kiss_ini
	''==================================================
    private Stream            ''// Stream 对象
    private FilePath      ''// 文件路径
   
    public Content            ''// 文件数据
    public Istrue            ''// 文件是否存在
    public IsAnsi            ''// 记录是否二进制
    public Codeset            ''// 数据编码
	''==================================================
   
    ''// 初始化
    private sub class_Initialize()
          set Stream      = Server.CreateObject("ADODB.Stream")
          Stream.Mode      = 3
          Stream.Type      = 2
          Codeset            = "gbk"
          IsAnsi            = true
          Istrue            = true
    end sub
   
    ''// 二进制流转换为字符串
    private function Bytes2bStr(bStr)
          if Lenb(bStr)=0 then
                Bytes2bStr = ""
                exit function
          end if
         
          dim BytesStream,StringReturn
          set BytesStream = Server.CreateObject("ADODB.Stream")
          With BytesStream
                .Type        = 2
                .Open
                .WriteText   bStr
                .Position    = 0
                .Charset     = Codeset
                .Position    = 2
                StringReturn = .ReadText
                .close
          end With
          Bytes2bStr       = StringReturn
          set BytesStream       = Nothing
          set StringReturn = Nothing
    end function
   
    ''// 设置文件路径
    property let openFile(iniFilePath)
          FilePath = iniFilePath
          Stream.Open
          On Error Resume next
          Stream.LoadFromFile(FilePath)
          ''// 文件不存在时返回给 Istrue
          if Err.Number<>0 then
                Istrue = false
                Err.Clear
          end if
          Content = Stream.ReadText(Stream.Size)
          if Not IsAnsi then Content=Bytes2bStr(Content)
    end property
   
    ''// 检测组是否存在[参数:组名]
    public function isGroup(GroupName)
          if Instr(Content,"["&GroupName&"]")>0 then
                isGroup = true
          else
                isGroup = false
          end if
    end function
   
    ''// 读取组信息[参数:组名]
    public function getGroup(GroupName)
          dim TempGroup
          if Not isGroup(GroupName) then exit function
          ''// 开始寻找头部截取
          TempGroup = Mid(Content,Instr(Content,"["&GroupName&"]"),Len(Content))
          ''// 剔除尾部
          if Instr(TempGroup,VbCrlf&"[")>0 then TempGroup=Left(TempGroup,Instr(TempGroup,VbCrlf&"[")-1)
          if Right(TempGroup,1)<>Chr(10) then TempGroup=TempGroup&VbCrlf
          getGroup = TempGroup
    end function
   
    ''// 检测节点是否存在[参数:组名,节点名]
    public function isNode(GroupName,NodeName)
          if Instr(getGroup(GroupName),NodeName&"=") then
                isNode = true
          else
                isNode = false
          end if
    end function
   
    ''// 创建组[参数:组名]
    public sub writeGroup(GroupName)
          if Not isGroup(GroupName) And GroupName<>"" then
                Content = Content & "[" & GroupName & "]" & VbCrlf
          end if
    end sub
   
    ''// 读取节点数据[参数:组名,节点名]
    public function readNode(GroupName,NodeName)
          if Not isNode(GroupName,NodeName) then exit function
          dim TempContent
          ''// 取组信息
          TempContent = getGroup(GroupName)
          ''// 取当前节点数据
          TempContent = Right(TempContent,Len(TempContent)-Instr(TempContent,NodeName&"=")+1)
          TempContent = Replace(Left(TempContent,Instr(TempContent,VbCrlf)-1),NodeName&"=","")
          readNode = ReplaceData(TempContent,0)
    end function
   
    ''// 写入节点数据[参数:组名,节点名,节点数据]
    public sub writeNode(GroupName,NodeName,NodeData)
          ''// 组不存在时写入组
          if Not isGroup(GroupName) then writeGroup(GroupName)
         
          ''// 寻找位置插入数据
          ''/// 获取组
          dim TempGroup : TempGroup = getGroup(GroupName)
         
          ''/// 在组尾部追加
          dim NewGroup
          if isNode(GroupName,NodeName) then
                NewGroup = Replace(TempGroup,NodeName&"="&ReplaceData(readNode(GroupName,NodeName),1),NodeName&"="&ReplaceData(NodeData,1))
          else
                NewGroup = TempGroup & NodeName & "=" & ReplaceData(NodeData,1) & VbCrlf
          end if
         
          Content = Replace(Content,TempGroup,NewGroup)
    end sub
   
    ''// 删除组[参数:组名]
    public sub deleteGroup(GroupName)
          Content = Replace(Content,getGroup(GroupName),"")
    end sub
   
   
    ''// 删除节点[参数:组名,节点名]
    public sub deleteNode(GroupName,NodeName)
          dim TempGroup
          dim NewGroup
          TempGroup = getGroup(GroupName)
          NewGroup = Replace(TempGroup,NodeName&"="&readNode(GroupName,NodeName)&VbCrlf,"")
          if Right(NewGroup,1)<>Chr(10) then NewGroup = NewGroup&VbCrlf
          Content = Replace(Content,TempGroup,NewGroup)
    end sub
   
    ''// 替换字符[实参:替换目标,数据流向方向]
    ''      字符转换[防止关键符号出错]
    ''      [            --->      {(@)}
    ''      ]            --->      {(#)}
    ''      =            --->      {($)}
    ''      回车      --->      {(1310)}
    public function ReplaceData(Data_Str,IsIn)
          if IsIn then
                ReplaceData = Replace(Replace(Replace(Data_Str,"[","{(@)}"),"]","{(#)}"),"=","{($)}")
                ReplaceData = Replace(ReplaceData,Chr(13)&Chr(10),"{(1310)}")
          else
                ReplaceData = Replace(Replace(Replace(Data_Str,"{(@)}","["),"{(#)}","]"),"{($)}","=")
                ReplaceData = Replace(ReplaceData,"{(1310)}",Chr(13)&Chr(10))
          end if
    end function
   
    ''// 保存文件数据
    public sub save()
          With Stream
                .close
                .Open
                .WriteText Content
                .saveToFile FilePath,2
          end With
    end sub
   
    ''// 关闭、释放
    public sub close()
          set Stream = Nothing
          set Content = Nothing
    end sub
   
end class

%>