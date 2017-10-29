Attribute VB_Name = "MainModule"
Option Explicit
Dim ac As Long
Dim enabled As Boolean
Public Form1S As Boolean

'防止生产DLL的软件BUG
Public Function FBUG() As Long
    '------------
End Function

' 接收应用AuthCode，酷Q读取应用信息后，如果接受该应用，将会调用这个函数并传递AuthCode。
' 不要在本函数处理其他任何代码，以免发生异常情况。如需执行初始化代码请在Startup事件中执行（Type=1001）。
Public Function Initialize(ByVal AuthCode As Long) As Long
    ac = AuthCode
    Initialize = 0
End Function

' 返回应用的ApiVer、Appid，打包后将不会调用
Public Function AppInfo() As Long
    Dim ZC() As Byte
    ZC = StrTByte(CQAPIVERTEXT & "," & CQAPPID)
    AppInfo = VarPtr(ZC(0))
End Function

' Type=1001 酷Q启动
' 无论本应用是否被启用，本函数都会在酷Q启动后执行一次，请在这里执行应用初始化代码。
' 如非必要，不建议在这里加载窗口。（可以添加菜单，让用户手动打开窗口）
Public Function eventStartup() As Long
    eventStartup = 0
End Function

' Type=1002 酷Q退出
' 无论本应用是否被启用，本函数都会在酷Q退出前执行一次，请在这里执行插件关闭代码。
' 本函数调用完毕后，酷Q将很快关闭，请不要再通过线程等方式执行其他代码。
Public Function eventExit() As Long
    eventExit = 0
End Function

' Type=1003 应用已被启用
' 当应用被启用后，将收到此事件。
' 如果酷Q载入时应用已被启用，则在_eventStartup(Type=1001,酷Q启动)被调用后，本函数也将被调用一次。
' 如非必要，不建议在这里加载窗口。（可以添加菜单，让用户手动打开窗口）
Public Function eventEnable() As Long
    enabled = True
    eventEnable = 0
End Function

' Type=1004 应用将被停用
' 当应用被停用前，将收到此事件。
' 如果酷Q载入时应用已被停用，则本函数*不会*被调用。
' 无论本应用是否被启用，酷Q关闭前本函数都*不会*被调用。
Public Function eventDisable() As Long
    enabled = False
    eventDisable = 0
End Function

' Type=21 私聊消息
' subType 子类型，11/来自好友 1/来自在线状态 2/来自群 3/来自讨论组
Public Function eventPrivateMsg(ByVal subType As Long, ByVal sendTime As Long, ByVal fromQQ As Currency, ByVal msg As Long, ByVal font As Long) As Long
    '如果要回复消息，请调用酷Q方法发送，并且这里 return EVENT_BLOCK - 截断本条消息，不再继续处理  注意：应用优先级设置为"最高"(10000)时，不得使用本返回值
    '如果不回复消息，交由之后的应用/过滤器处理，这里 return EVENT_IGNORE - 忽略本条消息
    If Form1S = True Then
        Form1.List1.AddItem ac & "/" & subType & "/" & sendTime & "/" & fromQQ * 10000 & "/" & msg & "/" & font
        Form1.Text1.SelStart = Len(Form1.Text1.Text)
        Form1.Text1.SelText = pGetStringFromPtr(msg)
    End If
    Dim ZC() As Byte
    ZC = StrTByte("test")
    CQ_sendPrivateMsg ac, fromQQ, VarPtr(ZC(0))
    eventPrivateMsg = EVENT_BLOCK
End Function

' Type=2 群消息
Public Function eventGroupMsg(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal fromQQ As Currency, ByVal fromAnonymous As Long, ByVal msg As Long, ByVal font As Long) As Long
    eventGroupMsg = EVENT_IGNORE '关于返回值说明, 见“eventPrivateMsg”函数
End Function

' Type=4 讨论组消息
Public Function eventDiscussMsg(ByVal subType As Long, ByVal sendTime As Long, ByVal fromDiscuss As Currency, ByVal fromQQ As Currency, ByVal msg As Long, ByVal font As Long) As Long
    eventDiscussMsg = EVENT_IGNORE '关于返回值说明, 见“eventPrivateMsg”函数
End Function

' Type=101 群事件-管理员变动
' subType 子类型，1/被取消管理员 2/被设置管理员
Public Function eventSystem_GroupAdmin(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal beingOperateQQ As Currency) As Long
    eventSystem_GroupAdmin = EVENT_IGNORE '关于返回值说明, 见“eventPrivateMsg”函数
End Function

' Type=102 群事件-群成员减少
' subType 子类型，1/群员离开 2/群员被踢 3/自己(即登录号)被踢
' fromQQ 操作者QQ(仅subType为2、3时存在)
' beingOperateQQ 被操作QQ
Public Function eventSystem_GroupMemberDecrease(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal fromQQ As Currency, ByVal beingOperateQQ As Currency) As Long
    eventSystem_GroupMemberDecrease = EVENT_IGNORE '关于返回值说明, 见“eventPrivateMsg”函数
End Function

' Type=103 群事件-群成员增加
' subType 子类型，1/管理员已同意 2/管理员邀请
' fromQQ 操作者QQ(即管理员QQ)
' beingOperateQQ 被操作QQ(即加群的QQ)
Public Function eventSystem_GroupMemberIncrease(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal fromQQ As Currency, ByVal beingOperateQQ As Currency) As Long
    eventSystem_GroupMemberIncrease = EVENT_IGNORE '关于返回值说明, 见“eventPrivateMsg”函数
End Function

'Type=201 好友事件-好友已添加
Public Function eventFriend_Add(ByVal subType As Long, ByVal sendTime As Long, ByVal fromQQ As Currency) As Long
    eventFriend_Add = EVENT_IGNORE '关于返回值说明, 见“eventPrivateMsg”函数
End Function

' Type=301 请求-好友添加
' msg 附言
' responseFlag 反馈标识(处理请求用)
Public Function eventRequest_AddFriend(ByVal subType As Long, ByVal sendTime As Long, ByVal fromQQ As Currency, ByVal msg As Long, ByVal responseflag As Long) As Long
    'CQ_setFriendAddRequest(ac, responseFlag, REQUEST_ALLOW, "");
    eventRequest_AddFriend = EVENT_IGNORE '关于返回值说明, 见“eventPrivateMsg”函数
End Function

' Type=302 请求-群添加
' subType 子类型，1/他人申请入群 2/自己(即登录号)受邀入群
' msg 附言
' responseFlag 反馈标识(处理请求用)
Public Function eventRequest_AddGroup(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal fromQQ As Currency, ByVal msg As Long, ByVal responseflag As Long) As Long
    'if (subType == 1) {
    '  CQ_setGroupAddRequestV2(ac, responseFlag, REQUEST_GROUPADD, REQUEST_ALLOW, "");
    '} else if (subType == 2) {
    '  CQ_setGroupAddRequestV2(ac, responseFlag, REQUEST_GROUPINVITE, REQUEST_ALLOW, "");
    '}
    eventRequest_AddGroup = EVENT_IGNORE '关于返回值说明, 见“eventPrivateMsg”函数
End Function

' 菜单，可在 .json 文件中设置菜单数目、函数名
' 如果不使用菜单，请在 .json 及此处删除无用菜单
Public Function menuA() As Long
    MsgBox "这是menuA，在这里载入窗口，或者进行其他工作。"
    menuA = 0
End Function

Public Function menuB() As Long
    Form1S = True
    Form1.Show
    menuB = 0
End Function


