Attribute VB_Name = "MainModule"
Option Explicit
Dim ac As Long
Dim enabled As Boolean
Public Form1S As Boolean

'��ֹ����DLL�����BUG
Public Function FBUG() As Long
    '------------
End Function

' ����Ӧ��AuthCode����Q��ȡӦ����Ϣ��������ܸ�Ӧ�ã���������������������AuthCode��
' ��Ҫ�ڱ��������������κδ��룬���ⷢ���쳣���������ִ�г�ʼ����������Startup�¼���ִ�У�Type=1001����
Public Function Initialize(ByVal AuthCode As Long) As Long
    ac = AuthCode
    Initialize = 0
End Function

' ����Ӧ�õ�ApiVer��Appid������󽫲������
Public Function AppInfo() As Long
    Dim ZC() As Byte
    ZC = StrTByte(CQAPIVERTEXT & "," & CQAPPID)
    AppInfo = VarPtr(ZC(0))
End Function

' Type=1001 ��Q����
' ���۱�Ӧ���Ƿ����ã������������ڿ�Q������ִ��һ�Σ���������ִ��Ӧ�ó�ʼ�����롣
' ��Ǳ�Ҫ����������������ش��ڡ���������Ӳ˵������û��ֶ��򿪴��ڣ�
Public Function eventStartup() As Long
    eventStartup = 0
End Function

' Type=1002 ��Q�˳�
' ���۱�Ӧ���Ƿ����ã������������ڿ�Q�˳�ǰִ��һ�Σ���������ִ�в���رմ��롣
' ������������Ϻ󣬿�Q���ܿ�رգ��벻Ҫ��ͨ���̵߳ȷ�ʽִ���������롣
Public Function eventExit() As Long
    eventExit = 0
End Function

' Type=1003 Ӧ���ѱ�����
' ��Ӧ�ñ����ú󣬽��յ����¼���
' �����Q����ʱӦ���ѱ����ã�����_eventStartup(Type=1001,��Q����)�����ú󣬱�����Ҳ��������һ�Ρ�
' ��Ǳ�Ҫ����������������ش��ڡ���������Ӳ˵������û��ֶ��򿪴��ڣ�
Public Function eventEnable() As Long
    enabled = True
    eventEnable = 0
End Function

' Type=1004 Ӧ�ý���ͣ��
' ��Ӧ�ñ�ͣ��ǰ�����յ����¼���
' �����Q����ʱӦ���ѱ�ͣ�ã��򱾺���*����*�����á�
' ���۱�Ӧ���Ƿ����ã���Q�ر�ǰ��������*����*�����á�
Public Function eventDisable() As Long
    enabled = False
    eventDisable = 0
End Function

' Type=21 ˽����Ϣ
' subType �����ͣ�11/���Ժ��� 1/��������״̬ 2/����Ⱥ 3/����������
Public Function eventPrivateMsg(ByVal subType As Long, ByVal sendTime As Long, ByVal fromQQ As Currency, ByVal msg As Long, ByVal font As Long) As Long
    '���Ҫ�ظ���Ϣ������ÿ�Q�������ͣ��������� return EVENT_BLOCK - �ضϱ�����Ϣ�����ټ�������  ע�⣺Ӧ�����ȼ�����Ϊ"���"(10000)ʱ������ʹ�ñ�����ֵ
    '������ظ���Ϣ������֮���Ӧ��/�������������� return EVENT_IGNORE - ���Ա�����Ϣ
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

' Type=2 Ⱥ��Ϣ
Public Function eventGroupMsg(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal fromQQ As Currency, ByVal fromAnonymous As Long, ByVal msg As Long, ByVal font As Long) As Long
    eventGroupMsg = EVENT_IGNORE '���ڷ���ֵ˵��, ����eventPrivateMsg������
End Function

' Type=4 ��������Ϣ
Public Function eventDiscussMsg(ByVal subType As Long, ByVal sendTime As Long, ByVal fromDiscuss As Currency, ByVal fromQQ As Currency, ByVal msg As Long, ByVal font As Long) As Long
    eventDiscussMsg = EVENT_IGNORE '���ڷ���ֵ˵��, ����eventPrivateMsg������
End Function

' Type=101 Ⱥ�¼�-����Ա�䶯
' subType �����ͣ�1/��ȡ������Ա 2/�����ù���Ա
Public Function eventSystem_GroupAdmin(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal beingOperateQQ As Currency) As Long
    eventSystem_GroupAdmin = EVENT_IGNORE '���ڷ���ֵ˵��, ����eventPrivateMsg������
End Function

' Type=102 Ⱥ�¼�-Ⱥ��Ա����
' subType �����ͣ�1/ȺԱ�뿪 2/ȺԱ���� 3/�Լ�(����¼��)����
' fromQQ ������QQ(��subTypeΪ2��3ʱ����)
' beingOperateQQ ������QQ
Public Function eventSystem_GroupMemberDecrease(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal fromQQ As Currency, ByVal beingOperateQQ As Currency) As Long
    eventSystem_GroupMemberDecrease = EVENT_IGNORE '���ڷ���ֵ˵��, ����eventPrivateMsg������
End Function

' Type=103 Ⱥ�¼�-Ⱥ��Ա����
' subType �����ͣ�1/����Ա��ͬ�� 2/����Ա����
' fromQQ ������QQ(������ԱQQ)
' beingOperateQQ ������QQ(����Ⱥ��QQ)
Public Function eventSystem_GroupMemberIncrease(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal fromQQ As Currency, ByVal beingOperateQQ As Currency) As Long
    eventSystem_GroupMemberIncrease = EVENT_IGNORE '���ڷ���ֵ˵��, ����eventPrivateMsg������
End Function

'Type=201 �����¼�-���������
Public Function eventFriend_Add(ByVal subType As Long, ByVal sendTime As Long, ByVal fromQQ As Currency) As Long
    eventFriend_Add = EVENT_IGNORE '���ڷ���ֵ˵��, ����eventPrivateMsg������
End Function

' Type=301 ����-�������
' msg ����
' responseFlag ������ʶ(����������)
Public Function eventRequest_AddFriend(ByVal subType As Long, ByVal sendTime As Long, ByVal fromQQ As Currency, ByVal msg As Long, ByVal responseflag As Long) As Long
    'CQ_setFriendAddRequest(ac, responseFlag, REQUEST_ALLOW, "");
    eventRequest_AddFriend = EVENT_IGNORE '���ڷ���ֵ˵��, ����eventPrivateMsg������
End Function

' Type=302 ����-Ⱥ���
' subType �����ͣ�1/����������Ⱥ 2/�Լ�(����¼��)������Ⱥ
' msg ����
' responseFlag ������ʶ(����������)
Public Function eventRequest_AddGroup(ByVal subType As Long, ByVal sendTime As Long, ByVal fromGroup As Currency, ByVal fromQQ As Currency, ByVal msg As Long, ByVal responseflag As Long) As Long
    'if (subType == 1) {
    '  CQ_setGroupAddRequestV2(ac, responseFlag, REQUEST_GROUPADD, REQUEST_ALLOW, "");
    '} else if (subType == 2) {
    '  CQ_setGroupAddRequestV2(ac, responseFlag, REQUEST_GROUPINVITE, REQUEST_ALLOW, "");
    '}
    eventRequest_AddGroup = EVENT_IGNORE '���ڷ���ֵ˵��, ����eventPrivateMsg������
End Function

' �˵������� .json �ļ������ò˵���Ŀ��������
' �����ʹ�ò˵������� .json ���˴�ɾ�����ò˵�
Public Function menuA() As Long
    MsgBox "����menuA�����������봰�ڣ����߽�������������"
    menuA = 0
End Function

Public Function menuB() As Long
    Form1S = True
    Form1.Show
    menuB = 0
End Function


