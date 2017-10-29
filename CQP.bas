Attribute VB_Name = "CQP"
Option Explicit
Public Const CQAPIVERTEXT = "9"
Public Const CQAPPID = "com.example.demovb"

Public Const EVENT_IGNORE = 0        '�¼�_����
Public Const EVENT_BLOCK = 1         '�¼�_����

Public Const REQUEST_ALLOW = 1       '����_ͨ��
Public Const REQUEST_DENY = 2        '����_�ܾ�

Public Const REQUEST_GROUPADD = 1    '����_Ⱥ���
Public Const REQUEST_GROUPINVITE = 2 '����_Ⱥ����

Public Const CQLOG_DEBUG = 0         '���� ��ɫ
Public Const CQLOG_INFO = 10         '��Ϣ ��ɫ
Public Const CQLOG_INFOSUCCESS = 11  '��Ϣ(�ɹ�) ��ɫ
Public Const CQLOG_INFORECV = 12     '��Ϣ(����) ��ɫ
Public Const CQLOG_INFOSEND = 13     '��Ϣ(����) ��ɫ
Public Const CQLOG_WARNING = 20      '���� ��ɫ
Public Const CQLOG_ERROR = 30        '���� ��ɫ
Public Const CQLOG_FATAL = 40        '�������� ���

' ����˽����Ϣ
' QQID Ŀ��QQ��
' msg ��Ϣ����
Public Declare Function CQ_sendPrivateMsg Lib "CQP" (ByVal AuthCode As Long, ByVal QQID As Currency, ByVal msg As Long) As Long

' ����Ⱥ��Ϣ
' groupid Ⱥ��
' msg ��Ϣ����
Public Declare Function CQ_sendGroupMsg Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal msg As Long) As Long

' ������������Ϣ
' discussid �������
' msg ��Ϣ����
Public Declare Function CQ_sendDiscussMsg Lib "CQP" (ByVal AuthCode As Long, ByVal discussid As Currency, ByVal msg As Long) As Long

' ������ �����ֻ���
' QQID QQ��
Public Declare Function CQ_sendLike Lib "CQP" (ByVal AuthCode As Long, ByVal QQID As Currency) As Long

' ��ȺԱ�Ƴ�
' groupid Ŀ��Ⱥ
' QQID QQ��
' rejectaddrequest ���ٽ��մ��˼�Ⱥ���룬������
Public Declare Function CQ_setGroupKick Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal QQID As Currency, ByVal rejectaddrequest As Long) As Long

' ��ȺԱ����
' groupid Ŀ��Ⱥ
' QQID QQ��
' duration ���Ե�ʱ�䣬��λΪ�롣���Ҫ�����������д0��
Public Declare Function CQ_setGroupBan Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal QQID As Currency, ByVal duration As Currency) As Long

' ��Ⱥ����Ա
' groupid Ŀ��Ⱥ
' QQID QQ��
' setadmin true:���ù���Ա false:ȡ������Ա
Public Declare Function CQ_setGroupAdmin Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal QQID As Currency, ByVal setadmin As Long) As Long

' ��ȫȺ����
' groupid Ŀ��Ⱥ
' enableban true:���� false:�ر�
Public Declare Function CQ_setGroupWholeBan Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal enableban As Long) As Long

' ������ȺԱ����
' groupid Ŀ��Ⱥ
' anomymous Ⱥ��Ϣ�¼��յ��� anomymous ����
' duration ���Ե�ʱ�䣬��λΪ�롣��֧�ֽ����
Public Declare Function CQ_setGroupAnonymousBan Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal anomymous As Long, ByVal duration As Currency) As Long

' ��Ⱥ��������
' groupid Ŀ��Ⱥ
' enableanomymous true:���� false:�ر�
Public Declare Function CQ_setGroupAnonymous Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal enableanomymous As Long) As Long

' ��Ⱥ��Ա��Ƭ
' groupid Ŀ��Ⱥ
' QQID Ŀ��QQ
' newcard ����Ƭ(�ǳ�)
Public Declare Function CQ_setGroupCard Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal QQID As Currency, ByVal newcard As Long) As Long

' ��Ⱥ�˳� ����, �˽ӿ���Ҫ�ϸ���Ȩ
' groupid Ŀ��Ⱥ
' isdismiss �Ƿ��ɢ true:��ɢ��Ⱥ(Ⱥ��) false:�˳���Ⱥ(����Ⱥ��Ա)
Public Declare Function CQ_setGroupLeave Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal isdismiss As Long) As Long

' ��Ⱥ��Աר��ͷ�� ��Ⱥ��Ȩ��
' groupid Ŀ��Ⱥ
' QQID Ŀ��QQ
' newspecialtitle ͷ�Σ����Ҫɾ����������գ�
' duration ר��ͷ����Ч�ڣ���λΪ�롣���������Ч��������д-1��
Public Declare Function CQ_setGroupSpecialTitle Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal QQID As Currency, ByVal newspecialtitle As Long, ByVal duration As Currency) As Long

' ���������˳�
' discussid Ŀ���������
Public Declare Function CQ_setDiscussLeave Lib "CQP" (ByVal AuthCode As Long, ByVal discussid As Currency) As Long

' �ú����������
' responseflag �����¼��յ��� responseflag ����
' responseoperation REQUEST_ALLOW �� REQUEST_DENY
' remark ��Ӻ�ĺ��ѱ�ע
Public Declare Function CQ_setFriendAddRequest Lib "CQP" (ByVal AuthCode As Long, ByVal responseflag As Long, ByVal responseoperation As Long, ByVal remark As Long) As Long

' ��Ⱥ�������
' responseflag �����¼��յ��� responseflag ����
' requesttype���������¼������������� REQUEST_GROUPADD �� REQUEST_GROUPINVITE
' responseoperation  REQUEST_ALLOW �� REQUEST_DENY
' reason �������ɣ��� REQUEST_GROUPADD �� REQUEST_DENY ʱ����
Public Declare Function CQ_setGroupAddRequestV2 Lib "CQP" (ByVal AuthCode As Long, ByVal responseflag As Long, ByVal requesttype As Long, ByVal responseoperation As Long, ByVal reason As Long) As Long

' ȡȺ��Ա��Ϣ
' groupid Ŀ��QQ����Ⱥ
' QQID Ŀ��QQ��
' nocache ��ʹ�û���
Public Declare Function CQ_getGroupMemberInfoV2 Lib "CQP" (ByVal AuthCode As Long, ByVal groupid As Currency, ByVal QQID As Currency, ByVal nocache As Long) As Long

' ȡİ������Ϣ
' QQID Ŀ��QQ
' nocache ��ʹ�û���
Public Declare Function CQ_getStrangerInfo Lib "CQP" (ByVal AuthCode As Long, ByVal QQID As Currency, ByVal nocache As Long) As Long

' ��־
' priority ���ȼ���CQLOG ��ͷ�ĳ���
' category ����
' content ����
Public Declare Function CQ_addLog Lib "CQP" (ByVal AuthCode As Long, ByVal priority As Long, ByVal category As Long, ByVal content As Long) As Long

' ȡCookies ����, �˽ӿ���Ҫ�ϸ���Ȩ
Public Declare Function CQ_getCookies Lib "CQP" (ByVal AuthCode As Long) As Long

' ȡCsrfToken ����, �˽ӿ���Ҫ�ϸ���Ȩ
Public Declare Function CQ_getCsrfToken Lib "CQP" (ByVal AuthCode As Long) As Long

' ȡ��¼QQ
Public Declare Function CQ_getLoginQQ Lib "CQP" (ByVal AuthCode As Long) As Long

' ȡ��¼QQ�ǳ�
Public Declare Function CQ_getLoginNick Lib "CQP" (ByVal AuthCode As Long) As Long

' ȡӦ��Ŀ¼�����ص�·��ĩβ��"\"
Public Declare Function CQ_getAppDirectory Lib "CQP" (ByVal AuthCode As Long) As Long

' ������������ʾ
' errorinfo ������Ϣ
Public Declare Function CQ_setFatal Lib "CQP" (ByVal AuthCode As Long, ByVal errorinfo As Long) As Long

' ����������������Ϣ�е�����(record),���ر����� \data\record\ Ŀ¼�µ��ļ���
' file �յ���Ϣ�е������ļ���(file)
' outformat Ӧ������������ļ���ʽ��Ŀǰ֧�� mp3 amr wma m4a spx ogg wav flac
Public Declare Function CQ_getRecord Lib "CQP" (ByVal AuthCode As Long, ByVal file As Long, ByVal outformat As Long) As Long
