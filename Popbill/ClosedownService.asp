<%
Class ClosedownService

Private m_PopbillBase

'�׽�Ʈ �÷���
Public Property Let IsTest(ByVal value)
    m_PopbillBase.IsTest = value
End Property

Public Property Let IPRestrictOnOff(ByVal value)
    m_PopbillBase.IPRestrictOnOff = value
End Property

Public Property Let UseStaticIP(ByVal value)
    m_PopbillBase.UseStaticIP = value
End Property

Public Property Let UseLocalTimeYN(ByVal value)
    m_PopbillBase.UseLocalTimeYN = value
End Property

Public Sub Class_Initialize
    Set m_PopbillBase = New PopbillBase
    m_PopbillBase.AddScope("170")
End Sub

Public Sub Initialize(linkID, SecretKey )
    m_PopbillBase.Initialize linkID,SecretKey
End Sub

'ȸ���ܾ���ȸ
Public Function GetBalance(CorpNum)
    GetBalance = m_PopbillBase.GetBalance(CorpNum)
End Function
'��Ʈ�� �ܾ���ȸ
Public Function GetPartnerBalance(CorpNum)
    GetPartnerBalance = m_PopbillBase.GetPartnerBalance(CorpNum)
End Function
'�˺� �⺻ URL
Public Function GetPopbillURL(CorpNum , UserID , TOGO )
    GetPopbillURL = m_PopbillBase.GetPopbillURL(CorpNum , UserID , TOGO )
End Function
'�˺� �α��� URL
Public Function GetAccessURL(CorpNum , UserID)
    GetAccessURL = m_PopbillBase.GetAccessURL(CorpNum , UserID )
End Function

'�˺� ����ȸ�� ����Ʈ ���� URL
Public Function GetChargeURL(CorpNum , UserID)
    GetChargeURL = m_PopbillBase.GetChargeURL(CorpNum , UserID )
End Function

'�˺� ����ȸ�� ����Ʈ �������� URL
Public Function GetPaymentURL(CorpNum, UserID)
    GetPaymentURL = m_PopbillBase.GetPaymentURL(CorpNum, UserID)
End Function

'�˺� ����ȸ�� ����Ʈ ��볻�� URL
Public Function GetUseHistoryURL(CorpNum, UserID)
    GetUseHistoryURL = m_PopbillBase.GetUseHistoryURL(CorpNum, UserID)
End Function

'��Ʈ�� ����Ʈ ���� �˾� URL - 2017/08/29 �߰�
Public Function GetPartnerURL(CorpNum, TOGO)
    GetPartnerURL = m_PopbillBase.GetPartnerURL(CorpNum,TOGO)
End Function

'ȸ������ ����
Public Function CheckIsMember(CorpNum , linkID)
    Set CheckIsMember = m_PopbillBase.CheckIsMember(CorpNum,linkID)
End Function
'ȸ������
Public Function JoinMember(JoinInfo)
    Set JoinMember = m_PopbillBase.JoinMember(JoinInfo)
End Function
'����� ���� Ȯ��
Public Function GetContactInfo(CorpNum, ContactID, UserID)
    Set GetContactInfo = m_PopbillBase.GetContactInfo(CorpNum, ContactID, UserID)
End Function 
'����� �����ȸ
Public Function ListContact(CorpNum, UserID)
    Set ListContact = m_popbillBase.ListContact(CorpNum,UserID)
End Function
'����� ��������
Public Function UpdateContact(CorpNum, contInfo, UserId)
    Set UpdateContact = m_popbillBase.UpdateContact(CorpNum, contInfo, UserId)
End Function
'����� �߰� 
Public Function RegistContact(CorpNum, contInfo, UserId)
    Set RegistContact = m_popbillBase.RegistContact(CorpNum, contInfo, UserId)
End Function
'ȸ������ ����
Public Function UpdateCorpInfo(CorpNum, corpInfo, UserId)
    Set UpdateCorpInfo = m_popbillBase.UpdateCorpInfo(CorpNum, corpInfo, UserId)
End Function
'ȸ������ Ȯ�� 
Public Function GetCorpInfo(CorpNum, UserId)
    Set GetCorpInfo = m_popbillBase.GetCorpInfo(CorpNum, UserId)
End Function
Public Function CheckID(id)
    Set CheckID = m_popbillBase.CheckID(id)
End Function
'�������� Ȯ��
Public Function GetChargeInfo ( CorpNum, UserID )
    Dim result : Set result = m_PopbillBase.httpGET("/CloseDown/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)

    Dim chrgInfo : Set chrgInfo = New ChargeInfo
    chrgInfo.fromJsonInfo result
    
    Set GetChargeInfo = chrgInfo
End Function
'''''''''''''  End of PopbillBase

'��ȸ�ܰ�Ȯ��
Public Function GetUnitCost(CorpNum)
    Dim result : Set result = m_PopbillBase.httpGET("/CloseDown/UnitCost", m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'�������ȸ �ܰ�
Public Function CheckCorpNum(MemberCorpNum, CorpNum) 
    If MemberCorpNum = "" Or isEmpty(MemberCorpNum) Then 
        Err.Raise -99999999, "POPBILL", "�˺�ȸ�� ����ڹ�ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    If CorpNum = "" Or isEmpty(CorpNum) Then 
        Err.Raise -99999999, "POPBILL", "��ȸ�� ����ڹ�ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    On Error Resume Next
    
    Dim result : Set result = m_PopbillBase.httpGet("/CloseDown?CN="+CorpNum, m_PopbillBase.getSession_token(MemberCorpNum),"")
    
    Dim stateObj : Set stateObj = New CorpState
    stateObj.fromJsonInfo result
    Set CheckCorpNum = stateObj
    
    On Error Resume Next
End Function 

'�������ȸ �뷮
Public Function CheckCorpNums(MemberCorpNum, CorpNumList)
    If MemberCorpNum = "" Or isEmpty(MemberCorpNum) Then 
        Err.Raise -99999999, "POPBILL", "�˺�ȸ�� ����ڹ�ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    If isNull(CorpNumList) Or isEmpty(CorpNumList) Then 
        Err.Raise -99999999, "POPBILL", "��ȸ�� ����ڹ�ȣ �迭�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim tmp : Set tmp = JSON.parse("[]")

    Dim i
    For i=0 To UBound(CorpNumList)-1
        tmp.Set i, CorpNumList(i)
    Next

    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    
    Dim result : Set result = m_PopbillBase.httpPOST("/CloseDown", m_PopbillBase.getSession_token(MemberCorpNum), "", postdata, "")

    Dim tmpDic : Set tmpDic = CreateObject("Scripting.Dictionary")

    For i=0 To result.length-1
        Dim stateObj : Set stateObj = New CorpState 
        stateObj.fromJsonInfo result.Get(i)
        tmpDic.Add i, stateObj
    Next

    Set CheckCorpNums = tmpDic
End Function
End Class

Class CorpState
Public corpNum
Public state
Public ctype
Public stateDate
Public checkDate
Public typeDate '�������� �������� �߰� 2017/08/17

Public Sub fromJsonInfo(jsonInfo)
    On Error Resume Next
        If Not isEmpty(jsonInfo.corpNum) Then
            corpNum = jsonInfo.corpNum
        End If 
        If Not isEmpty(jsonInfo.type) Then
            ctype = jsonInfo.type
        End If 
        If Not isEmpty(jsonInfo.state) Then
            state = jsonInfo.state
        End If 
        If Not isEmpty(jsonInfo.stateDate) Then
            stateDate = jsonInfo.stateDate
        End If 
        If Not isEmpty(jsonInfo.checkDate) Then
            checkDate = jsonInfo.checkDate
        End If 
        If Not isEmpty(jsonInfo.typeDate) Then
            typeDate = jsonInfo.typeDate
        End If 
    On Error GoTo 0
End Sub
End Class


%>