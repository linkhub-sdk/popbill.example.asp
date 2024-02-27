<%
Class MessageService

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

Public Property Let UseGAIP(ByVal value)
    m_PopbillBase.UseGAIP = value
End Property

Public Property Let UseLocalTimeYN(ByVal value)
    m_PopbillBase.UseLocalTimeYN = value
End Property

Public Sub Class_Initialize
    Set m_PopbillBase = New PopbillBase
    m_PopbillBase.AddScope("150")
    m_PopbillBase.AddScope("151")
    m_PopbillBase.AddScope("152")
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
Public Function GetPopbillURL(CorpNum , UserID , TOGO)
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
Public Function GetChargeInfo ( CorpNum, MType, UserID )
    Dim result : Set result = m_PopbillBase.httpGET ( "/Message/ChargeInfo?Type=" &MType, m_PopbillBase.getSession_token(CorpNum), UserID )

    Dim chrgInfo : Set chrgInfo = New ChargeInfo
    chrgInfo.fromJsonInfo result

    Set GetChargeInfo = chrgInfo
End Function

'������ �Աݽ�û
Public Function PaymentRequest(CorpNum, PaymentForm, UserID)
    Set PaymentRequest = m_popbillBase.PaymentRequest(CorpNum, PaymentForm, UserID)
End Function

'����ȸ�� ����Ʈ �������� ��ȸ
Public Function GetPaymentHistory(CorpNum, SDate, EDate, Page, PerPage, UserID)
    Set GetPaymentHistory = m_popbillBase.GetPaymentHistory(CorpNum, SDate, EDate, Page, PerPage, UserID)
End Function

'����ȸ�� ������ �Աݽ�û ����Ȯ��
Public Function GetSettleResult(CorpNum, SettleCode, UserID)
    Set GetSettleResult = m_popbillBase.GetSettleResult(CorpNum, SettleCode, UserID)
End Function

'����ȸ�� ����Ʈ ��볻�� Ȯ��
Public Function GetUseHistory(CorpNum, SDate, EDate, Page, PerPage, Order, UserID)
    Set GetUseHistory = m_PopbillBase.GetUseHistory(CorpNum, SDate, EDate, Page, PerPage, Order, UserID)
End Function

'����ȸ�� ����Ʈ ȯ�ҽ�û
Public Function Refund(CorpNum, RefundForm, UserID)
    Set Refund = m_popbillBase.Refund(CorpNum, RefundForm, UserID)
End Function

' ȯ�� ���� ����Ʈ ��ȸ
Public Function GetRefundableBalance(CorpNum, UserID)
    GetRefundableBalance = m_popbillBase.GetRefundableBalance(CorpNum, UserID)
End Function

'����ȸ�� ����Ʈ ȯ�ҳ��� Ȯ��
Public Function GetRefundHistory(CorpNum, Page, PerPage, UserID)
    Set GetRefundHistory = m_popbillBase.GetRefundHistory(CorpNum, Page, PerPage, UserID)
End Function

' ȯ�� ��û ���� ��ȸ
Public Function GetRefundInfo(CorpNum, RefundCode, UserID)
    Set GetRefundInfo = m_popbillBase.GetRefundInfo(CorpNum, RefundCode, UserID)
End Function

'ȸ�� Ż��
Public Function QuitMember(CorpNum, QuitReason, UserID)
    Set QuitMember = m_popbillBase.QuitMember(CorpNum, QuitReason, UserID)
End Function

'''''''''''''  End of PopbillBase

''�ܰ�Ȯ��
Public Function GetUnitCost(CorpNum, MType)
    Dim result : Set result = m_PopbillBase.httpGET("/Message/UnitCost?Type="&MType, m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'�߽Ź�ȣ ��Ͽ��� Ȯ��
Public Function CheckSenderNumber(CorpNum, SenderNumber, UserID)
    If SenderNumber = "" Or IsNull(SenderNumber) Then
        Err.Raise -99999999, "POPBILL", "�߽Ź�ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Set CheckSenderNumber = m_PopbillBase.httpGET("/Message/CheckSenderNumber/"&SenderNumber,m_PopbillBase.getSession_token(CorpNum),UserID)
End Function

'�ܹ� �޽��� ���� (RequestNum)
Public Function SendSMS(CorpNum, sender, Contents, Messages, reserveDT, adsYN, RequestNum, UserID)
    SendSMS = SendMessage("SMS", CorpNum, sender, "", Contents, Messages, reserveDT, adsYN, RequestNum, UserID)
End Function

'�幮 �޽��� ���� (RequestNum)
Public Function SendLMS(CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, RequestNum, UserID)
    SendLMS = SendMessage("LMS", CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, RequestNum, UserID)
End Function

'��/�幮 �޽��� �ڵ��ν� ����(RequestNum)
Public Function SendXMS(CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, RequestNum, UserID)
    SendXMS = SendMessage("XMS", CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, RequestNum, UserID)
End Function

'MMS �޽��� ���� (RequestNum)
Public Function SendMMS(CorpNum, sender, subject, Contents, msgList, FilePaths, reserveDT, adsYN, RequestNum, UserID)
    If IsNull(msgList) Or IsEmpty(msgList) Then
        Err.raise -99999999, "POPBILL", "������ �޽����� �Էµ��� �ʾҽ��ϴ�."
    End If

    If isNull(FilePaths) Then Err.Raise -99999999, "POPBILL", "������ ���ϰ�ΰ� �Էµ��� �ʾҽ��ϴ�."

    Dim tmp : Set tmp = JSON.parse("{}")

    If sender <> "" Then tmp.Set "snd", sender
    If Contents <> "" Then tmp.Set "content", Contents
    If subject <> "" Then tmp.Set "subject", subject
    If reserveDT <> "" Then tmp.Set "sndDT", reserveDT
    If adsYN <> "" Then tmp.Set "adsYN", adsYN
    If RequestNum <> "" Then tmp.Set "RequestNum", RequestNum

    Dim msgs : Set msgs = JSON.parse("[]")

    Dim i
    For i=0 To msgList.Count-1
        Dim msgObj : Set msgObj = New Messages
        msgObj.setValue msgList.Item(i)
        msgs.Set i, msgObj.toJsonInfo
    Next

    tmp.Set "msgs", msgs

    Dim postData : postData = m_PopbillBase.toString(tmp)

    Dim result : Set result = m_PopbillBase.httpPost_Files("/MMS", m_PopbillBase.getSession_Token(CorpNum), postData, FilePaths, UserID)
    SendMMS = result.ReceiptNum
End Function



Private Function SendMessage(MType, CorpNum, sender, subject, Contents, msgList, reserveDT, adsYN, RequestNum, UserID)
    If IsNull(msgList) Or IsEmpty(msgList) Then
        Err.raise -99999999, "POPBILL", "������ �޽����� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim tmp : Set tmp = JSON.parse("{}")

    If sender <> "" Then tmp.Set "snd", sender
    If Contents <> "" Then tmp.Set "content", Contents
    If subject <> "" Then tmp.Set "subject", subject
    If reserveDT <> "" Then tmp.Set "sndDT", reserveDT
    If RequestNum <> "" Then tmp.Set "RequestNum", RequestNum
    If adsYN Then tmp.Set "adsYN", adsYN


    Dim msgs : Set msgs = JSON.parse("[]")

    Dim i
    For i=0 To msgList.Count-1
        Dim msgObj : Set msgObj = New Messages
        msgObj.setValue msgList.Item(i)
        msgs.Set i, msgObj.toJsonInfo
    Next

    tmp.Set "msgs", msgs

    Dim postData : postData = m_PopbillBase.toString(tmp)

    Set result = m_PopbillBase.httpPost("/"&MType, m_PopbillBase.getSession_Token(CorpNum), "", postData, UserID)
    SendMessage = result.ReceiptNum

End Function

'���๮�� �������
Public Function CancelReserve(CorpNum, ReceiptNum, UserID)
    If ReceiptNum = "" Or IsNull(ReceiptNum) Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�"
    End If

    Set CancelReserve = m_PopbillBase.httpGet("/Message/"&ReceiptNum&"/Cancel",m_PopbillBase.getSession_token(CorpNum),UserID)
End Function

'���๮�� ������� (��û��ȣ �Ҵ�)
Public Function CancelReserveRN(CorpNum, RequestNum, UserID)
    If RequestNum = "" Or IsNull(RequestNum) Then
        Err.Raise -99999999, "POPBILL", "��û��ȣ�� �Էµ��� �ʾҽ��ϴ�"
    End If

    Set CancelReservRN = m_PopbillBase.httpGet("/Message/Cancel/"&RequestNum, m_PopbillBase.getSession_token(CorpNum), UserID)
End Function

'���๮�� ������� (������ȣ, ���Ź�ȣ)
Public Function CancelReservebyRCV(CorpNum, ReceiptNum, ReceiveNum, UserID)
    If ReceiptNum = "" Or IsNull(ReceiptNum) Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�"
    End If
    If ReceiveNum = "" Or IsNull(ReceiveNum) Then
        Err.Raise -99999999, "POPBILL", "���Ź�ȣ�� �Էµ��� �ʾҽ��ϴ�"
    End If

    Dim postData : postData = m_PopbillBase.toString(ReceiveNum)

    set CancelReservebyRCV = m_PopbillBase.httpPost("/Message/"&ReceiptNum&"/Cancel", m_PopbillBase.getSession_token(CorpNum), "", postData, UserID)
End Function

    '���๮�� ������� (��û��ȣ, ���Ź�ȣ)
Public Function CancelReserveRNbyRCV(CorpNum, RequestNum, ReceiveNum, UserID)
    If RequestNum = "" Or IsNull(RequestNum) Then
    Err.Raise -99999999, "POPBILL", "��û��ȣ�� �Էµ��� �ʾҽ��ϴ�"
    End If
    If ReceiveNum = "" Or IsNull(ReceiveNum) Then
    Err.Raise -99999999, "POPBILL", "���Ź�ȣ�� �Էµ��� �ʾҽ��ϴ�"
    End If

    Dim postData : postData = m_PopbillBase.toString(ReceiveNum)

    Set CancelReserveRNbyRCV = m_PopbillBase.httpPost("/Message/Cancel/"&RequestNum, m_PopbillBase.getSession_token(CorpNum), "", postData, UserID)
End Function

'���� ���� URL
Public Function GetURL(CorpNum, UserID, TOGO)
    Dim result : Set result = m_PopbillBase.httpGet("/Message/?TG="+TOGO,m_PopbillBase.getSession_token(CorpNum), UserID)
    GetURL = result.url
End Function

'���� ���۳��� �˾� URL
Public Function GetSentListURL(CorpNum, UserID)
    Dim result : Set result = m_PopbillBase.httpGet("/Message/?TG=BOX",m_PopbillBase.getSession_token(CorpNum), UserID)
    GetSentListURL = result.url
End Function

'�߽Ź�ȣ ���� �˾� URL
Public Function GetSenderNumberMgtURL(CorpNum, UserID)
    Dim result : Set result = m_PopbillBase.httpGet("/Message/?TG=SENDER",m_PopbillBase.getSession_token(CorpNum), UserID)
    GetSenderNumberMgtURL = result.url
End Function



'���� ���۳��� Ȯ��
Public Function GetMessages(CorpNum, ReceiptNum, UserID)
    If ReceiptNum = "" Or IsNull(ReceptNum) Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�"
    End If

    Dim result : Set result = m_PopbillBase.httpGet("/Message/"&ReceiptNum,m_PopbillBase.getSession_token(CorpNum),UserID)

    Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")

    Dim i
    For i=0 To result.length-1
        Dim msgInfo : Set msgInfo = New MessageInfo
        msgInfo.fromJsonInfo result.Get(i)
        tmp.Add i, msgInfo
    Next

    Set GetMessages = tmp

End Function


'���� ���۳��� Ȯ�� (��û��ȣ �Ҵ�)
Public Function GetMessagesRN(CorpNum, RequestNum, UserID)
    If RequestNum = "" Or IsNull(RequestNum) Then
        Err.Raise -99999999, "POPBILL", "��û��ȣ�� �Էµ��� �ʾҽ��ϴ�"
    End If

    Dim result : Set result = m_PopbillBase.httpGet("/Message/Get/"+RequestNum ,m_PopbillBase.getSession_token(CorpNum), UserID)

    Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")

    Dim i
    For i=0 To result.length-1
        Dim msgInfo : Set msgInfo = New MessageInfo
        msgInfo.fromJsonInfo result.Get(i)
        tmp.Add i, msgInfo
    Next

    Set GetMessagesRN = tmp

End Function

'���۳��� ������� Ȯ��
Public Function GetStates(CorpNum, ReceiptNumList, UserID)
    If IsNull(ReceiptNumList) Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�"
    End If

    Dim tmp : Set tmp = JSON.parse("[]")

    Dim i
    For i=0 To UBound(ReceiptNumList) - 1
        tmp.Set i, ReceiptNumList(i)
    Next

    Dim postData : postData = m_PopbillBase.toString(tmp)

    Dim result : Set result = m_PopbillBase.httpPost("/Message/States", m_PopbillBase.getSession_token(CorpNum), "", postData, UserID)

    Dim infoObj : Set infoObj = CreateObject("Scripting.Dictionary")

    For i=0 To result.length-1
        Dim infoTmp : Set infoTmp = New MessageBriefInfo
        infoTmp.fromJsonInfo result.Get(i)
        infoObj.Add i, infoTmp
    Next

    Set GetStates = infoObj

End Function

'�������۳��� ��ȸ
Public Function Search(CorpNum, SDate, EDate, State, Item, ReserveYN, SenderYN, Order, Page, PerPage, QString)
    If SDate = "" Then
        Err.Raise -99999999, "POPBILL", "�������ڰ� �Էµ��� �ʾҽ��ϴ�."
    End If
    If EDate = "" Then
        Err.Raise -99999999, "POPBILL", "�������ڰ� �̷µ��� �ʾҽ��ϴ�."
    End If
    Dim uri
    uri = "/Message/Search"
    uri = uri & "?SDate=" & SDate
    uri = uri & "&EDate=" & EDate

    Dim i
    uri = uri & "&State="
    For i=0 To UBound(State) -1
        If i = UBound(State) -1 then
            uri = uri & State(i)
        Else
            uri = uri & State(i) & ","
        End If
    Next

    uri = uri & "&Item="
    For i=0 To UBound(Item) -1
        If i = UBound(Item) -1  then
            uri = uri & Item(i)
        Else
            uri = uri & Item(i) & ","
        End If
    Next

    If ReserveYN Then
        uri = uri & "&ReserveYN=1"
    Else
        uri = uri & "&ReserveYN=0"
    End If

    If SenderYN Then
        uri = uri & "&SenderYN=1"
    Else
        uri = uri & "&SenderYN=0"
    End If

    uri = uri & "&Order=" & Order
    uri = uri & "&Page=" & CStr(Page)
    uri = uri & "&PerPage=" & CStr(PerPage)
    uri = uri & "&QString=" & Server.URLEncode(QString)

    Dim searchResult : Set searchResult = New MSGSearchResult
    Dim tmpObj : Set tmpObj = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), "")

    searchResult.fromJsonInfo tmpObj

    Set Search = searchResult
End Function

'080 ���Űźθ�� Ȯ��
Public Function GetAutoDenyList(CorpNum)
    Set GetAutoDenyList = m_PopbillBase.httpGET("/Message/Denied", m_PopbillBase.getSession_token(CorpNum), "")
End Function

' �߽Ź�ȣ ��� Ȯ��
Public Function GetSenderNumberList(CorpNum)
    Set GetSenderNumberList = m_PopbillBase.httpGET("/Message/SenderNumber", m_PopbillBase.getSession_token(CorpNum), "")
End Function

'080 ��ȣ Ȯ��
Public Function CheckAutoDenyNumber (CorpNum, UserID)
    Dim m_AutoDenyNumberInfo : Set m_AutoDenyNumberInfo = New AutoDenyNumberInfo
    Dim tmp : Set tmp = m_PopbillBase.httpGet("/Message/AutoDenyNumberInfo", m_popbillBase.getSession_token(CorpNum), UserID)

    m_AutoDenyNumberInfo.fromJsonInfo(tmp)

    Set CheckAutoDenyNumber = m_AutoDenyNumberInfo
End Function

End Class



Class Messages
    Public sender
    Public senderName
    Public receiver
    Public receiverName
    Public content
    Public subject
    Public interOPRefKey

    Public Sub setValue(msgList)
        sender = msgList.sender
        senderName = msgList.senderName
        receiver = msgList.receiver
        receiverName = msgList.receiverName
        content = msgList.content
        subject = msgList.subject
        interOPRefKey = msgList.interOPRefKey
    End Sub

    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.set "rcv", receiver
        If sender <> "" Then  toJsonInfo.set "snd", sender
        If senderName <> "" Then  toJsonInfo.set "sndnm", senderName
        If receiverName <> "" Then toJsonInfo.set "rcvnm", receiverName
        If content <> "" Then toJsonInfo.set "msg", content
        If subject <> "" Then toJsonInfo.set "sjt", subject
        If interOPRefKey <> "" Then toJsonInfo.set "interOPRefKey", interOPRefKey
    End Function

End Class


Class MessageBriefInfo
    Public sn
    Public rNum
    Public stat
    Public sDT
    Public rDT
    Public rlt
    Public net
    Public srt

    Public Sub fromJsonInfo(briefInfo)
        On Error Resume Next
        sn = briefInfo.sn
        rNum = briefInfo.rNum
        stat = briefInfo.stat
        sDT = briefInfo.sDT
        rDT = briefInfo.rDT
        rlt = briefInfo.rlt
        net = briefInfo.net
        srt = briefInfo.srt
        On Error GoTo 0
    End Sub
End Class


Class MessageInfo
    Public state
    Public result
    Public subject
    Public msgType
    Public content
    Public sendNum
    Public senderName
    Public ReceiveNum
    Public receiveName
    Public reserveDT
    Public sendDT
    Public resultDT
    Public sendResult
    Public tranNet
    Public receiptDT
    Public RequestNum
    Public ReceiptNum
    Public interOPRefKey

    Public Sub fromJsonInfo(msgInfo)
        On Error Resume Next
        state = msgInfo.state
        result = msgInfo.result
        subject = msgInfo.subject
        msgType = msgInfo.type
        content = msgInfo.content
        sendNum = msgInfo.sendNum
        senderName = msgInfo.senderName
        ReceiveNum = msgInfo.ReceiveNum
        receiveName = msgInfo.receiveName
        reserveDT = msgInfo.reserveDT
        sendDT = msgInfo.sendDT
        resultDT = msgInfo.resultDT
        sendResult = msgInfo.sendResult
        tranNet = msgInfo.tranNet
        receiptDT = msgInfo.receiptDT
        RequestNum = msgInfo.RequestNum
        ReceiptNum = msgInfo.ReceiptNum
        interOPRefKey = msgInfo.interOPRefKey
        On Error GoTo 0
    End Sub
End Class


Class MSGSearchResult
    Public code
    Public total
    Public perPage
    Public pageNum
    Public pageCount
    Public message
    Public list()

    Public Sub Class_Initialize
        ReDim list(-1)
    End Sub

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
        code = jsonInfo.code
        total = jsonInfo.total
        perPage = jsonInfo.perPage
        pageNum = jsonInfo.pageNum
        pageCount = jsonInfo.pageCount
        message = jsonInfo.message

        ReDim list(jsonInfo.list.length)
        Dim i
        For i = 0 To jsonInfo.list.length -1
            Dim tmpObj : Set tmpObj = New MessageInfo
            tmpObj.fromJsonInfo jsonInfo.list.Get(i)
            Set list(i) = tmpObj
        Next

        On Error GoTo 0
    End Sub
End Class

Class AutoDenyNumberInfo
    Public smsdenyNumber
    Public regDT

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
            smsdenyNumber = jsonInfo.smsdenyNumber
            regDT = jsonInfo.regDT
        On Error GoTo 0
    End Sub
End Class
%>
