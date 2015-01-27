<%
Class MessageService

Private m_PopbillBase

'�׽�Ʈ �÷���
Public Property Let IsTest(ByVal value)
    m_PopbillBase.IsTest = value
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
'ȸ������ ����
Public Function CheckIsMember(CorpNum , linkID)
    Set CheckIsMember = m_PopbillBase.CheckIsMember(CorpNum,linkID)
End Function
'ȸ������
Public Function JoinMember(JoinInfo)
    Set JoinMember = m_PopbillBase.JoinMember(JoinInfo)
End Function
'''''''''''''  End of PopbillBase

''�ܰ�Ȯ��
Public Function GetUnitCost(CorpNum, MType)
    Set result = m_PopbillBase.httpGET("/Message/UnitCost?Type="&MType, m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'�ܹ� �޽��� ����
Public Function SendSMS(CorpNum, sender, Contents, Messages, reserveDT, UserID)
	SendSMS = SendMessage("SMS", CorpNum, sender, "", Contents, Messages, reserveDT, UserID)	
End Function

'�幮 �޽��� ����
Public Function SendLMS(CorpNum, sender, subject, Contents, Messages, reserveDT, UserID)
	SendLMS = SendMessage("LMS", CorpNum, sender, subject, Contents, Messages, reserveDT, UserID)	
End Function

'��/�幮 �޽��� �ڵ��ν� ����
Public Function SendXMS(CorpNum, sender, subject, Contents, Messages, reserveDT, UserID)
	SendXMS = SendMessage("XMS", CorpNum, sender, subject, Contents, Messages, reserveDT, UserID)	
End Function

Private Function SendMessage(MType, CorpNum, sender, subject, Contents, msgList, reserveDT, UserID)
	If IsNull(msgList) Or IsEmpty(msgList) Then 
		Err.raise -99999999, "POPBILL", "������ �޽����� �Էµ��� �ʾҽ��ϴ�."
	End If

	Set tmp = JSON.parse("{}")
	    
    If sender <> "" Then tmp.Set "snd", sender
    If Contents <> "" Then tmp.Set "content", Contents
    If subject <> "" Then tmp.Set "subject", subject
    If reserveDT <> "" Then tmp.Set "sndDT", reserveDT

	Set msgs = JSON.parse("[]")

	For i=0 To msgList.Count-1
		Set msgObj = New Messages
		msgObj.setValue msgList.Item(i)
		msgs.Set i, msgObj.toJsonInfo
	Next

	tmp.Set "msgs", msgs

	postdata = m_PopbillBase.toString(tmp)

	Set result = m_PopbillBase.httpPost("/"&MType, m_PopbillBase.getSession_Token(CorpNum), "", postdata, UserID)
	SendMessage = result.receiptNum

End Function

'���๮�� �������
Public Function CancelReserve(CorpNum, ReceiptNum, UserID)
	If ReceiptNum = "" Or IsNull(ReceptNum) Then 
		Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�"
	End If
	
	Set CancelReserve = m_PopbillBase.httpGet("/Message/"&ReceiptNum&"/Cancel",m_PopbillBase.getSession_token(CorpNum),UserID)
End Function

'���� ���� URL
Public Function GetURL(CorpNum, UserID, TOGO)
	Set result = m_PopbillBase.httpGet("/Message/?TG="+TOGO,m_PopbillBase.getSession_token(CorpNum), UserID)
	GetURL = result.url
End Function


'���� ���۳��� Ȯ��
Public Function GetMessages(CorpNum, ReceiptNum, UserID)
	If ReceiptNum = "" Or IsNull(ReceptNum) Then 
		Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�"
	End If
	
	Set result = m_PopbillBase.httpGet("/Message/"&ReceiptNum,m_PopbillBase.getSession_token(CorpNum),UserID)
	
	Set tmp = CreateObject("Scripting.Dictionary")

	For i=0 To result.length-1
		Set msgInfo = New MessageInfo
		msgInfo.fromJsonInfo result.Get(i)
		tmp.Add i, msgInfo
	Next

	Set GetMessages = tmp

End Function 
End Class

Class Messages
Public sender
Public receiver
Public receiverName
Public content
Public subject

Public Sub setValue(msgList)
	sender = msgList.sender
	receiver = msgList.receiver
	receiverName = msgList.receiverName
	content = msgList.content
	subject = msgList.subject
End Sub

Public Function toJsonInfo()
	Set toJsonInfo = JSON.parse("{}")
	toJsonInfo.set "rcv", receiver
	If sender <> "" Then  toJsonInfo.set "snd", sender
	If receiverName <> "" Then toJsonInfo.set "rcvnm", receiverName
	If content <> "" Then toJsonInfo.set "msg", content
	If subject <> "" Then toJsonInfo.set "sjt", subject
End Function

End Class

Class MessageInfo
Public state
Public subject
Public msgType
Public content
Public sendNum
Public receiveNum
Public receiveName
Public reserveDT
Public sendDT
Public resultDT
Public sendResult

Public Sub fromJsonInfo(msgInfo)
	On Error Resume Next
	state = msgInfo.state
	subject = msgInfo.subject
	msgType = msgInfo.type
	content = msgInfo.content
	sendNum = msgInfo.sendNum
	receiveNum = msgInfo.receiveNum
	receiveName = msgInfo.receiveName
	reserveDT = msgInfo.reserveDT
	sendDT = msgInfo.sendDT
	resultDT = msgInfo.resultDT
	sendResult = msgInfo.sendResult
	On Error GoTo 0
End Sub
End Class

%>