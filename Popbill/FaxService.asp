<%
Class FaxService

Private m_PopbillBase

'�׽�Ʈ �÷���
Public Property Let IsTest(ByVal value)
    m_PopbillBase.IsTest = value
End Property

Public Sub Class_Initialize
	Set m_PopbillBase = New PopbillBase
	m_PopbillBase.AddScope("160")
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
'ȸ������ ����
Public Function CheckIsMember(CorpNum , linkID)
    Set CheckIsMember = m_PopbillBase.CheckIsMember(CorpNum,linkID)
End Function
'ȸ������
Public Function JoinMember(JoinInfo)
    Set JoinMember = m_PopbillBase.JoinMember(JoinInfo)
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
	Set result = m_PopbillBase.httpGET("/FAX/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)

	Set chrgInfo = New ChargeInfo
	chrgInfo.fromJsonInfo result
	
	Set GetChargeInfo = chrgInfo
End Function
'''''''''''''  End of PopbillBase

''�ܰ�Ȯ��
Public Function GetUnitCost(CorpNum)
    Set result = m_PopbillBase.httpGET("/FAX/UnitCost", m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'�ѽ� ���۳�����ȸ URL
Public Function GetURL(CorpNum, UserID, TOGO)
    Set result = m_PopbillBase.httpGET("/FAX/?TG="+TOGO, m_PopbillBase.getSession_token(CorpNum),UserID)
	GetURL = result.url
End Function


'�ѽ� �������� ���
Public Function CancelReserve(CorpNum, ReceiptNum, UserID)
	If isNull(ReceiptNum) Or IsEmpty(ReceiptNum) Then Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."

    Set CancelReserve = m_PopbillBase.httpGET("/FAX/"&ReceiptNum&"/Cancel", m_PopbillBase.getSession_token(CorpNum),UserID)
End Function

'�ѽ� ����
Public Function SendFAX(CorpNum , sendNum , receivers , FilePaths ,  reserveDT , UserID, adsYN )
	If isNull(receivers) Or IsEmpty(receivers) Then Err.Raise -99999999, "POPBILL", "���������� �� �Էµ��� �ʾҽ��ϴ�."
    If UBound(receivers) < 0 Then Err.Raise -99999999, "POPBILL", "���������� �� �Էµ��� �ʾҽ��ϴ�."
    If isNull(FilePaths) Or IsEmpty(FilePaths) Then Err.Raise -99999999, "POPBILL", "������ ���ϰ�ΰ� �Էµ��� �ʾҽ��ϴ�."
    If UBound(FilePaths) < 0 Then Err.Raise -99999999, "POPBILL", "������ ���ϰ�ΰ� �Էµ��� �ʾҽ��ϴ�."
    If UBound(FilePaths) >= 5 Then Err.Raise -99999999, "POPBILL", "1ȸ ���� ������ ���ϰ����� 5���Դϴ�."
  
    Set Form = JSON.parse("{}")
    
    Form.set "snd", sendNum
    If reserveDT <> "" Then Form.set "sndDT", reserveDT
  	If adsYN Then Form.Set "adsYN", adsYN  

    Form.set "fCnt", UBound(FilePaths) + 1
    
	Dim tmpArray() : ReDim tmpArray(UBound(receivers))
	For i = 0 to UBound(receivers)
		If  isNull(receivers(i)) Or IsEmpty(receivers(i)) Then
			Err.Raise -99999999, "POPBILL", CStr(i+1) & " ��° ������ ������ ������� �ʾҽ��ϴ�."
		else
			Set tmpArray(i) =  receivers(i).toJsonInfo()
		End if
	Next
    
    Form.set "rcvs", tmpArray
    
    postdata = m_PopbillBase.toString(Form)
    Set result = m_PopbillBase.httpPOST_Files("/FAX", m_PopbillBase.getSession_token(CorpNum), postdata, FilePaths, UserID)
    
    SendFAX = result.receiptNum
End Function


'�ѽ� ������
Public Function ResendFAX(CorpNum, receiptNum, sendNum, senderName, receivers,  reserveDT , UserID)
    If isNull(receiptNum) Or IsEmpty(receiptNum) Then Err.Raise -99999999, "POPBILL", "�ѽ� ������ȣ(receiptNum)�� �Էµ��� �ʾҽ��ϴ�."

    Set Form = JSON.parse("{}")
    
	If sendNum <> "" Then Form.set "snd", sendNum
	If senderName <> "" Then Form.set "sndnm", senderName
    If reserveDT <> "" Then Form.set "sndDT", reserveDT

	If UBound(receivers) >= 0 Then 
		Dim tmpArray() : ReDim tmpArray(UBound(receivers))
		For i = 0 to UBound(receivers)
			If  isNull(receivers(i)) Or IsEmpty(receivers(i)) Then
				Err.Raise -99999999, "POPBILL", CStr(i+1) & " ��° ������ ������ ������� �ʾҽ��ϴ�."
			else
				Set tmpArray(i) =  receivers(i).toJsonInfo()
			End if
		Next
		Form.set "rcvs", tmpArray
	End If 
	
    postdata = m_PopbillBase.toString(Form)
    Set result = m_PopbillBase.httpPOST("/FAX/"&receiptNum, m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)

    ResendFAX = result.receiptNum
End Function


'�ѽ� ���۰�� Ȯ��
Public Function GetFaxDetail(CorpNum, receiptNum, UserID)
	If  isEmpty(receiptNum) Then
			Err.Raise -99999999, "POPBILL", "�ѽ� ������ȣ(receiptNum)�� �Էµ��� �ʾҽ��ϴ�."
	End If

	Set result = m_PopbillBase.httpGET("/FAX/"&receiptNum, m_PopbillBase.getSession_token(CorpNum),UserID)
		
	Set tmp = CreateObject("Scripting.Dictionary")

	For i=0 To result.length-1
		Set faxInfo = New FaxState
		faxInfo.fromJsonInfo result.Get(i)
		tmp.Add i, faxInfo
	Next

	Set GetFaxDetail = tmp

End Function 

'�ѽ� ��� ��ȸ
Public Function Search(CorpNum, SDate, EDate, State, ReserveYN, SenderOnlyYN, Order, Page, PerPage)
	If SDate = "" Then
        Err.Raise -99999999, "POPBILL", "�������ڰ� �Էµ��� �ʾҽ��ϴ�."
	End If
	If EDate = "" Then
        Err.Raise -99999999, "POPBILL", "�������ڰ� �̷µ��� �ʾҽ��ϴ�."
	End If

	uri = "/FAX/Search"
	uri = uri & "?SDate=" & SDate
	uri = uri & "&EDate=" & EDate

	uri = uri & "&State="
	For i=0 To UBound(State) -1	
		If i = UBound(State) -1 then
			uri = uri & State(i)
		Else
			uri = uri & State(i) & ","
		End If
	Next
	
	If ReserveYN Then uri = uri & "&ReserveYN=1"
	If SedernOnlyYN Then uri = uri & "&SenderOnlyYN=1"
	
	uri = uri & "&Order=" & Order
	uri = uri & "&Page=" & CStr(Page)
	uri = uri & "&PerPage=" & CStr(PerPage)
	
	Set searchResult = New FAXSearchResult
	Set tmpObj = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), "")

	searchResult.fromJsonInfo tmpObj
	
	Set Search = searchResult
End Function

' �߽Ź�ȣ ��� Ȯ��
Public Function GetSenderNumberList(CorpNum)
	Set GetSenderNumberList = m_PopbillBase.httpGET("/FAX/SenderNumber", m_PopbillBase.getSession_token(CorpNum), "")
End Function

End Class

Class FaxState
Public sendState
Public convState
Public sendNum
Public senderName
Public receiveNum
Public receiveName
Public sendPageCnt
Public successPageCnt
Public failPageCnt
Public refundPageCnt
Public cancelPageCnt
Public reserveDT
Public sendDT
Public resultDT
Public sendResult
Public receiptDT
Public fileNames

Public Sub fromJsonInfo(jsonInfo)
	On Error Resume Next
	sendState = jsonInfo.sendState
	convState = jsonInfo.convState
	sendNum = jsonInfo.sendNum
	senderName = jsonInfo.senderName
	receiveNum = jsonInfo.receiveNum
	receiveName = jsonInfo.receiveName
	sendPageCnt = jsonInfo.sendPageCnt
	successPageCnt = jsonInfo.successPageCnt
	receiveName = jsonInfo.receiveName
	failPageCnt = jsonInfo.failPageCnt
	refundPageCnt = jsonInfo.refundPageCnt
	cancelPageCnt = jsonInfo.cancelPageCnt
	reserveDT = jsonInfo.reserveDT
	sendDT = jsonInfo.sendDT
	resultDT = jsonInfo.resultDT
	sendResult = jsonInfo.sendResult
	receiptDT = jsonInfo.receiptDT
	fileNames = jsonInfo.fileNames

	On Error GoTo 0
End Sub
End Class

Class FaxReceiver
Public receiverNum 
Public receiverName 

Public Function toJsonInfo() 
    Set toJsonInfo = JSON.parse("{}")
    toJsonInfo.set "rcv", receiverNum
    toJsonInfo.set "rcvnm", receiverName
End Function
End Class


Class FAXSearchResult
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
		For i = 0 To jsonInfo.list.length -1
			Set tmpObj = New FaxState
			tmpObj.fromJsonInfo jsonInfo.list.Get(i)
			Set list(i) = tmpObj
		Next

		On Error GoTo 0
	End Sub
End Class
%>