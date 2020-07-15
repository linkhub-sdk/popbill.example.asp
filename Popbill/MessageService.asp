<%
Class MessageService

Private m_PopbillBase

'테스트 플래그
Public Property Let IsTest(ByVal value)
    m_PopbillBase.IsTest = value
End Property

Public Property Let IPRestrictOnOff(ByVal value)
    m_PopbillBase.IPRestrictOnOff = value
End Property

Public Property Let UseStaticIP(ByVal value)
    m_PopbillBase.UseStaticIP = value
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

'회원잔액조회
Public Function GetBalance(CorpNum)
    GetBalance = m_PopbillBase.GetBalance(CorpNum)
End Function
'파트너 잔액조회
Public Function GetPartnerBalance(CorpNum)
    GetPartnerBalance = m_PopbillBase.GetPartnerBalance(CorpNum)
End Function
'팝빌 기본 URL
Public Function GetPopbillURL(CorpNum , UserID , TOGO)
	GetPopbillURL = m_PopbillBase.GetPopbillURL(CorpNum , UserID , TOGO )
End Function
'팝빌 로그인 URL
Public Function GetAccessURL(CorpNum , UserID)
    GetAccessURL = m_PopbillBase.GetAccessURL(CorpNum , UserID )
End Function

'팝빌 연동회원 포인트 충전 URL
Public Function GetChargeURL(CorpNum , UserID)
    GetChargeURL = m_PopbillBase.GetChargeURL(CorpNum , UserID )
End Function
'파트너 포인트 충전 팝업 URL - 2017/08/29 추가
Public Function GetPartnerURL(CorpNum, TOGO)
    GetPartnerURL = m_PopbillBase.GetPartnerURL(CorpNum,TOGO)
End Function

'회원가입 여부
Public Function CheckIsMember(CorpNum , linkID)
    Set CheckIsMember = m_PopbillBase.CheckIsMember(CorpNum,linkID)
End Function
'회원가입
Public Function JoinMember(JoinInfo)
    Set JoinMember = m_PopbillBase.JoinMember(JoinInfo)
End Function
'담당자 목록조회
Public Function ListContact(CorpNum, UserID)
	Set ListContact = m_popbillBase.ListContact(CorpNum,UserID)
End Function
'담당자 정보수정
Public Function UpdateContact(CorpNum, contInfo, UserId)
	Set UpdateContact = m_popbillBase.UpdateContact(CorpNum, contInfo, UserId)
End Function
'담당자 추가 
Public Function RegistContact(CorpNum, contInfo, UserId)
	Set RegistContact = m_popbillBase.RegistContact(CorpNum, contInfo, UserId)
End Function
'회사정보 수정
Public Function UpdateCorpInfo(CorpNum, corpInfo, UserId)
	Set UpdateCorpInfo = m_popbillBase.UpdateCorpInfo(CorpNum, corpInfo, UserId)
End Function
'회사정보 확인 
Public Function GetCorpInfo(CorpNum, UserId)
	Set GetCorpInfo = m_popbillBase.GetCorpInfo(CorpNum, UserId)
End Function
Public Function CheckID(id)
	Set CheckID = m_popbillBase.CheckID(id)
End Function

'과금정보 확인
Public Function GetChargeInfo ( CorpNum, MType, UserID )
	Set result = m_PopbillBase.httpGET ( "/Message/ChargeInfo?Type=" &MType, m_PopbillBase.getSession_token(CorpNum), UserID )

	Set chrgInfo = New ChargeInfo
	chrgInfo.fromJsonInfo result
	
	Set GetChargeInfo = chrgInfo
End Function 
'''''''''''''  End of PopbillBase

''단가확인
Public Function GetUnitCost(CorpNum, MType)
    Set result = m_PopbillBase.httpGET("/Message/UnitCost?Type="&MType, m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function


'단문 메시지 전송 (requestNum)
Public Function SendSMS(CorpNum, sender, Contents, Messages, reserveDT, adsYN, requestNum, UserID)
	SendSMS = SendMessage("SMS", CorpNum, sender, "", Contents, Messages, reserveDT, adsYN, requestNum, UserID)	
End Function

'장문 메시지 전송 (requestNum)
Public Function SendLMS(CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, requestNum, UserID)
	SendLMS = SendMessage("LMS", CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, requestNum, UserID)	
End Function

'단/장문 메시지 자동인식 전송(requestNum)
Public Function SendXMS(CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, requestNum, UserID)
	SendXMS = SendMessage("XMS", CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, requestNum, UserID)	
End Function

'MMS 메시지 전송 (requestNum)
Public Function SendMMS(CorpNum, sender, subject, Contents, msgList, FilePaths, reserveDT, adsYN, requestNum, UserID)
	If IsNull(msgList) Or IsEmpty(msgList) Then 
		Err.raise -99999999, "POPBILL", "전송할 메시지가 입력되지 않았습니다."
	End If

	If isNull(FilePaths) Then Err.Raise -99999999, "POPBILL", "전송할 파일경로가 입력되지 않았습니다."

	Set tmp = JSON.parse("{}")
	    
    If sender <> "" Then tmp.Set "snd", sender
    If Contents <> "" Then tmp.Set "content", Contents
    If subject <> "" Then tmp.Set "subject", subject
    If reserveDT <> "" Then tmp.Set "sndDT", reserveDT
	If adsYN <> "" Then tmp.Set "adsYN", adsYN
	If requestNum <> "" Then tmp.Set "requestNum", requestNum

	Set msgs = JSON.parse("[]")

	For i=0 To msgList.Count-1
		Set msgObj = New Messages
		msgObj.setValue msgList.Item(i)
		msgs.Set i, msgObj.toJsonInfo
	Next

	tmp.Set "msgs", msgs

	postdata = m_PopbillBase.toString(tmp)

	Set result = m_PopbillBase.httpPost_Files("/MMS", m_PopbillBase.getSession_Token(CorpNum), postdata, FilePaths, UserID)
	SendMMS = result.receiptNum
End Function



Private Function SendMessage(MType, CorpNum, sender, subject, Contents, msgList, reserveDT, adsYN, requestNum, UserID)
	If IsNull(msgList) Or IsEmpty(msgList) Then 
		Err.raise -99999999, "POPBILL", "전송할 메시지가 입력되지 않았습니다."
	End If

	Set tmp = JSON.parse("{}")
	    
    If sender <> "" Then tmp.Set "snd", sender
    If Contents <> "" Then tmp.Set "content", Contents
    If subject <> "" Then tmp.Set "subject", subject
    If reserveDT <> "" Then tmp.Set "sndDT", reserveDT
	If requestNum <> "" Then tmp.Set "requestNum", requestNum
	If adsYN Then tmp.Set "adsYN", adsYN


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

'예약문자 전송취소
Public Function CancelReserve(CorpNum, ReceiptNum, UserID)
	If ReceiptNum = "" Or IsNull(ReceptNum) Then 
		Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다"
	End If
	
	Set CancelReserve = m_PopbillBase.httpGet("/Message/"&ReceiptNum&"/Cancel",m_PopbillBase.getSession_token(CorpNum),UserID)
End Function

'예약문자 전송취소 (요청번호 할당)
Public Function CancelReserveRN(CorpNum, RequestNum, UserID)
	If RequestNum = "" Or IsNull(RequestNum) Then 
		Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다"
	End If
	
	Set CancelReservRN = m_PopbillBase.httpGet("/Message/Cancel/"&RequestNum, m_PopbillBase.getSession_token(CorpNum), UserID)
End Function

'문자 관련 URL
Public Function GetURL(CorpNum, UserID, TOGO)
	Set result = m_PopbillBase.httpGet("/Message/?TG="+TOGO,m_PopbillBase.getSession_token(CorpNum), UserID)
	GetURL = result.url
End Function

'문자 전송내역 팝업 URL
Public Function GetSentListURL(CorpNum, UserID)
	Set result = m_PopbillBase.httpGet("/Message/?TG=BOX",m_PopbillBase.getSession_token(CorpNum), UserID)
	GetSentListURL = result.url
End Function

'발신번호 관리 팝업 URL
Public Function GetSenderNumberMgtURL(CorpNum, UserID)
	Set result = m_PopbillBase.httpGet("/Message/?TG=SENDER",m_PopbillBase.getSession_token(CorpNum), UserID)
	GetSenderNumberMgtURL = result.url
End Function



'문자 전송내역 확인
Public Function GetMessages(CorpNum, ReceiptNum, UserID)
	If ReceiptNum = "" Or IsNull(ReceptNum) Then 
		Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다"
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


'문자 전송내역 확인 (요청번호 할당)
Public Function GetMessagesRN(CorpNum, RequestNum, UserID)
	If RequestNum = "" Or IsNull(RequestNum) Then 
		Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다"
	End If
	
	Set result = m_PopbillBase.httpGet("/Message/Get/"+RequestNum ,m_PopbillBase.getSession_token(CorpNum), UserID)
	
	Set tmp = CreateObject("Scripting.Dictionary")

	For i=0 To result.length-1
		Set msgInfo = New MessageInfo
		msgInfo.fromJsonInfo result.Get(i)
		tmp.Add i, msgInfo
	Next

	Set GetMessagesRN = tmp

End Function 

'전송내역 요약정보 확인
Public Function GetStates(CorpNum, ReceiptNumList, UserID)
	If IsNull(ReceiptNumList) Then 
		Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다"
	End If

	Set tmp = JSON.parse("[]")

	For i=0 To UBound(ReceiptNumList) - 1
		tmp.Set i, ReceiptNumList(i)
	Next

	postdata = m_PopbillBase.toString(tmp)
	
	Set result = m_PopbillBase.httpPost("/Message/States", m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)
	
	Set infoObj = CreateObject("Scripting.Dictionary")

	For i=0 To result.length-1
		Set infoTmp = New MessageBriefInfo
		infoTmp.fromJsonInfo result.Get(i)
		infoObj.Add i, infoTmp
	Next

	Set GetStates = infoObj

End Function 

'문자전송내역 조회 
Public Function Search(CorpNum, SDate, EDate, Item, ReserveYN, SenderYN, Order, Page, PerPage, QString)
	If SDate = "" Then
        Err.Raise -99999999, "POPBILL", "시작일자가 입력되지 않았습니다."
	End If
	If EDate = "" Then
        Err.Raise -99999999, "POPBILL", "종료일자가 이력되지 않았습니다."
	End If

	uri = "/Message/Search"
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
	uri = uri & "&QString=" & QString


	Set searchResult = New MSGSearchResult
	Set tmpObj = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), "")

	searchResult.fromJsonInfo tmpObj
	
	Set Search = searchResult
End Function

'080 수신거부목록 확인 
Public Function GetAutoDenyList(CorpNum)
	Set GetAutoDenyList = m_PopbillBase.httpGET("/Message/Denied", m_PopbillBase.getSession_token(CorpNum), "")
End Function

' 발신번호 목록 확인
Public Function GetSenderNumberList(CorpNum)
	Set GetSenderNumberList = m_PopbillBase.httpGET("/Message/SenderNumber", m_PopbillBase.getSession_token(CorpNum), "")
End Function

End Class



Class Messages
	Public sender
	Public senderName
	Public receiver
	Public receiverName
	Public content
	Public subject

	Public Sub setValue(msgList)
		sender = msgList.sender
		senderName = msgList.senderName
		receiver = msgList.receiver
		receiverName = msgList.receiverName
		content = msgList.content
		subject = msgList.subject
	End Sub

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.set "rcv", receiver
		If sender <> "" Then  toJsonInfo.set "snd", sender
		If senderName <> "" Then  toJsonInfo.set "sndnm", senderName
		If receiverName <> "" Then toJsonInfo.set "rcvnm", receiverName
		If content <> "" Then toJsonInfo.set "msg", content
		If subject <> "" Then toJsonInfo.set "sjt", subject
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
	Public receiveNum
	Public receiveName
	Public reserveDT
	Public sendDT
	Public resultDT
	Public sendResult
	Public tranNet
	Public receiptDT
	Public requestNum
	Public receiptNum

	Public Sub fromJsonInfo(msgInfo)
		On Error Resume Next
		state = msgInfo.state
		result = msgInfo.result
		subject = msgInfo.subject
		msgType = msgInfo.type
		content = msgInfo.content
		sendNum = msgInfo.sendNum
		senderName = msgInfo.senderName
		receiveNum = msgInfo.receiveNum
		receiveName = msgInfo.receiveName
		reserveDT = msgInfo.reserveDT
		sendDT = msgInfo.sendDT
		resultDT = msgInfo.resultDT
		sendResult = msgInfo.sendResult
		tranNet = msgInfo.tranNet
		receiptDT = msgInfo.receiptDT
		requestNum = msgInfo.requestNum
		receiptNum = msgInfo.receiptNum
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
		For i = 0 To jsonInfo.list.length -1
			Set tmpObj = New MessageInfo
			tmpObj.fromJsonInfo jsonInfo.list.Get(i)
			Set list(i) = tmpObj
		Next

		On Error GoTo 0
	End Sub
End Class

%>