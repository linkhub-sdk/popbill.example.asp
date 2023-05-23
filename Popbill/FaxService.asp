<%
Class FaxService

Private m_PopbillBase

'테스트 플래그
Public Property Let IsTest(ByVal value)
	m_PopbillBase.IsTest = value
End Property

Public Property Let IPRestrictOnOff(ByVal value)
	m_PopbillBase.IPRestrictOnOff = value
End Property

Public Property Let UseLocalTimeYN(ByVal value)
	m_PopbillBase.UseLocalTimeYN = value
End Property

Public Sub Class_Initialize
	Set m_PopbillBase = New PopbillBase
	m_PopbillBase.AddScope("160")
	m_PopbillBase.AddScope("161")
End Sub

Public Sub Initialize(linkID, SecretKey )
	m_PopbillBase.Initialize linkID,SecretKey
End Sub

Public Property Let UseStaticIP(ByVal value)
	m_PopbillBase.UseStaticIP = value
End Property

Public Property Let UseGAIP(ByVal value)
	m_PopbillBase.UseGAIP = value
End Property

'회원잔액조회
Public Function GetBalance(CorpNum)
	GetBalance = m_PopbillBase.GetBalance(CorpNum)
End Function
'파트너 잔액조회
Public Function GetPartnerBalance(CorpNum)
	GetPartnerBalance = m_PopbillBase.GetPartnerBalance(CorpNum)
End Function
'팝빌 기본 URL
Public Function GetPopbillURL(CorpNum , UserID , TOGO )
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

'팝빌 연동회원 포인트 결제내역 URL
Public Function GetPaymentURL(CorpNum, UserID)
	GetPaymentURL = m_PopbillBase.GetPaymentURL(CorpNum, UserID)
End Function

'팝빌 연동회원 포인트 사용내역 URL
Public Function GetUseHistoryURL(CorpNum, UserID)
	GetUseHistoryURL = m_PopbillBase.GetUseHistoryURL(CorpNum, UserID)
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
'담당자 정보 확인
Public Function GetContactInfo(CorpNum, ContactID, UserID)
	Set GetContactInfo = m_PopbillBase.GetContactInfo(CorpNum, ContactID, UserID)
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
Public Function GetChargeInfo ( CorpNum, UserID, ReceiveNumType )
	Dim result : Set result = m_PopbillBase.httpGET("/FAX/ChargeInfo?receiveNumType=" & ReceiveNumType, m_PopbillBase.getSession_token(CorpNum), UserID)

	Dim chrgInfo : Set chrgInfo = New ChargeInfo
	chrgInfo.fromJsonInfo result

	Set GetChargeInfo = chrgInfo
End Function

'무통장 입금신청
Public Function PaymentRequest(CorpNum, PaymentForm, UserID)
	Set PaymentRequest = m_popbillBase.PaymentRequest(CorpNum, PaymentForm, UserID)
End Function

'연동회원 포인트 결제내역 조회
Public Function GetPaymentHistory(CorpNum, SDate, EDate, Page, PerPage, UserID)
	Set GetPaymentHistory = m_popbillBase.GetPaymentHistory(CorpNum, SDate, EDate, Page, PerPage, UserID)
End Function

'연동회원 무통장 입금신청 정보확인
Public Function GetSettleResult(CorpNum, SettleCode, UserID)
	Set GetSettleResult = m_popbillBase.GetSettleResult(CorpNum, SettleCode, UserID)
End Function

'연동회원 포인트 사용내역 확인
Public Function GetUseHistory(CorpNum, SDate, EDate, Page, PerPage, Order, UserID)
	Set GetUseHistory = m_PopbillBase.GetUseHistory(CorpNum, SDate, EDate, Page, PerPage, Order, UserID)
End Function

'연동회원 포인트 환불신청
Public Function Refund(CorpNum, RefundForm, UserID)
	Set Refund = m_popbillBase.Refund(CorpNum, RefundForm, UserID)
End Function

' 환불 가능 포인트 조회
Public Function GetRefundableBalance(CorpNum, UserID)
	GetRefundableBalance = m_popbillBase.GetRefundableBalance(CorpNum, UserID)
End Function

'연동회원 포인트 환불내역 확인
Public Function GetRefundHistory(CorpNum, Page, PerPage, UserID)
	Set GetRefundHistory = m_popbillBase.GetRefundHistory(CorpNum, Page, PerPage, UserID)
End Function

' 환불 신청 상태 조회
Public Function GetRefundInfo(CorpNum, RefundCode, UserID)
	Set GetRefundInfo = m_popbillBase.GetRefundInfo(CorpNum, RefundCode, UserID)
End Function

'회원 탈퇴
Public Function QuitMember(CorpNum, QuitReason, UserID)
	Set QuitMember = m_popbillBase.QuitMember(CorpNum, QuitReason, UserID)
End Function

'''''''''''''  End of PopbillBase

''단가확인
Public Function GetUnitCost(CorpNum, ReceiveNumType)
	Dim result : Set result = m_PopbillBase.httpGET("/FAX/UnitCost?receiveNumType=" & ReceiveNumType, m_PopbillBase.getSession_token(CorpNum),"")
	GetUnitCost = result.unitCost
End Function

'팩스 전송내역조회 URL
Public Function GetURL(CorpNum, UserID, TOGO)
	Dim result : Set result = m_PopbillBase.httpGET("/FAX/?TG="+TOGO, m_PopbillBase.getSession_token(CorpNum),UserID)
	GetURL = result.url
End Function


'팩스 전송내역 팝업 URL
Public Function GetSentListURL(CorpNum, UserID)
	Dim result : Set result = m_PopbillBase.httpGet("/FAX/?TG=BOX",m_PopbillBase.getSession_token(CorpNum), UserID)
	GetSentListURL = result.url
End Function

'발신번호 관리 팝업 URL
Public Function GetSenderNumberMgtURL(CorpNum, UserID)
	Dim result : Set result = m_PopbillBase.httpGet("/FAX/?TG=SENDER",m_PopbillBase.getSession_token(CorpNum), UserID)
	GetSenderNumberMgtURL = result.url
End Function

'팩스 미리보기 URL
Public Function GetPreviewURL(CorpNum, ReceiptNum, UserID)
	If Len(ReceiptNum ) <> 18 Or IsNull(ReceiptNum) Then
    	Err.Raise -99999999, "POPBILL", "접수번호가 올바르지 않습니다"
	End If

	Dim result : Set result = m_PopbillBase.httpGET("/FAX/Preview/"+ReceiptNum, m_PopbillBase.getSession_token(CorpNum),UserID)
	GetPreviewURL = result.url
End Function

'팩스 예약전송 취소
Public Function CancelReserve(CorpNum, ReceiptNum, UserID)
	If isNull(ReceiptNum) Or IsEmpty(ReceiptNum) Then Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다."

	Set CancelReserve = m_PopbillBase.httpGET("/FAX/"&ReceiptNum&"/Cancel", m_PopbillBase.getSession_token(CorpNum),UserID)
End Function

'예약 전송취소 (요청번호 할당)
Public Function CancelReserveRN(CorpNum, RequestNum, UserID)
	If RequestNum = "" Or IsNull(RequestNum) Then
    	Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다"
	End If

	Set CancelReserveRN = m_PopbillBase.httpGet("/FAX/Cancel/"&RequestNum, m_PopbillBase.getSession_token(CorpNum), UserID)
End Function

'팩스 전송
Public Function SendFAX(CorpNum , sendNum , receivers , FilePaths ,  reserveDT , UserID, adsYN, title, requestNum)
	If isNull(receivers) Or IsEmpty(receivers) Then Err.Raise -99999999, "POPBILL", "수신자정보 가 입력되지 않았습니다."
	If UBound(receivers) < 0 Then Err.Raise -99999999, "POPBILL", "수신자정보 가 입력되지 않았습니다."
	If isNull(FilePaths) Or IsEmpty(FilePaths) Then Err.Raise -99999999, "POPBILL", "전송할 파일경로가 입력되지 않았습니다."
	If UBound(FilePaths) < 0 Then Err.Raise -99999999, "POPBILL", "전송할 파일경로가 입력되지 않았습니다."

	Dim Form : Set Form = JSON.parse("{}")

	Form.set "snd", sendNum
	If reserveDT <> "" Then Form.set "sndDT", reserveDT
  	If adsYN Then Form.Set "adsYN", adsYN
	If title <> "" Then Form.set "title", title
	If requestNum <> "" Then Form.set "requestNum", requestNum

	Form.set "fCnt", UBound(FilePaths) + 1

	Dim tmpArray() : ReDim tmpArray(UBound(receivers))
	Dim i
	For i = 0 to UBound(receivers)
    	If  isNull(receivers(i)) Or IsEmpty(receivers(i)) Then
        	Err.Raise -99999999, "POPBILL", CStr(i+1) & " 번째 수신자 정보가 기재되지 않았습니다."
    	else
        	Set tmpArray(i) =  receivers(i).toJsonInfo()
    	End if
	Next

	Form.set "rcvs", tmpArray

	Dim postData : postData = m_PopbillBase.toString(Form)
	Dim result : Set result = m_PopbillBase.httpPOST_Files("/FAX", m_PopbillBase.getSession_token(CorpNum), postData, FilePaths, UserID)

	SendFAX = result.receiptNum
End Function


'팩스 재전송
Public Function ResendFAX(CorpNum, receiptNum, sendNum, senderName, receivers,  reserveDT , UserID, title, requestNum)
	If isNull(receiptNum) Or IsEmpty(receiptNum) Then Err.Raise -99999999, "POPBILL", "팩스 접수번호(receiptNum)가 입력되지 않았습니다."

	Dim Form : Set Form = JSON.parse("{}")

	If sendNum <> "" Then Form.set "snd", sendNum
	If senderName <> "" Then Form.set "sndnm", senderName
	If reserveDT <> "" Then Form.set "sndDT", reserveDT
	If requestNum <> "" Then Form.set "requestNum", requestNum

	If title <> "" Then Form.set "title", title

	If UBound(receivers) >= 0 Then
    	Dim tmpArray() : ReDim tmpArray(UBound(receivers))
    	Dim i
    	For i = 0 to UBound(receivers)
        	If  isNull(receivers(i)) Or IsEmpty(receivers(i)) Then
            	Err.Raise -99999999, "POPBILL", CStr(i+1) & " 번째 수신자 정보가 기재되지 않았습니다."
        	else
            	Set tmpArray(i) =  receivers(i).toJsonInfo()
        	End if
    	Next
    	Form.set "rcvs", tmpArray
	End If

	Dim postData : postData = m_PopbillBase.toString(Form)
	Dim result : Set result = m_PopbillBase.httpPOST("/FAX/"&receiptNum, m_PopbillBase.getSession_token(CorpNum), "", postData, UserID)

	ResendFAX = result.receiptNum
End Function


'팩스 재전송 (요청번호 할당)
Public Function ResendFAXRN(CorpNum, orgRequestNum, sendNum, senderName, receivers,  reserveDT , UserID, title, requestNum)
	If isNull(orgRequestNum) Or IsEmpty(orgRequestNum) Then Err.Raise -99999999, "POPBILL", "원본 팩스 요청번호가 입력되지 않았습니다."
	Dim Form : Set Form = JSON.parse("{}")

	If sendNum <> "" Then Form.set "snd", sendNum
	If senderName <> "" Then Form.set "sndnm", senderName
	If reserveDT <> "" Then Form.set "sndDT", reserveDT
	If requestNum <> "" Then Form.set "requestNum", requestNum
	If title <> "" Then Form.set "title", title

	If UBound(receivers) >= 0 Then
    	Dim tmpArray() : ReDim tmpArray(UBound(receivers))
    	Dim i
    	For i = 0 to UBound(receivers)
        	If  isNull(receivers(i)) Or IsEmpty(receivers(i)) Then
            	Err.Raise -99999999, "POPBILL", CStr(i+1) & " 번째 수신자 정보가 기재되지 않았습니다."
        	else
            	Set tmpArray(i) =  receivers(i).toJsonInfo()
        	End if
    	Next
    	Form.set "rcvs", tmpArray
	End If

	Dim postData : postData = m_PopbillBase.toString(Form)
	Dim result : Set result = m_PopbillBase.httpPOST("/FAX/Resend/"&orgRequestNum, m_PopbillBase.getSession_token(CorpNum), "", postData, UserID)

	ResendFAXRN = result.receiptNum
End Function


'팩스 전송결과 확인
Public Function GetFaxDetail(CorpNum, receiptNum, UserID)
	If  isEmpty(receiptNum) Then
        	Err.Raise -99999999, "POPBILL", "팩스 접수번호(receiptNum)가 입력되지 않았습니다."
	End If

	Dim result : Set result = m_PopbillBase.httpGET("/FAX/"&receiptNum, m_PopbillBase.getSession_token(CorpNum),UserID)

	Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")

	Dim i
	For i=0 To result.length-1
    	Dim faxInfo : Set faxInfo = New FaxState
    	faxInfo.fromJsonInfo result.Get(i)
    	tmp.Add i, faxInfo
	Next

	Set GetFaxDetail = tmp

End Function


'팩스 전송내역 확인 (요청번호 할당)
Public Function GetFaxDetailRN(CorpNum, RequestNum, UserID)
	If RequestNum = "" Or IsNull(RequestNum) Then
    	Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다"
	End If

	Dim result : Set result = m_PopbillBase.httpGet("/FAX/Get/"+RequestNum ,m_PopbillBase.getSession_token(CorpNum), UserID)

	Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")

	Dim i
	For i=0 To result.length-1
    	Dim faxInfo : Set faxInfo = New FaxState
    	faxInfo.fromJsonInfo result.Get(i)
    	tmp.Add i, faxInfo
	Next

	Set GetFaxDetailRN = tmp

End Function


'팩스 목록 조회
Public Function Search(CorpNum, SDate, EDate, State, ReserveYN, SenderOnlyYN, Order, Page, PerPage, QString)
	If SDate = "" Then
    	Err.Raise -99999999, "POPBILL", "시작일자가 입력되지 않았습니다."
	End If
	If EDate = "" Then
    	Err.Raise -99999999, "POPBILL", "종료일자가 이력되지 않았습니다."
	End If

	Dim uri
	uri = "/FAX/Search"
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

	If ReserveYN Then uri = uri & "&ReserveYN=1"
	If SedernOnlyYN Then uri = uri & "&SenderOnlyYN=1"

	uri = uri & "&Order=" & Order
	uri = uri & "&Page=" & CStr(Page)
	uri = uri & "&PerPage=" & CStr(PerPage)
	uri = uri & "&QString=" & Server.URLEncode(QString)

	Dim searchResult : Set searchResult = New FAXSearchResult
	Dim tmpObj : Set tmpObj = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), "")

	searchResult.fromJsonInfo tmpObj

	Set Search = searchResult
End Function

' 발신번호 목록 확인
Public Function GetSenderNumberList(CorpNum)
	Set GetSenderNumberList = m_PopbillBase.httpGET("/FAX/SenderNumber", m_PopbillBase.getSession_token(CorpNum), "")
End Function

' 발신번호 등록여부 확인
Public Function CheckSenderNumber(CorpNum, SenderNumber, UserID)
	If SenderNumber = "" Or IsNull(SenderNumber) Then
    	Err.Raise -99999999, "POPBILL", "발신번호가 입력되지 않았습니다."
	End If

	Set CheckSenderNumber = m_PopbillBase.httpGET("/FAX/CheckSenderNumber/"&SenderNumber,m_PopbillBase.getSession_token(CorpNum),UserID)
End Function

End Class

Class FaxState

Public state
Public result
Public title
Public sendState
Public convState
Public sendNum
Public senderName
Public receiveNum
Public receiveNumType
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
Public requestNum
Public receiptNum
Public interOPRefKey
Public chargePageCnt
Public tiffFileSize

Public Sub fromJsonInfo(jsonInfo)
	On Error Resume Next
	state = jsonInfo.state
	result = jsonInfo.result
	title = jsonInfo.title

	sendState = jsonInfo.sendState
	convState = jsonInfo.convState
	sendNum = jsonInfo.sendNum
	senderName = jsonInfo.senderName
	receiveNum = jsonInfo.receiveNum
	receiveNum = jsonInfo.receiveNumType
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
	requestNum = jsonInfo.requestNum
	receiptNum = jsonInfo.receiptNum
	interOPRefKey = jsonInfo.interOPRefKey
	chargePageCnt = jsonInfo.chargePageCnt
	tiffFileSize = jsonInfo.tiffFileSize

	On Error GoTo 0
End Sub
End Class

Class FaxReceiver
Public receiverNum
Public receiverName
Public interOPRefKey

Public Function toJsonInfo()
	Set toJsonInfo = JSON.parse("{}")
	toJsonInfo.set "rcv", receiverNum
	toJsonInfo.set "rcvnm", receiverName
	toJsonInfo.set "interOPRefKey", interOPRefKey
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
    	Dim i
    	For i = 0 To jsonInfo.list.length -1
        	Dim tmpObj : Set tmpObj = New FaxState
        	tmpObj.fromJsonInfo jsonInfo.list.Get(i)
        	Set list(i) = tmpObj
    	Next

    	On Error GoTo 0
	End Sub
End Class
%>