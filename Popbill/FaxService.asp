<%
Class FaxService

Private m_PopbillBase

'테스트 플래그
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
'회원가입 여부
Public Function CheckIsMember(CorpNum , linkID)
    Set CheckIsMember = m_PopbillBase.CheckIsMember(CorpNum,linkID)
End Function
'회원가입
Public Function JoinMember(JoinInfo)
    Set JoinMember = m_PopbillBase.JoinMember(JoinInfo)
End Function
'''''''''''''  End of PopbillBase

''단가확인
Public Function GetUnitCost(CorpNum)
    Set result = m_PopbillBase.httpGET("/FAX/UnitCost", m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'팩스 전송내역조회 URL
Public Function GetURL(CorpNum, UserID, TOGO)
    Set result = m_PopbillBase.httpGET("/FAX/?TG="+TOGO, m_PopbillBase.getSession_token(CorpNum),UserID)
	GetURL = result.url
End Function


'팩스 예약전송 취소
Public Function CancelReserve(CorpNum, ReceiptNum, UserID)
	If isNull(ReceiptNum) Or IsEmpty(ReceiptNum) Then Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다."

    Set CancelReserve = m_PopbillBase.httpGET("/FAX/"&ReceiptNum&"/Cancel", m_PopbillBase.getSession_token(CorpNum),UserID)
End Function

'팩스 전송
Public Function SendFAX(CorpNum , sendNum , receivers , FilePaths ,  reserveDT , UserID )
	If isNull(receivers) Or IsEmpty(receivers) Then Err.Raise -99999999, "POPBILL", "수신자정보 가 입력되지 않았습니다."
    If UBound(receivers) < 0 Then Err.Raise -99999999, "POPBILL", "수신자정보 가 입력되지 않았습니다."
    If isNull(FilePaths) Or IsEmpty(FilePaths) Then Err.Raise -99999999, "POPBILL", "전송할 파일경로가 입력되지 않았습니다."
    If UBound(FilePaths) < 0 Then Err.Raise -99999999, "POPBILL", "전송할 파일경로가 입력되지 않았습니다."
    If UBound(FilePaths) >= 5 Then Err.Raise -99999999, "POPBILL", "1회 전송 가능한 파일갯수는 5개입니다."
  
    Set Form = JSON.parse("{}")
    
    Form.set "snd", sendNum
    
    If reserveDT <> "" Then Form.set "sndDT", reserveDT
    
    Form.set "fCnt", UBound(FilePaths) + 1
    
	Dim tmpArray() : ReDim tmpArray(UBound(receivers))
	For i = 0 to UBound(receivers)
		If  isNull(receivers(i)) Or IsEmpty(receivers(i)) Then
			Err.Raise -99999999, "POPBILL", CStr(i+1) & " 번째 수신자 정보가 기재되지 않았습니다."
		else
			Set tmpArray(i) =  receivers(i).toJsonInfo()
		End if
	Next
    
    Form.set "rcvs", tmpArray
    
    postdata = m_PopbillBase.toString(Form)
    Set result = m_PopbillBase.httpPOST_Files("/FAX", m_PopbillBase.getSession_token(CorpNum), postdata, FilePaths, UserID)
    
    SendFAX = result.receiptNum
End Function


'팩스 전송결과 확인
Public Function GetFaxDetail(CorpNum, receiptNum, UserID)
	If  isEmpty(receiptNum) Then
			Err.Raise -99999999, "POPBILL", "팩스 접수번호(receiptNum)가 입력되지 않았습니다."
	End If

	Set sentFaxList = m_PopbillBase.httpGET("/FAX/"&receiptNum, m_PopbillBase.getSession_token(CorpNum),UserID)
	
	Set faxResult = New FaxState 

	For i=0 To sentFaxList.length-1
		faxResult.fromJsonInfo sentFaxList.Get(i)
	Next

	Set GetFaxDetail = faxResult

End Function 
End Class

Class FaxState
Public sendState
Public convState
Public sendNum
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

Public Sub fromJsonInfo(jsonInfo)
	On Error Resume Next
	sendState = jsonInfo.sendState
	convState = jsonInfo.convState
	sendNum = jsonInfo.sendNum
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
%>