<%
Class CashbillService

Private m_PopbillBase

'테스트 플래그
Public Property Let IsTest(ByVal value)
    m_PopbillBase.IsTest = value
End Property

Public Sub Class_Initialize
	Set m_PopbillBase = New PopbillBase
	m_PopbillBase.AddScope("140")
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
Public Function GetChargeInfo ( CorpNum, UserID )
	Set result = m_PopbillBase.httpGET("/Cashbill/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)

	Set chrgInfo = New ChargeInfo
	chrgInfo.fromJsonInfo result
	
	Set GetChargeInfo = chrgInfo
End Function
'''''''''''''  End of PopbillBase

'단가확인
Public Function GetUnitCost(CorpNum)
    Set result = m_PopbillBase.httpGET("/Cashbill?cfg=UNITCOST", m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'연동관리번호 사용여부 확인
Public Function CheckMgtKeyInUse(CorpNum, mgtKey) 
	If mgtKey = "" Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	On Error Resume Next
	
	Set CheckMgtKeyInUse = m_PopbillBase.httpGet("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum),"")
	
	If Err.Number = -14000003 Then
		CheckMgtKeyInUse = False
	Else
		CheckMgtKeyInUse = True
	End If
	On Error Resume Next
End Function 


'팝빌 SSO URL확인
Public Function GetURL(CorpNum, UserID, TOGO)
	Set result = m_PopbillBase.httpGet("/Cashbill?TG=" + TOGO, m_PopbillBase.getSession_token(CorpNum),UserID)
	GetURL = result.url
End Function 


'현금영수증 보기 URL
Public Function GetPopUpURL(CorpNum, mgtKey, UserID)
	If mgtKey = "" Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=POPUP", m_PopbillBase.getSession_token(CorpNum), UserID)
	GetPopUpURL = result.url
End Function 


'현금영수증 인쇄 URL
Public Function GetPrintURL(CorpNum, mgtKey, UserID)
	If mgtKey = "" Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=PRINT", m_PopbillBase.getSession_token(CorpNum), UserID)
	GetPrintURL = result.url
End Function 


'현금영수증 인쇄 URL - 공급받는자
Public Function GetEPrintURL(CorpNum, mgtKey, UserID)
	If mgtKey = "" Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=EPRINT", m_PopbillBase.getSession_token(CorpNum), UserID)
	GetEPrintURL = result.url
End Function 


'현금영수증 이메일 링크 URL
Public Function GetMailURL(CorpNum, mgtKey, UserID)
	If mgtKey = "" Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=MAIL", m_PopbillBase.getSession_token(CorpNum), UserID)
	GetMailURL = result.url
End Function 


'다량 현금영수증 인쇄 URL 
Public Function GetMassPrintURL(CorpNum, mgtKeyList, UserID)
	If isEmpty(mgtKeyList) Or isNull(mgtKeyList) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set tmp = JSON.parse("[]")
	For i=0 To UBound(mgtKeyList)-1
		tmp.Set i, mgtKeyList(i)
	Next

	postdata = m_PopbillBase.toString(tmp)

	Set result = m_PopbillBase.httpPOST("/Cashbill/Prints", m_PopbillBase.getSession_token(CorpNum), "",postdata, UserID)

	GetMassPrintURL = result.url
End Function 


'현금영수증 임시저장
Public Function Register(CorpNum, ByRef Cashbill, UserID)
	Set tmpDic = Cashbill.toJsonInfo()
	postdata = m_PopbillBase.toString(tmpDic)

	Set Register = m_PopbillBase.httpPOST("/Cashbill", m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)
End Function 


'현금영수증 수정
Public Function Update(CorpNum, mgtKey, ByRef Cashbill, UserID)
	Set tmpDic = Cashbill.toJsonInfo()
	postdata = m_PopbillBase.toString(tmpDic)

	Set Update = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "PATCH", postdata, UserID)
End Function 


'현금영수증 발행
Public Function Issue(CorpNum, mgtKey, Memo, UserID)
	If mgtKey = "" Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set tmp = JSON.parse("{}")
	tmp.Set "memo", Memo

	postdata = m_PopbillBase.toString(tmp)
	Set Issue = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "ISSUE", postdata, UserID)
End Function 


'현금영수증 발행취소
Public Function CancelIssue(CorpNum, mgtKey, Memo, UserID)
	If mgtKey = "" Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set tmp = JSON.parse("{}")
	tmp.Set "memo", Memo

	postdata = m_PopbillBase.toString(tmp)
	Set CancelIssue = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "CANCELISSUE", postdata, UserID)
End Function 


'현금영수증 삭제
Public Function Delete(CorpNum, mgtKey, UserID)
	If mgtKey = "" Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set Delete = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "DELETE", "", UserID)
End Function 


'현금영수증 상태정보 조회
Public Function GetInfo(CorpNum, mgtKey, UserID)
	If mgtKey = "" Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set result = m_PopbillBase.httpGet("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), UserID)
	
	Set infoObj = New CashbillInfo	
	infoObj.fromJsonInfo result
	Set GetInfo = infoObj
End Function 


'다량 현금영수증 상태정보 조회
Public Function GetInfos(CorpNum, mgtKeyList, UserID)
	If isNull(mgtKeyList) Or isEmpty(mgtKeyList) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If
	
	Set tmp = JSON.parse("[]")

	For i=0 To UBound(mgtKeyList)-1
		tmp.Set i, mgtKeyList(i)
	Next

	postdata = m_PopbillBase.toString(tmp)
	
	Set result = m_PopbillBase.httpPOST("/Cashbill/States", m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)

	Set tmpDic = CreateObject("Scripting.Dictionary")

	For i=0 To result.length-1
		Set cbInfo = New CashbillInfo 
		cbInfo.fromJsonInfo result.Get(i)
		tmpDic.Add i, cbInfo
	Next

	Set GetInfos = tmpDic
End Function 


'현금영수증 이력확인
Public Function GetLogs(CorpNum, mgtKey, UserID)
	If isNull(mgtKey) Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If
	
	Set result = m_PopbillBase.httpGet("/Cashbill/"+mgtKey+"/Logs", m_PopbillBase.getSession_token(CorpNum),UserID)
	
	Set tmp = CreateObject("Scripting.Dictionary")

	For i=0 To result.length-1
		Set chLog = New CashbillLog
		chLog.fromJsonInfo result.Get(i)
		tmp.Add i, chLog
	Next 

	Set GetLogs = tmp 
End Function 


'상세정보 확인
Public Function GetDetailInfo(CorpNum, mgtKey, UserID)
	If isNull(mgtKey) Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set result = m_PopbillBase.httpGet("/Cashbill/"+mgtKey+"?Detail", m_PopbillBase.getSession_token(CorpNum),UserID)

	Set tmp = New Cashbill

	tmp.fromJsonInfo result
	
	Set GetDetailInfo = tmp
End Function 


'알림메일 재전송
Public Function SendEmail(CorpNum, mgtKey, Receiver, UsrID)
	If isNull(mgtKey) Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set tmp = JSON.parse("{}")
	tmp.Set "receiver", Receiver

	postdata = m_PopbillBase.toString(tmp)

	Set SendEmail = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey , m_PopbillBase.getSession_token(CorpNum), "EMAIL", postdata, UserID)
End Function 


'알림문자 전송
Public Function SendSMS(CorpNum, mgtKey, Sender, Receiver, Contents, UserID)
	If isNull(mgtKey) Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set tmp = JSON.parse("{}")
	tmp.Set "receiver", Receiver
	tmp.Set "sender", Sender
	tmp.Set "contents", Contents
	
	postdata = m_PopbillBase.toString(tmp)

	Set SendSMS = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey , m_PopbillBase.getSession_token(CorpNum), "SMS", postdata, UserID)
End Function 


'팩스 전송
Public Function SendFAX(CorpNum, mgtKey, Sender, Receiver, UserID)
	If isNull(mgtKey) Or isEmpty(mgtKey) Then 
		Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
	End If

	Set tmp = JSON.parse("{}")
	tmp.Set "receiver", Receiver
	tmp.Set "sender", Sender

	postdata = m_PopbillBase.toString(tmp)

	Set SendFAX = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey , m_PopbillBase.getSession_token(CorpNum), "FAX", postdata, UserID)
End Function 

'현금영수증 즉시발행
Public Function RegistIssue(CorpNum, ByRef Cashbill, Memo, UserID)
	Set tmpDic = Cashbill.toJsonInfo
	tmpDic.Set "memo", Memo

	postdata = m_PopbillBase.toString(tmpDic)
	
	Set RegistIssue = m_PopbillBase.httpPOST("/Cashbill", m_PopbillBase.getSession_token(CorpNum), "ISSUE", postdata, UserID)
End Function 

'현금영수증 목록 조회
Public Function Search(CorpNum, DType, SDate, EDate, State, TradeType, TradeUsage, TaxationType, Order, Page, PerPage)
    If DType = "" Then
        Err.Raise -99999999, "POPBILL", "검색일자 유형이 입력되지 않았습니다."
	End If
	If SDate = "" Then
        Err.Raise -99999999, "POPBILL", "시작일자가 입력되지 않았습니다."
	End If
	If EDate = "" Then
        Err.Raise -99999999, "POPBILL", "종료일자가 이력되지 않았습니다."
	End If

	uri = "/Cashbill/Search"
	uri = uri & "?DType=" & DType
	uri = uri & "&SDate=" & SDate
	uri = uri & "&EDate=" & EDate

	uri = uri & "&State="
	For i=0 To UBound(State) -1	
		If i = UBound(State) -1 then
			uri = uri & State(i)
		Else
			uri = uri & State(i) & ","
		End If
	Next
	
	uri = uri & "&TradeType="
	For i=0 To UBound(TradeType) -1	
		If i = UBound(TradeType) -1 then
			uri = uri & TradeType(i)
		Else
			uri = uri & TradeType(i) & ","
		End If
	Next

	uri = uri & "&TradeUsage="
	For i=0 To UBound(TradeUsage) -1	
		If i = UBound(TradeUsage) -1 then
			uri = uri & TradeUsage(i)
		Else
			uri = uri & TradeUsage(i) & ","
		End If
	Next

	uri = uri & "&TaxationType="
	For i=0 To UBound(TaxationType) -1	
		If i = UBound(TaxationType) -1 then
			uri = uri & TaxationType(i)
		Else
			uri = uri & TaxationType(i) & ","
		End If
	Next

	uri = uri & "&Order=" & Order
	uri = uri & "&Page=" & CStr(Page)
	uri = uri & "&PerPage=" & CStr(PerPage)
	
	Set searchResult = New CBSearchResult
	Set tmpObj = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), "")

	searchResult.fromJsonInfo tmpObj
	
	Set Search = searchResult
End Function

End Class

Class Cashbill
Public mgtKey
Public tradeDate
Public tradeUsage
Public tradeType
Public taxationType
Public supplyCost
Public tax
Public serviceFee
Public totalAmount

Public franchiseCorpNum
Public franchiseCorpName
Public franchiseCEOName
Public franchiseAddr
Public franchiseTEL

Public identityNum
Public customerName
Public itemName
Public orderNumber

Public email
Public hp
Public fax
Public smssendYN
Public faxsendYN

Public confirmNum

Public orgConfirmNum
Public orgTradeDate

Public Sub fromJsonInfo(jsonInfo)
	On Error Resume Next
	mgtKey = jsonInfo.mgtKey
	tradeDate = jsonInfo.tradeDate
	tradeUsage = jsonInfo.tradeUsage
	tradeType = jsonInfo.tradeType
	taxationType = jsonInfo.taxationType
	supplyCost = jsonInfo.supplyCost
	tax = jsonInfo.tax
	serviceFee = jsonInfo.serviceFee
	totalAmount = jsonInfo.totalAmount

	franchiseCorpNum = jsonInfo.franchiseCorpNum
	franchiseCorpName = jsonInfo.franchiseCorpName
	franchiseCEOName = jsonInfo.franchiseCEOName
	franchiseAddr = jsonInfo.franchiseAddr
	franchiseTEL = jsonInfo.franchiseTEL

	identityNum = jsonInfo.identityNum
	customerName = jsonInfo.customerName
	itemName = jsonInfo.itemName
	orderNumber = jsonInfo.orderNumber

	email = jsonInfo.email
	hp = jsonInfo.hp
	fax = jsonInfo.fax
	smssendYN = jsonInfo.smssendYN
	faxsendYN = jsonInfo.faxsendYN

	confirmNum = jsonInfo.confirmNum

	orgConfirmNum = jsonInfo.orgConfirmNum
	orgTradeDate = jsonInfo.orgTradeDate

	On Error GoTo 0 
End Sub 

Public Function toJsonInfo()
	Set toJsonInfo = JSON.parse("{}")
	toJsonInfo.Set "mgtKey", mgtKey
	toJsonInfo.Set "tradeDate", tradeDate
	toJsonInfo.Set "tradeUsage", tradeUsage
	toJsonInfo.Set "tradeType", tradeType
	toJsonInfo.Set "taxationType", taxationType
	toJsonInfo.Set "supplyCost", supplyCost
	toJsonInfo.Set "tax", tax
	toJsonInfo.Set "serviceFee", serviceFee
	toJsonInfo.Set "totalAmount", totalAmount

	toJsonInfo.Set "franchiseCorpNum", franchiseCorpNum
	toJsonInfo.Set "franchiseCorpName", franchiseCorpName
	toJsonInfo.Set "franchiseCEOName", franchiseCEOName
	toJsonInfo.Set "franchiseAddr", franchiseAddr
	toJsonInfo.Set "franchiseTEL", franchiseTEL

	toJsonInfo.Set "identityNum", identityNum
	toJsonInfo.Set "customerName", customerName
	toJsonInfo.Set "itemName", itemName
	toJsonInfo.Set "orderNumber", orderNumber

	toJsonInfo.Set "email", email
	toJsonInfo.Set "hp", hp
	toJsonInfo.Set "fax", fax
	toJsonInfo.Set "smssendYN", smssendYN
	toJsonInfo.Set "faxsendYN", faxsendYN

	toJsonInfo.Set "confirmNum", confirmNum

	toJsonInfo.Set "orgConfirmNum", orgConfirmNum
	toJsonInfo.Set "orgTradeDate", orgTradeDate
End Function 
End Class


Class CashbillLog
Public docLogType
Public log
Public procType
Public procMemo
Public procCorpName
Public regDT
Public ip

Public Sub fromJsonInfo(jsonInfo)
	On Error Resume Next
	docLogType = jsonInfo.docLogType
	log = jsonInfo.log
	procType = jsonInfo.procType
	procMemo = jsonInfo.procMemo
	procCorpName = jsonInfo.procCorpName
	regDT = jsonInfo.regDT
	ip = jsonInfo.ip
	On Error GoTo 0 
End Sub 
End Class

Class CashbillInfo
Public itemKey 
Public mgtKey 
Public tradeDate 
Public issueDT 
Public customerName 
Public itemName 
Public identityNum 
Public taxationType 

Public totalAmount 
Public tradeUsage 
Public tradeType 
Public stateCode 
Public stateDT 
Public printYN 

Public confirmNum 
Public orgTradeDate 
Public orgConfirmNum 

Public ntssendDT 
Public ntsresult 
Public ntsresultDT 
Public ntsresultCode 
Public ntsresultMessage 

Public regDT 

Public Sub fromJsonInfo(jsonInfo)
	On Error Resume Next
	itemKey = jsonInfo.itemKey 
	mgtKey = jsonInfo.mgtKey 
	tradeDate = jsonInfo.mgtKey
	issueDT = jsonInfo.issueDT
	customerName = jsonInfo.customerName 
	itemName = jsonInfo.itemName 
	identityNum = jsonInfo.identityNum 
	taxationType = jsonInfo.taxationType 

	totalAmount = jsonInfo.totalAmount 
	tradeUsage = jsonInfo.tradeUsage 
	tradeType = jsonInfo.tradeType 
	stateCode = jsonInfo.stateCode
	stateDT = jsonInfo.stateDT 
	printYN = jsonInfo.printYN 

	confirmNum = jsonInfo.confirmNum 
	orgTradeDate = jsonInfo.orgTradeDate 
	orgConfirmNum = jsonInfo.orgTradeDate 

	ntssendDT = jsonInfo.ntssendDT 
	ntsresult = jsonInfo.ntsresult 
	ntsresultDT = jsonInfo.ntsresultDT 
	ntsresultCode = jsonInfo.ntsresultCode 
	ntsresultMessage = jsonInfo.ntsresultMessage 

	regDT = jsonInfo.regDT
	On Error GoTo 0
End Sub
End Class

Class CBSearchResult
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
			Set tmpObj = New CashbillInfo
			tmpObj.fromJsonInfo jsonInfo.list.Get(i)
			Set list(i) = tmpObj
		Next

		On Error GoTo 0
	End Sub
End Class

%>