<%
Class CashbillService

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

Public Property Let UseLocalTimeYN(ByVal value)
    m_PopbillBase.UseLocalTimeYN = value
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
Public Function GetChargeInfo ( CorpNum, UserID )
    Dim result : Set result = m_PopbillBase.httpGET("/Cashbill/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)

    Dim chrgInfo : Set chrgInfo = New ChargeInfo
    chrgInfo.fromJsonInfo result
    
    Set GetChargeInfo = chrgInfo
End Function
'''''''''''''  End of PopbillBase

'단가확인
Public Function GetUnitCost(CorpNum)
    Dim result : Set result = m_PopbillBase.httpGET("/Cashbill?cfg=UNITCOST", m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'연동문서번호 사용여부 확인
Public Function CheckMgtKeyInUse(CorpNum, mgtKey) 
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    On Error Resume Next
    
    Dim result : Set result = m_PopbillBase.httpGet("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum),"")

    If Err.Number = -14000003 Then
        CheckMgtKeyInUse = False
    Else
        CheckMgtKeyInUse = True
    End If
    On Error Resume Next
End Function 


'팝빌 SSO URL확인
Public Function GetURL(CorpNum, UserID, TOGO)
    Dim result : Set result = m_PopbillBase.httpGet("/Cashbill?TG=" + TOGO, m_PopbillBase.getSession_token(CorpNum),UserID)
    GetURL = result.url
End Function 


'현금영수증 보기 URL
Public Function GetPopUpURL(CorpNum, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=POPUP", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetPopUpURL = result.url
End Function 

'현금영수증 보기 URL (메뉴/버튼 제외)
Public Function GetViewURL(CorpNum, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=VIEW", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetViewURL = result.url
End Function 

'현금영수증 인쇄 URL
Public Function GetPDFURL(CorpNum, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=PDF", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetPDFURL = result.url
End Function 

'현금영수증 인쇄 URL
Public Function GetPrintURL(CorpNum, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=PRINT", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetPrintURL = result.url
End Function 


'현금영수증 인쇄 URL - 공급받는자
Public Function GetEPrintURL(CorpNum, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=EPRINT", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetEPrintURL = result.url
End Function 


'현금영수증 이메일 링크 URL
Public Function GetMailURL(CorpNum, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Cashbill/"+mgtKey +"?TG=MAIL", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetMailURL = result.url
End Function 


'다량 현금영수증 인쇄 URL 
Public Function GetMassPrintURL(CorpNum, mgtKeyList, UserID)
    If isEmpty(mgtKeyList) Or isNull(mgtKeyList) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim tmp : Set tmp = JSON.parse("[]")
    Dim i
    For i=0 To UBound(mgtKeyList)-1
        tmp.Set i, mgtKeyList(i)
    Next

    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Dim result : Set result = m_PopbillBase.httpPOST("/Cashbill/Prints", m_PopbillBase.getSession_token(CorpNum), "",postdata, UserID)

    GetMassPrintURL = result.url
End Function 

Public Function AssignMgtKey(CorpNum, ItemKey, MgtKey)
    If ItemKey = "" Or isEmpty(ItemKey) Then 
        Err.Raise -99999999, "POPBILL", "아이템키가 입력되지 않았습니다."
    End If

    If MgtKey = "" Or isEmpty(MgtKey) Then
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Set AssignMgtKey = m_PopbillBase.httpPOST_ContentsType("/Cashbill/" & ItemKey,  _
                                m_PopbillBase.getSession_token(CorpNum), "", "MgtKey="+MgtKey, "", "application/x-www-form-urlencoded; charset=utf-8")
    
End Function

'현금영수증 임시저장
Public Function Register(CorpNum, ByRef Cashbill, UserID)
    Dim tmpDic : Set tmpDic = Cashbill.toJsonInfo()
    Dim postdata : postdata = m_PopbillBase.toString(tmpDic)

    Set Register = m_PopbillBase.httpPOST("/Cashbill", m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)
End Function 


'현금영수증 수정
Public Function Update(CorpNum, mgtKey, ByRef Cashbill, UserID)
    Dim tmpDic : Set tmpDic = Cashbill.toJsonInfo()
    Dim postdata : postdata = m_PopbillBase.toString(tmpDic)

    Set Update = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "PATCH", postdata, UserID)
End Function 


'현금영수증 발행
Public Function Issue(CorpNum, mgtKey, Memo, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "memo", Memo

    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    Set Issue = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "ISSUE", postdata, UserID)
End Function 


'현금영수증 발행취소
Public Function CancelIssue(CorpNum, mgtKey, Memo, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "memo", Memo

    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    Set CancelIssue = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "CANCELISSUE", postdata, UserID)
End Function 


'현금영수증 삭제
Public Function Delete(CorpNum, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Set Delete = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "DELETE", "", UserID)
End Function 


'현금영수증 상태정보 조회
Public Function GetInfo(CorpNum, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim result : Set result = m_PopbillBase.httpGet("/Cashbill/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), UserID)
    
    Dim infoObj : Set infoObj = New CashbillInfo	
    infoObj.fromJsonInfo result
    Set GetInfo = infoObj
End Function 


'다량 현금영수증 상태정보 조회
Public Function GetInfos(CorpNum, mgtKeyList, UserID)
    If isNull(mgtKeyList) Or isEmpty(mgtKeyList) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim tmp : Set tmp = JSON.parse("[]")

    Dim i
    For i=0 To UBound(mgtKeyList)-1
        tmp.Set i, mgtKeyList(i)
    Next

    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    
    Dim result : Set result = m_PopbillBase.httpPOST("/Cashbill/States", m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)

    Dim tmpDic : Set tmpDic = CreateObject("Scripting.Dictionary")

    For i=0 To result.length-1
        Dim cbInfo : Set cbInfo = New CashbillInfo 
        cbInfo.fromJsonInfo result.Get(i)
        tmpDic.Add i, cbInfo
    Next

    Set GetInfos = tmpDic
End Function 


'현금영수증 이력확인
Public Function GetLogs(CorpNum, mgtKey, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If
    
    Dim result : Set result = m_PopbillBase.httpGet("/Cashbill/"+mgtKey+"/Logs", m_PopbillBase.getSession_token(CorpNum),UserID)
    
    Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")

    Dim i
    For i=0 To result.length-1
        Dim chLog : Set chLog = New CashbillLog
        chLog.fromJsonInfo result.Get(i)
        tmp.Add i, chLog
    Next 

    Set GetLogs = tmp 
End Function 


'상세정보 확인
Public Function GetDetailInfo(CorpNum, mgtKey, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim result : Set result = m_PopbillBase.httpGet("/Cashbill/"+mgtKey+"?Detail", m_PopbillBase.getSession_token(CorpNum),UserID)

    Dim tmp : Set tmp = New Cashbill

    tmp.fromJsonInfo result
    
    Set GetDetailInfo = tmp
End Function 


'알림메일 재전송
Public Function SendEmail(CorpNum, mgtKey, Receiver, UsrID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "receiver", Receiver

    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Set SendEmail = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey , m_PopbillBase.getSession_token(CorpNum), "EMAIL", postdata, UserID)
End Function 


'알림문자 전송
Public Function SendSMS(CorpNum, mgtKey, Sender, Receiver, Contents, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "receiver", Receiver
    tmp.Set "sender", Sender
    tmp.Set "contents", Contents
    
    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Set SendSMS = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey , m_PopbillBase.getSession_token(CorpNum), "SMS", postdata, UserID)
End Function 


'팩스 전송
Public Function SendFAX(CorpNum, mgtKey, Sender, Receiver, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "문서번호가 입력되지 않았습니다."
    End If

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "receiver", Receiver
    tmp.Set "sender", Sender

    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Set SendFAX = m_PopbillBase.httpPOST("/Cashbill/"+mgtKey , m_PopbillBase.getSession_token(CorpNum), "FAX", postdata, UserID)
End Function 

'현금영수증 즉시발행
Public Function RegistIssue(CorpNum, ByRef Cashbill, Memo, UserID, EmailSubject)
    Dim tmpDic : Set tmpDic = Cashbill.toJsonInfo
    tmpDic.Set "memo", Memo

    If EmailSubject <> "" Then
        tmpDic.Set "emailSubject", EmailSubject
    End If

    Dim postdata : postdata = m_PopbillBase.toString(tmpDic)
    
    Set RegistIssue = m_PopbillBase.httpPOST("/Cashbill", m_PopbillBase.getSession_token(CorpNum), "ISSUE", postdata, UserID)
End Function 


'취소현금영수증 즉시발행. 2017/08/17 추가
Public Function RevokeRegistIssue(CorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID)

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "mgtKey", mgtKey
    tmp.Set "orgConfirmNum", orgConfirmNum
    tmp.Set "orgTradeDate", orgTradeDate
    tmp.Set "memo", memo
    
    If smssendYN Then
        tmp.Set "smssendYN", True
    End If

    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    
    Set RevokeRegistIssue = m_PopbillBase.httpPOST("/Cashbill", m_PopbillBase.getSession_token(CorpNum), "REVOKEISSUE", postdata, UserID)
End Function 

'취소현금영수증 임시저장. 2017/08/17 추가
Public Function RevokeRegister(CorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, userID)

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "mgtKey", mgtKey
    tmp.Set "orgConfirmNum", orgConfirmNum
    tmp.Set "orgTradeDate", orgTradeDate

    If smssendYN Then
        tmp.Set "smssendYN", True
    End If


    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    
    Set RevokeRegister = m_PopbillBase.httpPOST("/Cashbill", m_PopbillBase.getSession_token(CorpNum), "REVOKE", postdata, UserID)
End Function 

'부분취소 현금영수증 즉시발행
Public Function RevokeRegistIssue_Part(CorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID, isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount )

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "mgtKey", mgtKey
    tmp.Set "orgConfirmNum", orgConfirmNum
    tmp.Set "orgTradeDate", orgTradeDate
    tmp.Set "memo", memo
    
    If smssendYN Then
        tmp.Set "smssendYN", True
    End If

    tmp.Set "isPartCancel", isPartCancel
    tmp.Set "cancelType", cancelType
    tmp.Set "supplyCost", supplyCost
    tmp.Set "tax", tax
    tmp.Set "serviceFee", serviceFee
    tmp.Set "totalAmount", totalAmount

    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    
    Set RevokeRegistIssue_Part = m_PopbillBase.httpPOST("/Cashbill", m_PopbillBase.getSession_token(CorpNum), "REVOKEISSUE", postdata, UserID)
End Function 

'부분취소 현금영수증 임시저장
Public Function RevokeRegister_Part(CorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, userID, isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount)

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "mgtKey", mgtKey
    tmp.Set "orgConfirmNum", orgConfirmNum
    tmp.Set "orgTradeDate", orgTradeDate
    If smssendYN Then
        tmp.Set "smssendYN", True
    End If
    tmp.Set "isPartCancel", isPartCancel
    tmp.Set "cancelType", cancelType
    tmp.Set "supplyCost", supplyCost
    tmp.Set "tax", tax
    tmp.Set "serviceFee", serviceFee
    tmp.Set "totalAmount", totalAmount

    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    
    Set RevokeRegister_Part = m_PopbillBase.httpPOST("/Cashbill", m_PopbillBase.getSession_token(CorpNum), "REVOKE", postdata, UserID)
End Function 




'현금영수증 목록 조회
Public Function Search(CorpNum, DType, SDate, EDate, State, TradeType, TradeUsage, TradeOpt, TaxationType, Order, Page, PerPage, QString)
    If DType = "" Then
        Err.Raise -99999999, "POPBILL", "검색일자 유형이 입력되지 않았습니다."
    End If
    If SDate = "" Then
        Err.Raise -99999999, "POPBILL", "시작일자가 입력되지 않았습니다."
    End If
    If EDate = "" Then
        Err.Raise -99999999, "POPBILL", "종료일자가 이력되지 않았습니다."
    End If
    Dim uri
    uri = "/Cashbill/Search"
    uri = uri & "?DType=" & DType
    uri = uri & "&SDate=" & SDate
    uri = uri & "&EDate=" & EDate

    uri = uri & "&State="
    Dim i
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

    uri = uri & "&TradeOpt="
    For i=0 To UBound(TradeOpt) -1	
        If i = UBound(TradeOpt) -1 then
            uri = uri & TradeOpt(i)
        Else
            uri = uri & TradeOpt(i) & ","
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
    uri = uri & "&QString=" & QString
    uri = uri & "&Order=" & Order
    uri = uri & "&Page=" & CStr(Page)
    uri = uri & "&PerPage=" & CStr(PerPage)
    
    Dim searchResult : Set searchResult = New CBSearchResult
    Dim tmpObj : Set tmpObj = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), "")

    searchResult.fromJsonInfo tmpObj
    
    Set Search = searchResult
End Function

'알림메일 전송목록 조회
Public Function listEmailConfig(CorpNum, UserID)
    If CorpNum = "" Or isEmpty(CorpNum) Then 
        Err.Raise -99999999, "POPBILL", "사업자등록번호가 올바르지 않습니다."
    End If

    Dim result : Set result = m_PopbillBase.httpGet("/Cashbill/EmailSendConfig", m_PopbillBase.getSession_token(CorpNum), UserID)
    
    Dim tmpDic : Set tmpDic = CreateObject("Scripting.Dictionary")

    Dim i
    For i=0 To result.length-1
        Dim emailObj : Set emailObj = New EmailSendConfig	
        emailObj.fromJsonInfo result.Get(i)
        tmpDic.Add i, emailObj
    Next
    
    Set listEmailConfig = tmpDic
End Function 

'알림메일 전송설정 수정
Public Function updateEmailConfig(CorpNum, mailType, sendYN, UserID)
    If CorpNum = "" Or isEmpty(CorpNum) Then 
        Err.Raise -99999999, "POPBILL", "사업자등록번호가 올바르지 않습니다."
    End If

    If mailType = "" Or isEmpty(mailType) Then 
        Err.Raise -99999999, "POPBILL", "메일전송 타입이 입력되지 않았습니다."
    End If

    If sendYN = "" Or isEmpty(sendYN) Then 
        Err.Raise -99999999, "POPBILL", "메일전송 여부 항목이 입력되지 않았습니다."
    End If

    If (sendYN) Then
        sendYN="true"
    Else
        sendYN="false"
    End If
    
    Dim uri : uri = "/Cashbill/EmailSendConfig?EmailType="+mailType+"&SendYN="+sendYN

    Set updateEmailConfig = m_PopbillBase.httpPOST(uri, m_PopbillBase.getSession_token(CorpNum), "", "", UserID)
End Function

End Class

Class Cashbill
    Public mgtKey
    Public tradeDate
    Public tradeUsage
    Public tradeType
    Public tradeOpt
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
    Public cancelType

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
        mgtKey = jsonInfo.mgtKey
        tradeDate = jsonInfo.tradeDate
        tradeUsage = jsonInfo.tradeUsage
        tradeOpt = jsonInfo.tradeOpt
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

        cancelType = jsonInfo.cancelType

        On Error GoTo 0 
    End Sub 

    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.Set "mgtKey", mgtKey
        toJsonInfo.Set "tradeDate", tradeDate
        toJsonInfo.Set "tradeUsage", tradeUsage
        toJsonInfo.Set "tradeOpt", tradeOpt	
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
        toJsonInfo.Set "cancelType", cancelType
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
    Public tradeOpt

    Public totalAmount 
    Public tradeUsage 
    Public tradeType 
    Public stateCode 
    Public stateMemo
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
        tradeDate = jsonInfo.tradeDate
        issueDT = jsonInfo.issueDT
        customerName = jsonInfo.customerName 
        itemName = jsonInfo.itemName 
        identityNum = jsonInfo.identityNum 
        taxationType = jsonInfo.taxationType 
        tradeOpt = jsonInfo.tradeOpt

        totalAmount = jsonInfo.totalAmount 
        tradeUsage = jsonInfo.tradeUsage 
        tradeType = jsonInfo.tradeType 
        stateCode = jsonInfo.stateCode
        stateMemo = jsonInfo.stateMemo
        stateDT = jsonInfo.stateDT 
        printYN = jsonInfo.printYN 

        confirmNum = jsonInfo.confirmNum 
        orgTradeDate = jsonInfo.orgTradeDate 
        orgConfirmNum = jsonInfo.orgConfirmNum 

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
        Dim i
        For i = 0 To jsonInfo.list.length -1
            Dim tmpObj : Set tmpObj = New CashbillInfo
            tmpObj.fromJsonInfo jsonInfo.list.Get(i)
            Set list(i) = tmpObj
        Next

        On Error GoTo 0
    End Sub
End Class



Class EmailSendConfig
    Public emailType
    Public sendYN

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
        emailType = jsonInfo.emailType
        sendYN = jsonInfo.sendYN
        On Error GoTo 0 
    End Sub 

    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.Set "emailType", emailType
        toJsonInfo.Set "sendYN", sendYN
    End Function 
End Class
%>