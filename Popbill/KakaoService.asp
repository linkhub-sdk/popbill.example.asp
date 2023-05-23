<%
Class KakaoService

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

Public Property Let UseGAIP(ByVal value)
    m_PopbillBase.UseGAIP = value
End Property

Public Property Let UseLocalTimeYN(ByVal value)
    m_PopbillBase.UseLocalTimeYN = value
End Property

Public Sub Class_Initialize
    Set m_PopbillBase = New PopbillBase
    m_PopbillBase.AddScope("153")
    m_PopbillBase.AddScope("154")
    m_PopbillBase.AddScope("155")
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
Public Function GetChargeInfo ( CorpNum, KType, UserID )
    Dim result : Set result = m_PopbillBase.httpGET ( "/KakaoTalk/ChargeInfo?Type=" &KType, m_PopbillBase.getSession_token(CorpNum), UserID )

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

'카카오톡 관련 URL
Public Function GetURL(CorpNum, UserID, TOGO)
    Dim result
    If TOGO = "SENDER" Then
        Set result = m_PopbillBase.httpGet("/Message/?TG="+TOGO, m_PopbillBase.getSession_token(CorpNum), UserID)
    Else
        Set result = m_PopbillBase.httpGet("/KakaoTalk/?TG="+TOGO, m_PopbillBase.getSession_token(CorpNum), UserID)
    End If
    GetURL = result.url
End Function

'플러스친구 계정관리 팝업 URL
Public Function GetPlusFriendMgtURL(CorpNum, UserID)
    Set result = m_PopbillBase.httpGet("/KakaoTalk/?TG=PLUSFRIEND", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetPlusFriendMgtURL = result.url
End Function

'발신번호 관리 팝업 URL
Public Function GetSenderNumberMgtURL(CorpNum, UserID)
    Set result = m_PopbillBase.httpGet("/Message/?TG=SENDER", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetSenderNumberMgtURL = result.url
End Function

'알림톡 템플릿관리 팝업 URL
Public Function GetATSTemplateMgtURL(CorpNum, UserID)
    Set result = m_PopbillBase.httpGet("/KakaoTalk/?TG=TEMPLATE", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetATSTemplateMgtURL = result.url
End Function

'카카오톡 전송내역 팝업 URL
Public Function GetSentListURL(CorpNum, UserID)
    Set result = m_PopbillBase.httpGet("/KakaoTalk/?TG=BOX", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetSentListURL = result.url
End Function


'플러스친구 계정 목록 확인
Public Function ListPlusFriendID(CorpNum)
    Set ListPlusFriendID = m_PopbillBase.httpGET("/KakaoTalk/ListPlusFriendID", m_PopbillBase.getSession_token(CorpNum), "")
End Function

'발신번호 목록 확인
Public Function GetSenderNumberList(CorpNum)
    Set GetSenderNumberList = m_PopbillBase.httpGET("/Message/SenderNumber", m_PopbillBase.getSession_token(CorpNum), "")
End Function

'발신번호 등록여부 확인
Public Function CheckSenderNumber(CorpNum, SenderNumber, UserID)
    If SenderNumber = "" Or IsNull(SenderNumber) Then
        Err.Raise -99999999, "POPBILL", "발신번호가 입력되지 않았습니다."
    End If

    Set CheckSenderNumber = m_PopbillBase.httpGET("/KakaoTalk/CheckSenderNumber/"&SenderNumber,m_PopbillBase.getSession_token(CorpNum),UserID)
End Function

'단가확인
Public Function GetUnitCost(CorpNum, KType)
    Dim result : Set result = m_PopbillBase.httpGET("/KakaoTalk/UnitCost?Type="&KType, m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'알림톡 템플릿 정보 확인
Public Function GetATSTemplate(CorpNum, templateCode, UserID)

    If templateCode = "" Or isEmpty(templateCode) Then
        Err.raise -99999999, "POPBILL", "템플릿코드가 입력되지 않았습니다."
    End If

    Dim template : Set template = new KakaoATSTemplate
    Dim result : Set result = m_PopbillBase.httpGet("/KakaoTalk/GetATSTemplate/" + templateCode, m_PopbillBase.getSession_token(CorpNum), UserID)

    template.fromJsonInfo result

    Set GetATSTemplate = template

End Function

'알림톡 템플릿 목록 확인
Public Function ListATSTemplate(CorpNum)
    Dim result : Set result = m_PopbillBase.httpGET("/KakaoTalk/ListATSTemplate", m_PopbillBase.getSession_token(CorpNum), "")

    Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")

    Dim i
    For i=0 To result.length-1
        Dim atsList : Set atsList = New KakaoATSTemplate
        atsList.fromJsonInfo result.Get(i)
        tmp.Add i, atsList
    Next
    Set ListATSTemplate = tmp
End Function

'예약전송취소
Public Function CancelReserve(CorpNum, ReceiptNum, UserID)
    If ReceiptNum = "" Or IsNull(ReceptNum) Then
        Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다"
    End If

    Set CancelReserve = m_PopbillBase.httpGet("/KakaoTalk/"&ReceiptNum&"/Cancel",m_PopbillBase.getSession_token(CorpNum),UserID)
End Function


'예약 전송취소 (요청번호 할당)
Public Function CancelReserveRN(CorpNum, RequestNum, UserID)
    If RequestNum = "" Or IsNull(RequestNum) Then
        Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다"
    End If

    Set CancelReserveRN = m_PopbillBase.httpGet("/KakaoTalk/Cancel/"&RequestNum, m_PopbillBase.getSession_token(CorpNum), UserID)
End Function

'알림톡 전송
Public Function SendATS(CorpNum, templateCode, senderNum, content, altContent, altSendType, reserveDT, receiverList, RequestNum, UserID, btnList, altSubject)
    If templateCode = "" Or IsNull(templateCode) Then
        Err.Raise -99999999, "POPBILL", "알림톡 템플릿 코드(TemplateCode)가 입력되지 않았습니다"
    End If

    Dim tmp : Set tmp = JSON.parse("{}")

    If templateCode <> "" Then tmp.Set "templateCode", templateCode
    If senderNum <> "" Then tmp.Set "snd", senderNum
    If content <> "" Then tmp.Set "content", content
    If altContent <> "" Then tmp.Set "altContent", altContent
    If altSendType <> "" Then tmp.Set "altSendType", altSendType
    If reserveDT <> "" Then tmp.Set "sndDT", reserveDT
    If RequestNum <> "" Then tmp.Set "RequestNum", RequestNum
    If altSubject <> "" Then tmp.Set "altSubject", altSubject

    Dim msgs : Set msgs = JSON.parse("[]")

    Dim i
    For i=0 To receiverList.Count-1
        Dim msgObj : Set msgObj = New KakaoReceiver
        msgObj.setValue receiverList.Item(i)
        msgs.Set i, msgObj.toJsonInfo
    Next

    tmp.Set "msgs", msgs

    If False = IsNull(btnList)  Then
        Dim btns : Set btns = JSON.parse("[]")
        For i=0 To btnList.Count -1
            Dim btnObj : Set btnObj = New KakaoButton
            btnObj.setValue btnList.Item(i)
            btns.Set i, btnObj.toJsonInfo
        Next

        tmp.Set "btns", btns
    End If

    Dim postData : postData = m_PopbillBase.toString(tmp)

    Dim result : Set result = m_PopbillBase.httpPost("/ATS", m_PopbillBase.getSession_Token(CorpNum), "", postData, UserID)
    SendATS = result.ReceiptNum
End Function


'친구톡 텍스트 전송
Public Function SendFTS(CorpNum, plusFriendID, snd, content, altContent, altSendType, sndDT, adsYN, receiverList, btnList, RequestNum, UserID, altSubject)

    If plusFriendID = "" Or IsNull(plusFriendID) Then
        Err.Raise -99999999, "POPBILL", "친구톡 플러스친구 아이디(plusFriendID)가 입력되지 않았습니다"
    End If

    Dim tmp : Set tmp = JSON.parse("{}")

    If plusFriendID <> "" Then tmp.Set "plusFriendID", plusFriendID
    If senderNum <> "" Then tmp.Set "snd", senderNum
    If content <> "" Then tmp.Set "content", content
    If altContent <> "" Then tmp.Set "altContent", altContent
    If altSendType <> "" Then tmp.Set "altSendType", altSendType
    If reserveDT <> "" Then tmp.Set "sndDT", reserveDT
    If adsYN Then tmp.Set "adsYN", adsYN
    If RequestNum <> "" Then tmp.Set "RequestNum", RequestNum
    If altSubject <> "" Then tmp.Set "altSubject", altSubject

    Dim msgs : Set msgs = JSON.parse("[]")

    Dim i
    For i=0 To receiverList.Count-1
        Dim msgObj : Set msgObj = New KakaoReceiver
        msgObj.setValue receiverList.Item(i)
        msgs.Set i, msgObj.toJsonInfo
    Next

    tmp.Set "msgs", msgs

    Dim btns : Set btns = JSON.parse("[]")

    For i=0 To btnList.Count -1
        Dim btnObj : Set btnObj = New KakaoButton
        btnObj.setValue btnList.Item(i)
        btns.Set i, btnObj.toJsonInfo
    Next

    tmp.Set "btns", btns

    Dim postData : postData = m_PopbillBase.toString(tmp)

    Dim result : Set result = m_PopbillBase.httpPost("/FTS", m_PopbillBase.getSession_Token(CorpNum), "", postData, UserID)
    SendFTS = result.ReceiptNum
End Function


'친구톡 이미지 전송
Public Function SendFMS(CorpNum, plusFriendID, snd, content, altContent, altSendType, sndDT, adsYN, receiverList, btnList, filePath, imageURL, RequestNum, UserID, altSubject)

    If plusFriendID = "" Or IsNull(plusFriendID) Then
        Err.Raise -99999999, "POPBILL", "친구톡 플러스친구 아이디(plusFriendID)가 입력되지 않았습니다"
    End If

    Dim tmp : Set tmp = JSON.parse("{}")

    If plusFriendID <> "" Then tmp.Set "plusFriendID", plusFriendID
    If senderNum <> "" Then tmp.Set "snd", senderNum
    If content <> "" Then tmp.Set "content", content
    If altContent <> "" Then tmp.Set "altContent", altContent
    If altSendType <> "" Then tmp.Set "altSendType", altSendType
    If reserveDT <> "" Then tmp.Set "sndDT", reserveDT
    If imageURL <> "" Then tmp.Set "imageURL", imageURL
    If adsYN Then tmp.Set "adsYN", adsYN
    If RequestNum <> "" Then tmp.Set "RequestNum", RequestNum
    If altSubject <> "" Then tmp.Set "altSubject", altSubject

    Dim msgs : Set msgs = JSON.parse("[]")

    Dim i
    For i=0 To receiverList.Count-1
        Dim msgObj : Set msgObj = New KakaoReceiver
        msgObj.setValue receiverList.Item(i)
        msgs.Set i, msgObj.toJsonInfo
    Next

    tmp.Set "msgs", msgs

    Dim btns : Set btns = JSON.parse("[]")

    For i=0 To btnList.Count -1
        Dim btnObj : Set btnObj = New KakaoButton
        btnObj.setValue btnList.Item(i)
        btns.Set i, btnObj.toJsonInfo
    Next

    tmp.Set "btns", btns

    Dim postData : postData = m_PopbillBase.toString(tmp)
    Dim result : Set result = m_PopbillBase.httpPost_Files("/FMS", m_PopbillBase.getSession_Token(CorpNum), postData, filePath, UserID)
    SendFMS = result.ReceiptNum
End Function


'카카오톡 전송내역 확인
Public Function GetMessages(CorpNum, ReceiptNum, UserID)
    If ReceiptNum = "" Or IsNull(ReceptNum) Then
        Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다"
    End If

    Dim result : Set result = m_PopbillBase.httpGet("/KakaoTalk/"&ReceiptNum,m_PopbillBase.getSession_token(CorpNum),UserID)

    Dim resultObj : Set resultObj = New KakaoSentResult

    resultObj.fromJsonInfo result

    Set GetMessages = resultObj
End Function

'카카오톡 전송내역 확인 (요청번호 할당)
Public Function GetMessagesRN(CorpNum, RequestNum, UserID)
    If RequestNum = "" Or IsNull(RequestNum) Then
        Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다"
    End If

    Dim result : Set result = m_PopbillBase.httpGet("/KakaoTalk/Get/"+RequestNum ,m_PopbillBase.getSession_token(CorpNum), UserID)

    Dim resultObj : Set resultObj = New KakaoSentResult

    resultObj.fromJsonInfo result

    Set GetMessagesRN = resultObj

End Function


'카카오톡 전송내역 목록 조회
Public Function Search(CorpNum, SDate, EDate, State, Item, ReserveYN, SenderYN, Order, Page, PerPage, QString)
    If SDate = "" Then
        Err.Raise -99999999, "POPBILL", "시작일자가 입력되지 않았습니다."
    End If
    If EDate = "" Then
        Err.Raise -99999999, "POPBILL", "종료일자가 이력되지 않았습니다."
    End If

    Dim uri
    uri = "/KakaoTalk/Search"
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

    uri = uri & "&ReserveYN=" & ReserveYN


    If SenderYN Then
        uri = uri & "&SenderYN=1"
    Else
        uri = uri & "&SenderYN=0"
    End If

    uri = uri & "&Order=" & Order
    uri = uri & "&Page=" & CStr(Page)
    uri = uri & "&PerPage=" & CStr(PerPage)
    uri = uri & "&QString=" & Server.URLEncode(QString)

    Dim searchResult : Set searchResult = New KakaoSearchResult
    Dim tmpObj : Set tmpObj = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), "")

    searchResult.fromJsonInfo tmpObj

    Set Search = searchResult
End Function

'예약전송 일부 취소 (접수번호)
Public Function CancelReservebyRCV(CorpNum, ReceiptNum,receiverNum, UserID)
    IF isNull(ReceiptNum) Or ReceiptNum = "" Then
        Err.raise ReceiptNum, "POPBILL", "접수번호가 입력되지 않았습니다."
    End IF
    IF isNull(receiverNum) Or receiverNum = "" Then
        Err.raise receiverNum, "POPBILL", "수신번호가 입력되지 않았습니다."
    End IF

    Set m_CancelReserve = New CancelReserveObj
    m_CancelReserve.receiverNum = receiverNum
    Dim tmp : Set tmp = m_CancelReserve.toJsonInfo

    Dim uri : uri = "/KakaoTalk/" & ReceiptNum & "/Cancel"
    Dim postData: postData = m_popbillBase.toString(tmp)

    Set CancelReservebyRCV = m_popbillBase.httpPOST(uri, m_PopbillBase.getSession_token(CorpNum), "", postData, UserID)
End Function

'예약전송 일부 취소 (전송 요청번호)
Public Function CancelReserveRNbyRCV(CorpNum, RequestNum, receiverNum, UserID)
    IF isNull(RequestNum) Or RequestNum = "" Then
        Err.raise RequestNum, "POPBILL", "전송요청번호가 입력되지 않았습니다."
    End IF
    IF isNull(receiverNum) Or  receiverNum = "" Then
        Err.raise receiverNum, "POPBILL", "수신번호가 입력되지 않았습니다."
    End IF

    Set m_CancelReserve = New CancelReserveObj
    m_CancelReserve.receiverNum = receiverNum
    Dim tmp : Set tmp = m_CancelReserve.toJsonInfo

    Dim uri : uri = "/KakaoTalk/Cancel/" & RequestNum
    Dim postData: postData = m_popbillBase.toString(tmp)

    Set CancelReserveRNbyRCV = m_popbillBase.httpPOST(uri, m_PopbillBase.getSession_token(CorpNum), "", postData, UserID)
End Function


End Class ' end of KakaoService class

Class KakaoSentDetail
    Public state
    Public sendDT
    Public ReceiveNum
    Public receiveName
    Public contentType
    Public content
    Public result
    Public resultDT
    Public altSubject
    Public altContent
    Public altContentType
    Public altSendDT
    Public altResult
    Public altResultDT
    Public RequestNum
    Public ReceiptNum
    Public interOPRefKey

    Public Sub fromJsonInfo(detailInfo)
        On Error Resume Next
            state = detailInfo.state
            sendDT = detailInfo.sendDT
            ReceiveNum = detailInfo.ReceiveNum
            receiveName = detailInfo.receiveName
            contentType = detailInfo.contentType
            content = detailInfo.content
            result = detailInfo.result
            resultDT = detailInfo.resultDT
            altSubject = detailInfo.altSubject
            altContent = detailInfo.altContent
            altContentType = detailInfo.altContentType
            altSendDT = detailInfo.altSendDT
            altResult = detailInfo.altResult
            altResultDT = detailInfo.altResultDT
            RequestNum = detailInfo.RequestNum
            ReceiptNum = detailInfo.ReceiptNum
            interOPRefKey = detailInfo.interOPRefKey
        on Error GoTo 0
    End Sub
End Class  ' end of KakaoSentDetail Class


Class KakaoSearchResult
    Public code
    Public message
    Public total
    Public perPage
    Public pageNum
    Public pageCount
    Public list()

    Public Sub Class_Initialize
        ReDim list(-1)
    End Sub

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
            code = jsonInfo.code
            message = jsonInfo.message
            total = jsonInfo.total
            perPage = jsonInfo.perPage
            pageNum = jsonInfo.pageNum
            pageCount = jsonInfo.pageCount

            ReDim list(jsonInfo.list.length)
            Dim i
            For i = 0 To jsonInfo.list.length-1
                Dim tmpObj : Set tmpObj = New KakaoSentDetail
                tmpObj.fromJsonInfo jsonInfo.list.Get(i)
                Set list(i) = tmpObj
            Next
        On Error GoTo 0
    End Sub
End Class ' End of KakaoSearchResult

Class KakaoSentResult
    Public contentType
    Public templateCode
    Public plusFriendID
    Public sendNum
    Public altSubject
    Public altContent
    Public altSendType
    Public reserveDT
    Public adsYN
    Public imageURL
    Public sendCnt
    Public successCnt
    Public failCnt
    Public altCnt
    Public cancelCnt
    Public btns()
    Public msgs()

    Public Sub Class_Initialize
        ReDim btns(-1)
        ReDim msgs(-1)
    End Sub

    Public Sub fromJsonInfo(detailInfo)

        On Error Resume Next
            contentType = detailInfo.contentType
            templateCode = detailInfo.templateCode
            plusFriendID = detailInfo.plusFriendID
            sendNum = detailInfo.sendNum
            altSubject = detailInfo.altSubject
            altContent = detailInfo.altContent
            altSendType = detailInfo.altSendType
            reserveDT = detailInfo.reserveDT
            adsYN = detailInfo.adsYN
            imageURL = detailInfo.imageURL
            sendCnt = detailInfo.sendCnt
            successCnt = detailInfo.successCnt
            failCnt = detailInfo.failCnt
            altCnt = detailInfo.altCnt
            cancelCnt = detailInfo.cancelCnt

            ReDim btns(detailInfo.btns.length)
            Dim i, tmpObj
            For i = 0 To detailInfo.btns.length -1
                tmpObj : Set tmpObj = New KakaoButton
                tmpObj.fromJsonInfo detailInfo.btns.Get(i)
                Set btns(i) = tmpObj
            Next

            ReDim msgs(detailInfo.msgs.length)
            For i = 0 To detailInfo.msgs.length -1
                tmpObj : Set tmpObj = New KakaoSentDetail
                tmpObj.fromJsonInfo detailInfo.msgs.Get(i)
                Set msgs(i) = tmpObj
            Next
        On Error GoTo 0

    End Sub
End Class ' End of KakaoSentResult class

Class KakaoReceiver
    Public rcv
    Public rcvnm
    Public msg
    Public altsjt
    Public altmsg
    Public interOPRefKey
    Public btns()

    Public Sub Class_Initialize
        ReDim btns(-1)
    End Sub

    Public Function toJsonInfo()

        Set toJsonInfo = JSON.parse("{}")
        If rcv <> "" Then toJsonInfo.Set "rcv", rcv
        If rcvnm <> "" Then toJsonInfo.Set "rcvnm", rcvnm
        If msg <> "" Then toJsonInfo.Set "msg", msg
        If altsjt <> "" Then toJsonInfo.Set "altsjt", altsjt
        If altmsg <> "" Then toJsonInfo.Set "altmsg", altmsg
        If interOPRefKey <> "" Then toJsonInfo.Set "interOPRefKey", interOPRefKey
        IF Ubound(btns) >= -1 Then
            Dim btnsJsonInfo()
            ReDim btnsJsonInfo(UBound(btns))
            Dim i, btn
            i = 0
            For Each btn In btns
                Set btnsJsonInfo(i) = btns(i).toJsonInfo
                i = i + 1
            Next
        End If
        toJsonInfo.set "btns", btnsJsonInfo
    End Function

    Public Sub setValue(msgList)
        On Error Resume Next
        rcv = msgList.rcv
        rcvnm = msgList.rcvnm
        msg = msgList.msg
        altsjt = msgList.altsjt
        altmsg = msgList.altmsg
        interOPRefKey = msgList.interOPRefKey

        Redim btns(Ubound(msgList.btns))
        Dim i : i = 0
        For Each btn In msgList.btns
            Set btns(i) = btn
            i = i + 1
        Next
        On Error GoTo 0
    End Sub

    Public Sub AddBtn(btnInfo)
        Redim Preserve btns(Ubound(btns) + 1)
        Set btns(Ubound(btns)) = btnInfo
    End Sub
End Class  ' End of KakaoReceiver class

Class KakaoButton
    Public n
    Public t
    Public u1
    Public u2

    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        If n <> "" Then  toJsonInfo.set "n", n
        If t <> "" Then  toJsonInfo.set "t", t
        If u1 <> "" Then  toJsonInfo.set "u1", u1
        If u2 <> "" Then  toJsonInfo.set "u2", u2
    End Function

    Public Sub setValue(btnInfo)
        n = btnInfo.n
        t = btnInfo.t
        u1 = btnInfo.u1
        u2 = btnInfo.u2
    End Sub

    Public Sub fromJsonInfo(btnInfo)
        On Error Resume Next
            n = btnInfo.n
            t = btnInfo.t
            u1 = btnInfo.u1
            u2 = btnInfo.u2
        On Error GoTo 0
    End Sub
End Class ' End of KakaoButton class

Class KakaoATSTemplate
    Public templateCode
    Public templateName
    Public template
    Public plusFriendID
    Public ads
    Public appendix
    Public secureYN
    Public state
    Public stateDT
    Public btns()

    Public Sub Class_Initialize
        ReDim btns(-1)
    End Sub

    Public Sub fromJsonInfo(atsInfo)
        On Error Resume Next
            templateCode = atsInfo.templateCode
            templateName = atsInfo.templateName
            template = atsInfo.template
            plusFriendID = atsInfo.plusFriendID
            ads = atsInfo.ads
            appendix = atsInfo.appendix
            secureYN = atsInfo.secureYN
            state = atsInfo.state
            stateDT = atsInfo.stateDT

            ReDim btns(atsInfo.btns.length)
            Dim i
            For i = 0 To atsInfo.btns.length -1
                Dim tmpObj : Set tmpObj = New KakaoButton
                tmpObj.fromJsonInfo atsInfo.btns.Get(i)
                Set btns(i) = tmpObj
            Next
        On Error GoTo 0
    End Sub
End Class ' end of KakaoATSTemplate

Class CancelReserveObj
    Public receiverNum

    Public Function toJsonInfo()

        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.Set "receiverNum", receiverNum
    End Function
End Class

%>