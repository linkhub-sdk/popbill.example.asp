<!--#include file="Linkhub/Linkhub.asp"-->
<%

Application("LINKHUB_TOKEN_SCOPE_POPBILL") = Array("member")
Const ServiceID_REAL = "POPBILL"
Const ServiceID_TEST = "POPBILL_TEST"
Const ServiceURL_REAL = "https://popbill.linkhub.co.kr"
Const ServiceURL_TEST = "https://popbill-test.linkhub.co.kr"

Const ServiceURL_Static_REAL = "https://static-popbill.linkhub.co.kr"
Const ServiceURL_Static_TEST = "https://static-popbill-test.linkhub.co.kr"

Const ServiceURL_GA_REAL = "https://ga-popbill.linkhub.co.kr"
Const ServiceURL_GA_TEST = "https://ga-popbill-test.linkhub.co.kr"

Const APIVersion = "1.0"
Const adTypeBinary = 1
Const adTypeText = 2

Class PopbillBase

Private m_IsTest
Private m_TokenDic
Private m_Linkhub
Private m_IPRestrictOnOff
Private m_UseStaticIP
Private m_UseGAIP
Private m_UseLocalTimeYN

'테스트 플래그
Public Property Let IsTest(ByVal value)
    m_IsTest = value
End Property

Public Property Let IPRestrictOnOff(ByVal value)
    m_IPRestrictOnOff = value
End Property

Public Property Let UseStaticIP(ByVal value)
    m_UseStaticIP = value
End Property

Public Property Let UseGAIP(ByVal value)
    m_UseGAIP = value
End Property

Public Property Let UseLocalTimeYN(ByVal value)
    m_UseLocalTimeYN = value
End Property

Public Sub Class_Initialize
    On Error Resume next
    If  Not(POPBILL_TOKEN_CACHE Is Nothing) Then
        Set m_TokenDic = POPBILL_TOKEN_CACHE
    Else
        Set m_TokenDic = server.CreateObject("Scripting.Dictionary")
    End If
    On Error GoTo 0

    If isEmpty( m_TokenDic) Then
        Set m_TokenDic = server.CreateObject("Scripting.Dictionary")
    End If

    m_IsTest = False
    m_IPRestrictOnOff = True
    m_UseStaticIP = False
    m_UseGAIP = False
    m_UseLocalTimeYN = True
    Set m_Linkhub = New Linkhub
End Sub

Public Sub Class_Terminate
    Set m_Linkhub = Nothing
End Sub

Private Property Get m_scope
    m_scope = Application("LINKHUB_TOKEN_SCOPE_POPBILL")
End Property

Public Sub AddScope(scope)
    Dim t : t = Application("LINKHUB_TOKEN_SCOPE_POPBILL")
    ReDim Preserve t(Ubound(t)+1)
    t(Ubound(t)) = scope
    Application("LINKHUB_TOKEN_SCOPE_POPBILL") = t
End Sub


Public Sub Initialize(linkID, SecretKey )
    m_Linkhub.LinkID = linkID
    m_Linkhub.SecretKey = SecretKey
End Sub

Public Function getSession_token(CorpNum)
    Dim refresh :  refresh = False
    Dim m_Token : Set m_Token = Nothing

    If m_TokenDic.Exists(CorpNum) Then
        Set m_Token = m_TokenDic.Item(CorpNum)
    End If

    If m_Token Is Nothing Then
        refresh = True
    Else
        'CheckScope
        Dim scope
        For Each scope In m_scope
            If InStr(m_Token.strScope,scope) = 0 Then
                refresh = True
                Exit for
            End if
        Next
        If refresh = False then
            Dim utcnow : utcnow = CDate(Replace(left(m_linkhub.getTime(m_UseStaticIP, m_UseLocalTimeYN, m_UseGAIP),19),"T" , " " ))
            refresh = CDate(Replace(left(m_Token.expiration,19),"T" , " " )) < utcnow
        End if
    End If

    If refresh Then
        If m_TokenDic.Exists(CorpNum) Then m_TokenDic.remove CorpNum
        Set m_Token = m_Linkhub.getToken(IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), CorpNum, m_scope, IIf(m_IPRestrictOnOff, "", "*"), m_UseStaticIP, m_UseLocalTimeYN, m_UseGAIP)
        m_Token.set "strScope", Join(m_scope,"|")
        m_TokenDic.Add CorpNum, m_Token
    End If

    getSession_token = m_Token.session_token

End Function

'회원잔액조회
Public Function GetBalance(CorpNum)
    GetBalance = m_Linkhub.GetBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), m_UseStaticIP, m_UseGAIP)
End Function

'파트너 잔액조회
Public Function GetPartnerBalance(CorpNum)
    GetPartnerBalance = m_Linkhub.GetPartnerBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), m_UseStaticIP, m_UseGAIP)
End Function

'파트너 포인트 충전 URL - 2017/08/29 추가
Public Function GetPartnerURL(CorpNum, TOGO)
    GetPartnerURL = m_Linkhub.GetPartnerURL(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), TOGO, m_UseStaticIP, m_UseGAIP)
End Function

'팝빌 기본 URL
Public Function GetPopbillURL(CorpNum , UserID , TOGO)

    Dim result : Set result = httpGET("/?TG=" + TOGO, getSession_token(CorpNum), UserID)
    GetPopbillURL = result.url
End Function

'팝빌 로그인 URL
Public Function GetAccessURL(CorpNum , UserID)

    Dim result : Set result = httpGET("/?TG=LOGIN", getSession_token(CorpNum), UserID)
    GetAccessURL = result.url
End Function

'팝빌 연동회원 포인트 충전 URL
Public Function GetChargeURL(CorpNum , UserID)

    Dim result : Set result = httpGET("/?TG=CHRG", getSession_token(CorpNum), UserID)
    GetChargeURL = result.url
End Function

'팝빌 연동회원 포인트 결제내역 URL
Public Function GetPaymentURL(CorpNum, UserID)

    Dim result : Set result = httpGET("/?TG=PAYMENT", getSession_token(CorpNum), UserID)
    GetPaymentURL = result.url
End Function

'팝빌 연동회원 포인트 사용내역 URL
Public Function GetUseHistoryURL(CorpNum, UserID)

    Dim result : Set result = httpGET("/?TG=USEHISTORY", getSession_token(CorpNum), UserID)
    GetUseHistoryURL = result.url
End Function

'회원가입 여부
Public Function CheckIsMember(CorpNum , linkID)
    Set CheckIsMember = httpGET("/Join?CorpNum=" + CorpNum + "&LID=" + linkID, "","")
End Function

'회원가입
Public Function JoinMember(JoinInfo)
    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.set "LinkID", JoinInfo.linkID
    tmp.set "CorpNum", JoinInfo.CorpNum
    tmp.set "CorpName", JoinInfo.CorpName
    tmp.set "CEOName", JoinInfo.CEOName
    tmp.set "Addr", JoinInfo.Addr
    tmp.set "ZipCode", JoinInfo.ZipCode
    tmp.set "BizClass", JoinInfo.BizClass
    tmp.set "BizType", JoinInfo.BizType
    tmp.set "ContactName", JoinInfo.contactName
    tmp.set "ContactEmail", JoinInfo.ContactEmail
    tmp.set "ContactFAX", JoinInfo.ContactFAX
    tmp.set "ContactHP", JoinInfo.ContactHP
    tmp.set "ContactTEL", JoinInfo.ContactTEL
    tmp.set "ID", JoinInfo.ID
    tmp.set "PWD", JoinInfo.PWD
    tmp.set "Password", JoinInfo.Password
    Dim postdata : postdata = m_Linkhub.toString(tmp)

    Set JoinMember = httpPOST("/Join", "", "", postdata, "")
End Function

'담당자 정보 확인
Public Function GetContactInfo(CorpNum, ContactID, UserID)

    postdata = "{'id':" + "'" + ContactID  +"'}"

    Dim contInfo : Set contInfo = New ContactInfo
    Dim result : Set result = httpPOST("/Contact", getSession_token(CorpNum), "", postdata, UserID)

    contInfo.fromJsonInfo result

    Set GetContactInfo = contInfo
End Function

' 담당자 목록조회
Public Function ListContact(CorpNum, UserID)
    Dim result : Set result = httpGET("/IDs",getSession_token(CorpNum), UserID)

    Dim infoObj : Set infoObj = CreateObject("Scripting.Dictionary")
    Dim i
    For i = 0 To result.length - 1
        Dim contInfo : Set contInfo = New ContactInfo
        contInfo.fromJsonInfo result.Get(i)
        infoObj.Add i, contInfo
    Next

    Set ListContact = infoObj
End Function

'담당자 수정
Public Function UpdateContact(CorpNum, ContactInfo, UserID)
    Dim tmp : Set tmp = ContactInfo.toJsonInfo
    Dim postdata : postdata = m_Linkhub.toString(tmp)

    Set UpdateContact = httpPOST("/IDs", getSession_token(CorpNum), "", postdata, UserID)
End Function

'담당자 추가
Public Function RegistContact(CorpNum, ContactInfo, UserId)
    Dim tmp : Set tmp = ContactInfo.toJsonInfo
    Dim postdata : postdata = m_Linkhub.toString(tmp)

    Set RegistContact = httpPOST("/IDs/New", getSession_token(CorpNum), "", postdata, UserId)
End Function

'회사정보 확인
Public Function GetCorpInfo(CorpNum, UserID)
    Dim result : Set result = httpGET("/CorpInfo",getSession_token(CorpNum), UserID)

    Dim infoObj : Set infoObj = New CorpInfo
    infoObj.fromJsonInfo result

    Set GetCorpInfo = infoObj
End Function

'회사정보 수정
Public Function UpdateCorpInfo(CorpNum, CorpInfo, UserID)
    Dim tmp : Set tmp = CorpInfo.toJsonInfo
    Dim postdata : postdata = m_Linkhub.toString(tmp)

    Set UpdateCorpInfo = httpPOST("/CorpInfo", getSession_token(CorpNum), "", postdata, UserID)
End Function

'아이디 중복확인
Public Function CheckID(id)
    Set CheckID = httpGET("/IDCheck?ID="+id, "", "")
End Function

' 무통장 입금신청 (PaymentRequest)
Public Function PaymentRequest(CorpNum, PaymentForm, UserID)
    Dim tmp: Set tmp = PaymentForm.toJsonInfo
    Dim postData: postData = m_Linkhub.toString(tmp)
    Dim t_result : Set t_result = httpPOST("/Payment", getSession_token(CorpNum), "", postData, UserID)

    Dim m_paymentResult : Set m_paymentResult = New PaymentResponse
    m_paymentResult.fromJsonInfo t_result
    Set PaymentRequest = m_paymentResult
End Function

' 무통장 입금신청 정보확인 (GetSettleResult)
Public Function GetSettleResult(CorpNum, SettleCode, UserID)
    If SettleCode = "" then
        Err.Raise -99999999, "POPBILL", "정산코드가 입력되지 않았습니다."
    End If

    Dim tmp : Set tmp = httpGET("/Paymet/"& SettleCode,getSession_token(CorpNum),UserID)
    Dim m_paymentHistory: Set m_paymentHistory = New PaymentHistory
    m_paymentHistory.fromJsonInfo  tmp

    Set GetSettleResult = m_paymentHistory
End Function

' 포인트 사용내역 (GetUseHistory)
Public Function GetUseHistory(CorpNum, SDate, EDate, Page, PerPage, Order, UserID)
    Dim tmp : Set tmp = httpGET("/UseHistory?SDate=" & SDate &  "&EDate=" & EDate & "&Page=" & Page & "&PerPage=" & PerPage &  "&Order=" & Order,getSession_token(CorpNum),UserID)
    Dim infoObj : Set infoObj = CreateObject("Scripting.Dictionary")

    Dim useHistoryResult : Set useHistoryResult = New UseHistoryResult
    useHistoryResult.fromJsonInfo tmp


    Set GetUseHistory = useHistoryResult
End Function

' 포인트 결제내역 (GetPaymentHistory)
Public Function GetPaymentHistory(CorpNum, SDate, EDate, Page, PerPage, UserID)
    Dim tmp: Set tmp = httpGET("/PaymentHistory?SDate=" &SDate&  "&EDate="&EDate &  "&Page="&Page&  "&PerPage=" &PerPage,getSession_token(CorpNum),UserID)

    Dim infoObj : Set infoObj = CreateObject("Scripting.Dictionary")
    Dim paymentHistoryResult : Set paymentHistoryResult = New PaymentHistoryResult
    paymentHistoryResult.fromJsonInfo mytmp

    Set GetPaymentHistory = paymentHistoryResult
End Function

' 환불 신청 (Refund)
Public Function Refund(CorpNum, RefundForm,  UserID)
    Dim tmp: Set tmp = RefundForm.toJsonInfo
    Dim postData: postData = m_Linkhub.toString(tmp)

    response.write(postdata)

    Dim tmpResult:Set tmpResult = httpPOST("/Refund", getSession_token(CorpNum), "", postData, UserID)

    Dim refundResponse: Set refundResponse = New RefundResponse
    refundResponse.fromJsonInfo tmpResult
    Set Refund = refundResponse
End Function

' 환불 신청내역 (GetRefundHistory)
Public Function GetRefundHistory(CorpNum, Page, PerPage, UserID)
    Dim tmp : Set tmp  = httpGET("/RefundHistory?Page="&Page & "&PerPage="&PerPage,getSession_token(CorpNum),UserID)

    Dim refundHistoryResult : Set refundHistoryResult = New RefundHistoryResult
    refundHistoryResult.fromJsonInfo tmp

    Set GetRefundHistory = refundHistoryResult
End Function

' 환불 신청상태 확인 (GetRefundInfo)
Public Function GetRefundInfo(CorpNum, RefundCode, UserID)
    If RefundCode = "" then
        Err.Raise -99999999, "POPBILL", "환불코드가 입력되지 않았습니다."
    End If

	Set tmp = httpGET("/Refund/"&RefundCode,getSession_token(CorpNum),UserID)

    Dim refundHistory : Set refundHistory = New RefundHistory
    refundHistory.fromJsonInfo tmp

    Set GetRefundInfo = refundHistory

End Function

' 환불 가능 포인트 조회 (GetRefundableBalance)
Public Function GetRefundableBalance(CorpNum, UserID)
	Dim m_balance: Set m_balance = httpGET("/RefundPoint",getSession_token(CorpNum),UserID)
	GetRefundableBalance = CDbl(m_balance.refundableBalance)
End Function

' 팝빌회원 탈퇴 (QuitMember)
Public Function QuitMember(CorpNum, QuitReason, UserID)
    Dim t_QuitReason: Set t_QuitReason = QuitReason.toJsonInfo
    Dim postdata: postdata = m_Linkhub.toString(t_QuitReason)
    Dim tmp: Set tmp  = httpPOST("/QuitMember", getSession_token(CorpNum), "", postData, UserID)
    Set QuitMember = tmp
End Function

'''''''''''''  End of PopbillBase

'Private Functions
Public Function httpGET(url , BearerToken , UserID )

    Dim winhttp1: Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("GET", getTargetURL() + url, false)

    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP POPBILL SDK")
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)

    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If

    winhttp1.Send
    winhttp1.WaitForResponse
    Dim result : result = winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
        Dim parsedDic : Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If

    Set winhttp1 = Nothing

    Set httpGET = m_Linkhub.parse(result)

End Function


Public Function httpPOST(url , BearerToken , override , postdata ,  UserID)

    Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("POST", getTargetURL() + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP POPBILL SDK")

    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If

    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If

    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If

    winhttp1.Send (postdata)
    winhttp1.WaitForResponse
    Dim result : result = winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
        Dim parsedDic : Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If

    Set winhttp1 = Nothing
    Set httpPOST = m_Linkhub.parse(result)

End Function

Public Function httpBulkPOST(url, BearerToken, override, SubmitID, postdata, userID)

    Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("POST", getTargetURL() + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP POPBILL SDK")
    Call winhttp1.setRequestHeader("x-pb-message-digest", m_linkhub.b64sha1(postdata))
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If

    If SubmitID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-submit-id", SubmitID)
    End If

    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If

    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If

    winhttp1.Send (postdata)
    winhttp1.WaitForResponse
    Dim result : result = winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
        Dim parsedDic : Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If

    Set winhttp1 = Nothing
    Set httpBulkPOST = m_Linkhub.parse(result)

End Function

Public Function httpPOST_ContentsType(url , BearerToken , override , postdata , UserID, ContentsType)
    Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("POST", getTargetURL() + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP POPBILL SDK")

    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If

    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If

    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If

    If ContentsType <> "" Then
        Call winhttp1.setRequestHeader("Content-Type", ContentsType)
    Else
        Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    End If

    winhttp1.Send (postdata)
    winhttp1.WaitForResponse
    Dim result : result = winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
        Dim parsedDic : Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If

    Set winhttp1 = Nothing
    Set httpPOST_ContentsType = m_Linkhub.parse(result)
End Function



Public Function httpPOST_File(url , BearerToken , FilePath , UserID )

    Dim boundary : boundary = "---------------------popbill"

    Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("POST", getTargetURL() + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP POPBILL SDK")

    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If

    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If

    Call winhttp1.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary)

    Dim Stream : Set Stream = Server.CreateObject("ADODB.Stream")
    Stream.Type = adTypeBinary
    Stream.Open

    Dim fileHead : fileHead = vbCrLf & "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""Filedata""; filename=""" & GetOnlyFileName(FilePath) & """" + vbCrLf & _
           "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
    Stream.Write StringToBytes(fileHead)
    Stream.Write GetFile(FilePath)

    Dim tail : tail = vbCrLf & "--" & boundary & "--" & vbCrLf
    Stream.Write  StringToBytes(tail)

    Stream.Flush
    Stream.Position = 0
    Dim postData : postData = Stream.Read
    Stream.Close : Set Stream = Nothing

    winhttp1.Send (postData)
    winhttp1.WaitForResponse

    Dim result : result = winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
        Dim parsedDic : Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If

    Set winhttp1 = Nothing

    Set httpPOST_File = m_Linkhub.parse(Result)

End Function


Public Function httpPOST_Files(url , BearerToken ,postData, FilePaths , UserID )

    Dim boundary : boundary = "---------------------popbill"

    Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("POST", getTargetURL() + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP POPBILL SDK")

    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If

    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If

    Call winhttp1.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary)

    Dim Stream : Set Stream = Server.CreateObject("ADODB.Stream")
    Stream.Type = adTypeBinary
    Stream.Open

    If postdata <> "" Then
        Dim applicationform : applicationform = vbCrLf & "--" & boundary & vbCrLf & _
                          "Content-Disposition: form-data; name=""form""" & vbCrLf & _
                          "Content-Type: Application/json" & vbCrLf & vbCrLf & _
                          postdata
        Stream.Write StringToBytes(applicationform)
    End If

    Dim FilePath
    For Each FilePath In FilePaths
        Dim fileHead : fileHead = vbCrLf & "--" & boundary & vbCrLf & _
               "Content-Disposition: form-data; name=""file""; filename=""" & GetOnlyFileName(FilePath) & """" + vbCrLf & _
               "Content-Type: application/octet-stream" & vbCrLf & vbCrLf

        Stream.Write StringToBytes(fileHead)
        Stream.Write GetFile(FilePath)
    Next

    Dim tail : tail = vbCrLf & "--" & boundary & "--" & vbCrLf
    Stream.Write  StringToBytes(tail)

    Stream.Flush
    Stream.Position = 0
    Dim btPostData : btPostData = Stream.Read
    Stream.Close : Set Stream = Nothing

    winhttp1.Send (btPostData)
    winhttp1.WaitForResponse

    Dim result : result = winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
        Dim parsedDic : Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If

    Set winhttp1 = Nothing

    Set httpPOST_Files = m_Linkhub.parse(Result)

End Function

Private Function getTargetURL()
    If m_UseGAIP Then
        getTargetURL = IIf(m_IsTest, ServiceURL_GA_TEST, ServiceURL_GA_REAL)
    ElseIf m_UseStaticIP Then
        getTargetURL = IIf(m_IsTest, ServiceURL_Static_TEST, ServiceURL_Static_REAL)
    Else
        getTargetURL = IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL)
    End If
End Function

Private Function StringToBytes(Str)
  Dim Stream
  Set Stream = Server.CreateObject("ADODB.Stream")
  Stream.Type = adTypeText
  Stream.Charset = "UTF-8"
  Stream.Open
  Stream.WriteText Str
  Stream.Flush
  Stream.Position = 0
  Stream.Type = adTypeBinary
  Dim buffer : buffer= Stream.Read
  Stream.Close
  'Remove BOM.
  Set Stream = Server.CreateObject("ADODB.Stream")
  Stream.Type = adTypeBinary
  Stream.Open
  Stream.write buffer
  Stream.Flush
  Stream.Position = 3
  StringToBytes= Stream.Read
  Stream.Close
  Set Stream = Nothing

End Function

Private Function GetFile(FileName)
    Dim Stream : Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = adTypeBinary
    Stream.Open
    Stream.LoadFromFile FileName
    GetFile = Stream.Read
    Stream.Close
End Function

Private Function GetOnlyFileName(ByVal FilePath )
     Dim Temp : Temp = Split(FilePath, "\")
     GetOnlyFileName = Split(FilePath, "\")(UBound(Temp))
End Function

Private Function IIf(condition , trueState,falseState)
    If condition Then
        IIf = trueState
    Else
        IIf = falseState
    End if
End Function
public Function toString(object)
    toString = m_Linkhub.toString(object)
End Function

Public Function parse(jsonString)
    Set parse = m_Linkhub.parse(jsonString)
End Function
End Class

'회원가입 정보
Class JoinForm
    Public LinkID
    Public CorpNum
    Public CEOName
    Public CorpName
    Public Addr
    Public ZipCode
    Public BizType
    Public BizClass
    Public ID
    Public PWD
    Public Password
    Public ContactName
    Public ContactTEL
    Public ContactHP
    Public ContactFAX
    Public ContactEmail
End Class

'담당자 정보
Class ContactInfo
    Public id
    Public pwd
    Public Password
    Public email
    Public hp
    Public personName
    Public searchAllAllowYN
    Public searchRole
    Public tel
    Public fax
    Public mgrYN
    Public regDT
    Public state

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next

        id = jsonInfo.id
        email = jsonInfo.email
        hp = jsonInfo.hp
        personName = jsonInfo.personName
        searchAllAllowYN = jsonInfo.searchAllAllowYN
        searchRole = jsonInfo.searchRole
        tel = jsonInfo.tel
        fax = jsonInfo.fax
        mgrYN = jsonInfo.mgrYN
        regDT = jsonInfo.regDT
        State = jsonInfo.state

        On Error GoTo 0
    End Sub

    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.set "id", id
        toJsonInfo.set "pwd", pwd
        toJsonInfo.set "Password", Password
        toJsonInfo.set "email", email
        toJsonInfo.set "hp", hp
        toJsonInfo.set "personName", personName
        toJsonInfo.set "searchAllAllowYN", searchAllAllowYN
        toJsonInfo.set "searchRole", searchRole
        toJsonInfo.set "tel", tel
        toJsonInfo.set "fax", fax
        toJsonInfo.set "mgrYN", mgrYN
    End Function

End Class

'회사정보
Class CorpInfo
    Public ceoname
    Public corpName
    Public addr
    Public bizType
    Public bizClass

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
        ceoname = jsonInfo.ceoname
        corpName = jsonInfo.corpName
        addr = jsonInfo.addr
        bizType = jsonInfo.bizType
        bizClass = jsonInfo.bizClass
        On Error GoTo 0
    End Sub

    Public Function toJsonINfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.Set "ceoname", ceoname
        toJsonInfo.Set "corpName", corpName
        toJsonInfo.Set "addr", addr
        toJsonInfo.Set "bizType", bizType
        toJsonInfo.Set "bizClass", bizClass
    End Function

End Class

'과금정보
Class ChargeInfo
    Public unitCost
    Public chargeMethod
    Public rateSystem

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
        unitCost = jsonInfo.unitCost
        chargeMethod = jsonInfo.chargeMethod
        rateSystem = jsonInfo.rateSystem
        On Error GoTo 0
    End Sub

End Class

'무통장 입금신청 객체정보
Class PaymentForm
    Public SettlerName
    Public SettlerEmail
    Public NotifyHP
    Public PaymentName
    Public SettleCost

    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.set "settlerName", SettlerName
        toJsonInfo.set "settlerEmail", SettlerEmail
        toJsonInfo.set "notifyHP", NotifyHP
        toJsonInfo.set "paymentName", PaymentName
        toJsonInfo.set "settleCost", SettleCost
    End Function

End Class

' 환불신청 객체정보
Class RefundForm
    Public ContactName
    Public TEL
    Public RequestPoint
    Public AccountBank
    Public AccountNum
    Public AccountName
    Public Reason

    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.set "ContactName", ContactName
        toJsonInfo.set "TEL", TEL
        toJsonInfo.set "RequestPoint", RequestPoint
        toJsonInfo.set "AccountBank", AccountBank
        toJsonInfo.set "AccountNum", AccountNum
        toJsonInfo.set "AccountName", AccountName
        toJsonInfo.set "Reason", Reason
    End Function

End Class

'회원탈퇴 신청 객체정보
Class QuitReason
    public quitReason

    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.set "QuitReason", quitReason
    End Function

End Class

Class UseHistoryResult
Public code
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
        total = jsonInfo.total
        perPage = jsonInfo.perPage
        pageNum = jsonInfo.pageNum
        pageCount = jsonInfo.pageCount

        ReDim list(jsonInfo.list.length)
        Dim i
        For i = 0 To jsonInfo.list.length -1
            Dim tmpObj : Set tmpObj = New UseHistory
            tmpObj.fromJsonInfo jsonInfo.list.Get(i)
            Set list(i) = tmpObj
        Next

        On Error GoTo 0
    End Sub
End Class

Class UseHistory
    Public itemCode
    Public txType
    Public txPoint
    Public balance
    Public txDT
    Public userID
    Public userName

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
            itemCode = jsonInfo.itemCode
            txType = jsonInfo.txType
            txPoint = jsonInfo.txPoint
            balance = jsonInfo.balance
            txDT = jsonInfo.txDT
            userID = jsonInfo.userID
            userName = jsonInfo.userName
        On Error GoTo 0
    End Sub
End Class

Class RefundHistory

    Public reqDT
    Public requestPoint
    Public accountBank
    Public accountNum
    Public accountName
    Public state
    Public reason

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
            reqDT = jsonInfo.reqDT
            requestPoint = jsonInfo.requestPoint
            accountBank = jsonInfo.accountBank
            accountNum = jsonInfo.accountNum
            accountName = jsonInfo.accountName
            state = jsonInfo.state
            reason = jsonInfo.reason
        On Error GoTo 0
    End Sub

End Class

Class RefundHistoryResult
Public code
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
        total = jsonInfo.total
        perPage = jsonInfo.perPage
        pageNum = jsonInfo.pageNum
        pageCount = jsonInfo.pageCount

        ReDim list(jsonInfo.list.length)
        Dim i
        For i = 0 To jsonInfo.list.length -1
            Dim tmpObj : Set tmpObj = New RefundHistory
            tmpObj.fromJsonInfo jsonInfo.list.Get(i)
            Set list(i) = tmpObj
        Next

        On Error GoTo 0
    End Sub
End Class

Class PaymentHistory
    Public productType
    Public productName
    Public settleType
    Public settlerName
    Public settlerEmail
    Public settleCost
    Public settlePoint
    Public settleState
    Public regDT
    Public stateDT

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
            productType = jsonInfo.productType
            productName = jsonInfo.productName
            settleType = jsonInfo.settleType
            settlerName = jsonInfo.settlerName
            settlerEmail = jsonInfo.settlerEmail
            settleCost = jsonInfo.settleCost
            settlePoint = jsonInfo.settlePoint
            settleState = jsonInfo.settleState
            regDT = jsonInfo.regDT
            stateDT = jsonInfo.stateDT
        On Error GoTo 0
    End Sub

End Class


Class PaymentHistoryResult
Public code
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
        total = jsonInfo.total
        perPage = jsonInfo.perPage
        pageNum = jsonInfo.pageNum
        pageCount = jsonInfo.pageCount

        ReDim list(jsonInfo.list.length)
        Dim i
        For i = 0 To jsonInfo.list.length -1
            Dim tmpObj : Set tmpObj = New PaymentHistory
            tmpObj.fromJsonInfo jsonInfo.list.Get(i)
            Set list(i) = tmpObj
        Next

        On Error GoTo 0
    End Sub
End Class


Class RefundResponse
	Public code
    Public message
    Public refundCode

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next

        code = jsonInfo.code
        message = jsonInfo.message
        refundCode = jsonInfo.refundCode

        On Error GoTo 0
    End Sub
End Class

Class PaymentResponse
    Public code
    Public message
    Public settleCode

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next

        code = jsonInfo.code
        message = jsonInfo.message
        settleCode = jsonInfo.settleCode

        On Error GoTo 0
    End Sub
End Class
%>