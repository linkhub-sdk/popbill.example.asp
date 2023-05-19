<%
Const SELL = "SELL"
Const BUY = "BUY"
Const TRUSTEE = "TRUSTEE"

Class HTTaxinvoiceService

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
        m_PopbillBase.AddScope("111")
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

    '파트너 포인트 충전 팝업 URL - 2017/08/29 추가
    Public Function GetPartnerURL(CorpNum, TOGO)
        GetPartnerURL = m_PopbillBase.GetPartnerURL(CorpNum,TOGO)
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
        Dim result : Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)

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

    '수집요청
    Public Function RequestJob(CorpNum , KeyType, DType, SDate, Edate, UserID)
        Dim uri
        uri = "/HomeTax/Taxinvoice/" & KeyType
        uri = uri + "?DType=" & DType
        uri = uri + "&SDate=" & SDate
        uri = uri + "&EDate=" & EDate
        Dim result : Set result = m_PopbillBase.httpPOST( uri, m_PopbillBase.getSession_token(CorpNum),"", "", UserID )

        RequestJob = result.jobID
    End Function

    '수집 상태 확인
    Public Function GetJobState(CorpNum, JobID, UserID)
        If Len(JobID) <> 18  Then
            Err.Raise -99999999, "POPBILL", "작업아이디가 올바르지 않습니다."
        End If

        Dim result : Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/" & JobID & "/State", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim jobInfo : Set jobInfo = New HTTIJobState
        jobInfo.fromJsonInfo result
        Set GetJobState = jobInfo
    End Function

    '수집 상태 목록 확인
    Public Function ListActiveJob(CorpNum, UserID)
        Dim result : Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/JobList", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim jobList : Set jobList = CreateObject("Scripting.Dictionary")

        Dim i
        For i=0 To result.length-1
            Dim jobInfo : Set jobInfo = New HTTIJobState
            jobInfo.fromJsonInfo result.Get(i)
            jobList.Add i, jobInfo
        Next

        Set ListActiveJob = jobList
    End Function

    '수집 결과 조회
    Public Function Search ( CorpNum, JobID, TIType, TaxType, PurposeType, TaxRegIDYN, TaxRegIDType, TaxRegID, Page, PerPage, Order, UserID, SearchString )
        If  Not ( Len ( JobID ) = 18 )  Then
            Err.Raise -99999999, "POPBILL", "작업아이디가 올바르지 않습니다."
        End If
        Dim uri
        uri = "/HomeTax/Taxinvoice/" & JobID

        Dim i
        uri = uri & "?Type="
        For i = 0 To UBound(TIType) -1
            If i = UBound(TIType) -1 Then
                uri = uri & TIType(i)
            Else
                uri = uri & TIType(i) & ","
            End if
        Next

        uri = uri & "&TaxType="
        For i = 0 To UBound(TaxType) -1
            If i = UBound(TaxType) -1 Then
                uri = uri & TaxType(i)
            Else
                uri = uri & TaxType(i) & ","
            End if
        Next

        uri = uri & "&PurposeType="
        For i = 0 To UBound(PurposeType) -1
            If i = UBound(PurposeType) -1 Then
                uri = uri & PurposeType(i)
            Else
                uri = uri & PurposeType(i) & ","
            End if
        Next

        If TaxRegIDYN <> "" Then
            uri = uri & "&TaxRegIDYN=" & TaxRegIDYN
        End If

        uri = uri & "&TaxRegIDType=" & TaxRegIDType

        uri = uri & "&TaxRegID=" & TaxRegID

        If SearchString <> "" Then
            uri = uri & "&SearchString=" & Server.URLEncode(SearchString)
        End If

        uri = uri & "&Page=" & CStr(Page)
        uri = uri & "&PerPage=" & CStr(PerPage)
        uri = uri & "&Order=" & Order

        Dim result : Set result = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim searchResult : Set searchResult = New HTTaxinvoiceSerach
        searchResult.fromJsonInfo result
        Set Search = searchResult

    End Function

    '수집 결과 요약정보 조회
    Public Function Summary ( CorpNum, JobID, TIType, TaxType, PurposeType, TaxRegIDYN, TaxRegIDType, TaxRegID, UserID, SearchString )
        If Not ( Len ( JobID ) = 18 ) Then
            Err.Raise -99999999, "POPBILL", "작업아이디가 올바르지 않습니다."
        End If
        Dim uri
        uri = "/HomeTax/Taxinvoice/" & JobID & "/Summary"

        Dim i
        uri = uri & "?Type="
        For i = 0 To UBound(TIType) -1
            If i = UBound(TIType) -1 Then
                uri = uri & TIType(i)
            Else
                uri = uri & TIType(i) & ","
            End if
        Next

        uri = uri & "&TaxType="
        For i = 0 To UBound(TaxType) -1
            If i = UBound(TaxType) -1 Then
                uri = uri & TaxType(i)
            Else
                uri = uri & TaxType(i) & ","
            End if
        Next

        uri = uri & "&PurposeType="
        For i = 0 To UBound(PurposeType) -1
            If i = UBound(PurposeType) -1 Then
                uri = uri & PurposeType(i)
            Else
                uri = uri & PurposeType(i) & ","
            End if
        Next

        uri = uri & "&TaxRegIDType=" & TaxRegIDType

        If TaxRegIDYN <> "" Then
            uri = uri & "&TaxRegIDYN=" & TaxRegIDYN
        End If

        uri = uri & "&TaxRegID=" & TaxRegID

        If SearchString <> "" Then
            uri = uri & "&SearchString=" & Server.URLEncode(SearchString)
        End If

        Dim result : Set result = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim summaryResult : Set summaryResult = New HTTaxinvoiceSummary
        summaryResult.fromJsonInfo result
        Set Summary = summaryResult

    End Function

    '상세정보 조회 - JSON
    Public Function GetTaxinvoice ( CorpNum, NTSConfirmNum, UserID )
        If Not ( Len ( NTSConfirmNum ) = 24 ) Then
            Err.Raise -99999999, "POPBILL", "국세청승인번호가 올바르지 않습니다."
        End If

        Dim responseObj : Set responseObj = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/" & NTSConfirmNUm, _
                                m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim taxinvoiceDetail : Set taxinvoiceDetail = New HTTaxinvoice
        taxinvoiceDetail.fromJsonInfo responseObj
        Set GetTaxinvoice = taxinvoiceDetail
    End Function

    '상세정보 조회 - XML
    Public Function GetXML ( CorpNum, NTSConfirmNum, UserID )
        If Not ( Len ( NTSConfirmNum ) = 24 ) Then
            Err.Raise -99999999, "POPBILL", "국세청승인번호가 올바르지 않습니다."
        End If

        Dim responseObj : Set responseObj = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/" & NTSConfirmNum & "?T=xml", _
                                m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim taxinvoiceXML : Set taxinvoiceXML = New HTTaxinvoiceXML
        taxinvoiceXML.fromJsonInfo responseObj
        Set GetXML = taxinvoiceXML
    End Function

    '정액제 신청 팝업 URL
    Public Function GetFlatRatePopUpURL ( CorpNum, UserID )
        Dim result : Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice?TG=CHRG", _
                            m_PopbillBase.getSession_token(CorpNum), UserID)
        GetFlatRatePopUpURL = result.url
    End Function

    '정액제 상태 확인
    Public Function GetFlatRateState ( CorpNum, UserID )
        Dim responseObj : Set responseObj = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/Contract", _
                            m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim flatRateObj : Set flatRateObj = New HTTIFlatRate
        flatRateObj.fromJsonInfo responseObj
        Set GetFlatRateState = flatrateObj
    End Function

    '공인인증서 등록 URL
    Public Function GetCertificatePopUpURL ( CorpNum, UserID )
        Dim result : Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice?TG=CERT", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)
        GetCertificatePopUpURL = result.url
    End Function

    '공인인증서 만료일자 확인
    Public Function GetCertificateExpireDate ( CorpNum, UserID )
        Dim result : Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/CertInfo", _
                    m_PopbillBase.getSession_token(CorpNum), UserID)
        GetCertificateExpireDate = result.certificateExpiration
    End Function

    '홈택스 전자세금계산서 보기 팝업 URL
    Public Function GetPopUpURL ( CorpNum, NTSConfirmNum, UserID )
        If Not ( Len ( NTSConfirmNum ) = 24 ) Then
            Err.Raise -99999999, "POPBILL", "국세청승인번호가 올바르지 않습니다."
        End If

        Dim result : Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/" & NTSConfirmNum & "/PopUp", _
                                m_PopbillBase.getSession_token(CorpNum), UserID)

        GetPopUpURL = result.url
    End Function

    '홈택스 전자세금계산서 인쇄 팝업 URL
    Public Function GetPrintURL ( CorpNum, NTSConfirmNum, UserID )
        If Not ( Len ( NTSConfirmNum ) = 24 ) Then
            Err.Raise -99999999, "POPBILL", "국세청승인번호가 올바르지 않습니다."
        End If

        Dim result : Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/" & NTSConfirmNum & "/Print", _
                                m_PopbillBase.getSession_token(CorpNum), UserID)

        GetPrintURL = result.url
    End Function


    '홈택스 공인인증서 로그인 테스트
    Public Function CheckCertValidation ( CorpNum, UserID )
        Set CheckCertValidation = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/CertCheck", m_PopbillBase.getSession_token(CorpNum), UserID)
    End Function

    '부서사용자 계정등록
    Public Function RegistDeptUser ( CorpNum, DeptUserID, DeptUserPWD, UserID )
        If DeptUserID = "" Then
            Err.Raise -99999999, "POPBILL", "홈택스 부서사용자 계정 아이디가 입력되지 않았습니다."
        End If
        If DeptUserPWD = "" Then
            Err.Raise -99999999, "POPBILL", "홈택스 부서사용자 계정 비밀번호가 입력되지 않았습니다."
        End If

        Dim tmp : Set tmp = JSON.parse("{}")
        tmp.Set "id", DeptUserID
        tmp.Set "pwd", DeptUserPWD

        Dim postData : postData = m_PopbillBase.toString(tmp)

        Set RegistDeptUser = m_PopbillBase.httpPOST("/HomeTax/Taxinvoice/DeptUser", m_PopbillBase.getSession_token(CorpNum),"", postData, UserID)
    End Function

    '부서사용자 등록정보 확인
    Public Function CheckDeptUser ( CorpNum, UserID )
        Set CheckDeptUser = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/DeptUser", m_PopbillBase.getSession_token(CorpNum), UserID)
    End Function

    '부서사용자 로그인 테스트
    Public Function CheckLoginDeptUser ( CorpNum, UserID )
        Set CheckLoginDeptUser = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/DeptUser/Check", m_PopbillBase.getSession_token(CorpNum), UserID)
    End Function

    '부서사용자 등록정보 삭제
    Public Function DeleteDeptUser ( CorpNum, UserID )
        Set DeleteDeptUser = m_PopbillBase.httpPOST("/HomeTax/Taxinvoice/DeptUser", m_PopbillBase.getSession_token(CorpNum),"DELETE", "", UserID)
    End Function


''End Of Class HTTaxinvoiceService
End Class

Class HTTIFlatRate
    Public referenceID
    Public contractDT
    Public useEndDate
    Public baseDate
    Public state
    Public closeRequestYN
    Public useRestrictYN
    Public closeOnExpired
    Public unPaidYN

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
            referenceID = jsonInfo.referenceID
            contractDT = jsonInfo.contractDT
            useEndDate = jsonInfo.useEndDate
            baseDate = jsonInfo.baseDate
            state = jsonInfo.state
            closeRequestYN = jsonInfo.closeRequestYN
            useRestrictYN = jsonInfo.useRestrictYN
            closeOnExpired = jsonInfo.closeOnExpired
            unPaidYN = jsonInfo.unPaidYN
        On Error GoTo 0
    End Sub
End class

Class HTTaxinvoiceXML
    Public ResultCode
    Public Message
    Public retObject

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
            ResultCode = jsonInfo.ResultCode
            Message = jsonInfo.Message
            retObject = jsonInfo.retObject
        On Error GoTo 0
    End Sub
End Class

Class HTTaxinvoice
    Public writeDate
    Public issueDT
    Public invoiceType
    Public taxType
    Public taxTotal
    Public supplyCostTotal
    Public totalAmount
    Public purposeType
    Public serialNum
    Public cash
    Public chkBill
    Public credit
    Public note
    Public remark1
    Public remark2
    Public remark3
    Public ntsconfirmNum
    Public modifyCode
    Public orgNTSConfirmNum
    Public invoicerCorpNum
    Public invoicerMgtKey
    Public invoicerTaxRegID
    Public invoicerCorpName
    Public invoicerCEOName
    Public invoicerAddr
    Public invoicerBizType
    Public invoicerBizClass
    Public invoicerContactName
    Public invoicerDeptName
    Public invoicerTEL
    Public invoicerEmail

    Public invoiceeCorpNum
    Public invoiceeType
    Public invoiceeMgtKey
    Public invoiceeTaxRegID
    Public invoiceeCorpName
    Public invoiceeCEOName
    Public invoiceeAddr
    Public invoiceeBizType
    Public invoiceeBizClass
    Public invoiceeContactName1
    Public invoiceeDeptName1
    Public invoiceeTEL1
    Public invoiceeEmail1
    Public invoiceeContactName2
    Public invoiceeTEL2
    Public invoiceeEmail2

    Public trusteeCorpNum
    Public trusteeMgtKey
    Public trusteeTaxRegID
    Public trusteeCorpName
    Public trusteeCEOName
    Public trusteeAddr
    Public trusteeBizType
    Public trusteeBizClass
    Public trusteeContactName
    Public trusteeDeptName
    Public trusteeTEL
    Public trusteeEmail

    Public detailList()

    Public Sub Class_Initialize
        ReDim detailList(-1)
    End Sub



    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
            writeDate = jsonInfo.writeDate
            issueDT = jsonInfo.issueDT
            invoiceType = jsonInfo.invoiceType
            taxType = jsonInfo.taxType
            taxTotal = jsonInfo.taxTotal
            supplyCostTotal = jsonInfo.supplyCostTotal
            totalAmount = jsonInfo.totalAmount
            purposeType = jsonInfo.purposeType
            serialNum = jsonInfo.serialNum
            cash = jsonInfo.cash
            chkBill = jsonInfo.chkBill
            credit = jsonInfo.credit
            note = jsonInfo.note
            remark1 = jsonInfo.remark1
            remark2 = jsonInfo.remark2
            remark3 = jsonInfo.remark3
            ntsconfirmNum = jsonInfo.ntsconfirmNum

            modifyCode = jsonInfo.modifyCode
            orgNTSConfirmNum = jsonInfo.orgNTSConfirmNum

            invoicerCorpNum = jsonInfo.invoicerCorpNum
            invoicerMgtKey = jsonInfo.invoicerMgtKey
            invoicerTaxRegID = jsonInfo.invoicerTaxRegID
            invoicerCorpName = jsonInfo.invoicerCorpName
            invoicerCEOName = jsonInfo.invoicerCEOName
            invoicerAddr = jsonInfo.invoicerAddr
            invoicerBizType = jsonInfo.invoicerBizType
            invoicerBizClass = jsonInfo.invoicerBizClass
            invoicerContactName = jsonInfo.invoicerContactName
            invoicerTEL = jsonInfo.invoicerTEL
            invoicerEmail = jsonInfo.invoicerEmail

            invoiceeCorpNum = jsonInfo.invoiceeCorpNum
            invoiceeType = jsonInfo.invoiceeType
            invoiceeMgtKey = jsonInfo.invoiceeMgtKey
            invoiceeTaxRegID = jsonInfo.invoiceeTaxRegID
            invoiceeCorpName = jsonInfo.invoiceeCorpName
            invoiceeCEOName = jsonInfo.invoiceeCEOName
            invoiceeAddr = jsonInfo.invoiceeAddr
            invoiceeBizType = jsonInfo.invoiceeBizType
            invoiceeBizClass = jsonInfo.invoiceeBizClass
            invoiceeContactName1 = jsonInfo.invoiceeContactName1
            invoiceeTEL1 = jsonInfo.invoiceeTEL1
            invoiceeEmail1 = jsonInfo.invoiceeEmail1
            invoiceeContactName2 = jsonInfo.invoiceeContactName2
            invoiceeTEL2 = jsonInfo.invoiceeTEL2
            invoiceeEmail2 = jsonInfo.invoiceeEmail2

            trusteeCorpNum = jsonInfo.trusteeCorpNum
            trusteeMgtKey = jsonInfo.trusteeMgtKey
            trusteeTaxRegID = jsonInfo.trusteeTaxRegID
            trusteeCorpName = jsonInfo.trusteeCorpName
            trusteeCEOName = jsonInfo.trusteeCEOName
            trusteeAddr = jsonInfo.trusteeAddr
            trusteeBizType = jsonInfo.trusteeBizType
            trusteeBizClass = jsonInfo.trusteeBizClass
            trusteeContactName = jsonInfo.trusteeContactName
            trusteeDeptName = jsonInfo.trusteeDeptName
            trusteeTEL = jsonInfo.trusteeTEL
            trusteeEmail = jsonInfo.trusteeEmail

            ReDim detailList(jsonInfo.detailList.length)
            Dim i
            For i = 0 To jsonInfo.detailList.length -1
                Dim tmpDetail : Set tmpDetail = New HTTaxinvoiceDetail
                tmpDetail.FromJsonInfo jsonInfo.detailList.Get(i)
                Set detailList(i) = tmpDetail
            Next

        On Error GoTo 0
    End Sub
End Class

Class HTTaxinvoiceDetail
    Public serialNum
    Public purchaseDT
    Public itemName
    Public spec
    Public qty
    Public unitCost
    Public supplyCost
    Public tax
    Public remark

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
        serialNum = jsonInfo.serialNum
        purchaseDT = jsonInfo.purchaseDT
        itemName = jsonInfo.itemName
        spec = jsonInfo.spec
        qty = jsonInfo.qty
        unitCost = jsonInfo.unitCost
        supplyCost = jsonInfo.supplyCost
        tax = jsonInfo.tax
        remark = jsonInfo.remark
        On Error GoTo 0
    End Sub

End Class


Class HTTaxinvoiceSummary
    Public count
    Public supplyCostTotal
    Public taxTotal
    Public amountTotal

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
        count = jsonInfo.count
        supplyCostTotal = jsonInfo.supplyCostTotal
        taxTotal = jsonInfo.taxTotal
        amountTotal = jsonInfo.amountTotal
        On Error GoTo 0
    End Sub
End Class

Class HTTaxinvoiceSerach
    Public code
    Public message
    Public total
    Public perPage
    Public pageNum
    Public pageCount
    Public list()

    Public Sub classs_initialize
        ReDim list(-1)
    End Sub

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
        code = jsonInfo.code
        message = jsonInfo.message
        total = jsonInfo.total
        perPage = jsonInfo.perPage
        pageNum = jsonInfo.pageNum
        pageCount = jsonInfo.pageCount

        ReDim list ( jsonInfo.list.length )
        Dim i
        For i = 0 To jsonInfo.list.length -1
            Dim tmpObj : Set tmpObj = New HTTaxinvoiceAbbr
            tmpObj.fromJsonInfo jsonInfo.list.Get(i)
            Set list(i) = tmpObj
        next

        On Error GoTo 0
    End Sub
End Class

Class HTTaxinvoiceAbbr
    Public ntsconfirmNum
    Public writeDate
    Public issueDate
    Public sendDate
    Public taxType
    Public purposeType
    Public supplyCostTotal
    Public taxTotal
    Public totalAmount
    Public remark1
    Public modifyYN
    Public orgNTSConfirmNum

    Public purchaseDate
    Public itemName
    Public spec
    Public qty
    Public unitCost
    Public supplyCost
    Public tax
    Public remark

    Public invoicerCorpNum
    Public invoicerTaxRegID
    Public invoicerCorpName
    Public invoicerCEOName
    Public invoicerEmail

    Public invoiceeCorpNum
    Public invoiceeType
    Public invoiceeTaxRegID
    Public invoiceeCorpName
    Public invoiceeCEOName
    Public invoiceeEmail1
    Public invoiceeEmail2

    Public trusteeCorpNum
    Public trusteeTaxRegID
    Public trusteeCorpName
    Public trusteeCEOName
    Public trusteeEmail

    '매입/매출 구분 필드 추가 (2017/08/29)
    Public invoiceType

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
        ntsconfirmNum = jsonInfo.ntsconfirmNum
        writeDate = jsonInfo.writeDate
        issueDate = jsonInfo.issueDate
        sendDate = jsonInfo.sendDate
        taxType = jsonInfo.taxType
        purposeType = jsonInfo.purposeType
        supplyCostTotal = jsonInfo.supplyCostTotal
        taxTotal = jsonInfo.taxTotal
        totalAmount = jsonInfo.totalAmount
        remark1 = jsonInfo.remark1


        modifyYN = jsonInfo.modifyYN
        orgNTSConfirmNum = jsonInfo.orgNTSConfirmNUm

        purchaseDate = jsonInfo.purchaseDate
        itemName = jsonInfo.itemName
        spec = jsonInfo.spec
        qty = jsonInfo.qty
        unitCost = jsonInfo.unitCost
        supplyCost = jsonInfo.supplyCost
        tax = jsonInfo.taxt
        remark = jsonInfo.remark

        invoicerCorpNum = jsonInfo.invoicerCorpNum
        invoicerTaxRegID = jsonInfo.invoicerTaxRegID
        invoicerCorpName = jsonInfo.invoicerCorpName
        invoicerCEOName = jsonInfo.invoicerCEOName
        invoicerEmail = jsonInfo.invoicerEmail

        invoiceeCorpNum = jsonInfo.invoiceeCorpNum
        invoiceeType = jsonInfo.invoiceeType
        invoiceeTaxRegID = jsonInfo.invoiceeTaxRegID
        invoiceeCorpName = jsonInfo.invoiceeCorpName
        invoiceeCEOName = jsonInfo.invoiceeCEOName
        invoiceeEmail1 = jsonInfo.invoiceeEmail1
        invoiceeEmail2 = jsonInfo.invoiceeEmail2

        trusteeCorpNum = jsonInfo.trusteeCorpNum
        trusteeTaxRegID = jsonInfo.trusteeTaxRegID
        trusteeCorpName = jsonInfo.trusteeCorpName
        trusteeCEOName = jsonInfo.trusteeCEOName
        trusteeEmail = jsonInfo.trusteeEmail

        invoiceType = jsonInfo.invoiceType
        On Error GoTo 0
    End Sub

End class


Class HTTIJobState
    Public jobID
    Public jobState
    Public queryType
    Public queryDateType
    Public queryStDate
    Public queryEnDate
    Public errorCode
    Public errorReason
    Public jobStartDT
    Public jobEndDT
    Public collectCount
    Public regDT

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
            jobID = jsonInfo.jobID
            jobState = jsonInfo.jobState
            queryType = jsonInfo.queryType
            queryDateType = jsonInfo.queryDateType
            queryStDate = jsonInfo.queryStDate
            queryEnDate = jsonInfo.queryEnDate
            errorCode = jsonInfo.errorCode
            errorReason = jsonInfo.errorReason
            jobStartDT = jsonInfo.jobStartDT
            jobEndDT = jsonInfo.jobEndDT
            collectCount = jsonInfo.collectCount
            regDT = jsonInfo.regDT
        On Error GoTo 0
    End sub
End Class

%>