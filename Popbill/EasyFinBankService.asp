<%

Class EasyFinBankSErvice
    
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
        m_PopbillBase.AddScope("180")
    End Sub
    
    Public Sub Initialize(linkID, SecretKey )
        m_PopbillBase.Initialize linkID,SecretKey
    End Sub

    '회원 포인트조회
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

    '팝빌 연동회원 포인트 충전내역 URL
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
        Dim result : Set result = m_PopbillBase.httpGET("/EasyFin/Bank/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)
        Dim chrgInfo : Set chrgInfo = New ChargeInfo
        chrgInfo.fromJsonInfo result
        
        Set GetChargeInfo = chrgInfo
    End Function 

    '''''''''''''  End of PopbillBase

    
    Public Function RegistBankAccount(CorpNum, BankInfoObj, UserID)
        Dim uri
        uri = "/EasyFin/Bank/BankAccount/Regist"
        uri = uri + "?UsePeriod=" & BankInfoObj.UsePeriod

        Dim tmp : Set tmp = BankInfoObj.toJsonInfo
        Dim postdata : postdata = m_PopbillBase.toString(tmp)

        Set RegistBankAccount = m_PopbillBase.httpPOST(uri, m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)
    End Function

    Public Function UpdateBankAccount(CorpNum, BankInfoObj, UserID)

        Dim uri : uri = "/EasyFin/Bank/BankAccount/"+BankInfoObj.BankCode+"/"+BankInfoObj.AccountNumber+"/Update"

        Dim tmp : Set tmp = BankInfoObj.toJsonInfo
        Dim postdata : postdata = m_PopbillBase.toString(tmp)

        Set UpdateBankAccount = m_PopbillBase.httpPOST(uri, m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)
    End Function


    Public Function CloseBankAccount(CorpNum, BankCode, AccountNumber, CloseType, UserID)
        Dim uri
        uri = "/EasyFin/Bank/BankAccount/Close"
        uri = uri + "?BankCode=" & BankCode
        uri = uri + "&AccountNumber=" & AccountNumber
        uri = uri + "&CloseType=" & CloseType

        Set CloseBankAccount = m_PopbillBase.httpPOST( uri, m_PopbillBase.getSession_token(CorpNum),"", "", UserID )

    End Function
    
    Public Function RevokeCloseBankAccount(CorpNum, BankCode, AccountNumber, UserID)
        Dim uri
        uri = "/EasyFin/Bank/BankAccount/RevokeClose"
        uri = uri + "?BankCode=" & BankCode
        uri = uri + "&AccountNumber=" & AccountNumber

        Set RevokeCloseBankAccount = m_PopbillBase.httpPOST( uri, m_PopbillBase.getSession_token(CorpNum),"", "", UserID )

    End function

    '종량제 계좌 삭제
    Public Function DeleteBankAccount (CorpNum, BankInfoObj, UserID)

         Dim uri : uri = "/EasyFin/Bank/BankAccount/Delete"

        Dim tmp : Set tmp = BankInfoObj.toJsonInfo
        Dim postdata : postdata = m_PopbillBase.toString(tmp)

        Set DeleteBankAccount = m_PopbillBase.httpPOST(uri, m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)
        
    End Function 

    Public Function GetBankAccountMgtURL ( CorpNum, UserID )
        Dim result : Set result = m_PopbillBase.httpGET("/EasyFin/Bank?TG=BankAccount", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)
        GetBankAccountMgtURL = result.url
    End Function


    Public Function GetBankAccountInfo(CorpNum, BankCode, AccountNumber, UserID)

        Dim uri : uri = "/EasyFin/Bank/BankAccount/" & BankCode & "/" & AccountNumber

        Dim result : Set result = m_PopbillBase.httpGET(uri, _
                        m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim infoObj : Set infoObj = New EasyFinBankAccount
        infoObj.fromJsonInfo result
        
        Set GetBankAccountInfo = infoObj	
    End Function 


    Public Function ListBankAccount(CorpNum, UserID)
        Dim result : Set result = m_PopbillBase.httpGET("/EasyFin/Bank/ListBankAccount", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)
        
        Dim bankAccountList : Set bankAccountList = CreateObject("Scripting.Dictionary")
        Dim i
        For i=0 To result.length-1
            Dim tmpInfo : Set tmpInfo = New EasyFinBankAccount
            tmpInfo.fromJsonInfo result.Get(i)
            bankAccountList.Add i, tmpInfo
        Next
        Set ListBankAccount = bankAccountList
    End Function

    Public Function RequestJob(CorpNum , BankCode, AccountNumber, SDate, EDate, UserID)
        Dim uri
        uri = "/EasyFin/Bank/BankAccount"
        uri = uri + "?BankCode=" & BankCode
        uri = uri + "&AccountNumber=" & AccountNumber
        uri = uri + "&SDate=" & SDate
        uri = uri + "&EDate=" & EDate

        Dim result : Set result = m_PopbillBase.httpPOST( uri, m_PopbillBase.getSession_token(CorpNum),"", "", UserID )

        RequestJob = result.jobID
    End Function

    Public Function GetJobState(CorpNum, JobID, UserID)
        If Len(JobID) <> 18  Then
            Err.Raise -99999999, "POPBILL", "작업아이디가 올바르지 않습니다."
        End If

        Dim result : Set result = m_PopbillBase.httpGET("/EasyFin/Bank/" & JobID & "/State", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim jobInfo : Set jobInfo = New EasyFinJobState	
        jobInfo.fromJsonInfo result
        Set GetJobState = jobInfo
    End Function

    Public Function ListActiveJob(CorpNum, UserID)
        Dim result : Set result = m_PopbillBase.httpGET("/EasyFin/Bank/JobList", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)
        
        Dim jobList : Set jobList = CreateObject("Scripting.Dictionary")

        Dim i
        For i=0 To result.length-1
            Dim jobInfo : Set jobInfo = New EasyFinJobState
            jobInfo.fromJsonInfo result.Get(i)
            jobList.Add i, jobInfo
        Next

        Set ListActiveJob = jobList
    End Function

    Public Function Search ( CorpNum, JobID, TradeType, SearchString, Page, PerPage, Order, UserID )

        If  Not ( Len ( JobID ) = 18 )  Then
            Err.Raise -99999999, "POPBILL", "작업아이디가 올바르지 않습니다."
        End If 

        Dim uri
        uri = "/EasyFin/Bank/" & JobID

        Dim i
        uri = uri & "?TradeType="
        For i = 0 To UBound(TradeType) -1 
            If i = UBound(TradeType) -1 Then
                uri = uri & TradeType(i)
            Else
                uri = uri & TradeType(i) & ","
            End if
        Next
        
        If SearchString <> "" Then
            uri = uri & "&SearchString=" & SearchString
        End If 

        uri = uri & "&Page=" & CStr(Page)
        uri = uri & "&PerPage=" & CStr(PerPage)
        uri = uri & "&Order=" & Order

        Dim result : Set result = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim searchResult : Set searchResult = New EasyFinBankSearchResult
        searchResult.fromJsonInfo result
        Set Search = searchResult 

    End Function 

    Public Function Summary ( CorpNum, JobID, TradeType, SearchString, UserID)

        If Not ( Len ( JobID ) = 18 ) Then
            Err.Raise -99999999, "POPBILL", "작업아이디가 올바르지 않습니다."
        End If 

        Dim uri
        uri = "/EasyFin/Bank/" & JobID & "/Summary"

        Dim i
        uri = uri & "?TradeType="
        For i = 0 To UBound(TradeType) -1 
            If i = UBound(TradeType) -1 Then
                uri = uri & TradeType(i)
            Else
                uri = uri & TradeType(i) & ","
            End if
        Next
        
        If SearchString <> "" Then
            uri = uri & "&SearchString=" & SearchString
        End If 

        Dim result : Set result = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), UserID)
    
        Dim summaryResult : Set summaryResult = New EasyFinBankSummaryResult
        summaryResult.fromJsonInfo result
        Set Summary = summaryResult

    End Function

    Public Function SaveMemo(CorpNum , TID, Memo, UserID)

        If TID = "" Then
            Err.Raise -99999999, "POPBILL", "거래내역 아이디가 입력되지 않았습니다."
        End If

        Dim uri
        uri = "/EasyFin/Bank/SaveMemo"
        uri = uri + "?TID=" & TID
        uri = uri + "&Memo=" & Memo
        Set SaveMemo = m_PopbillBase.httpPOST( uri, m_PopbillBase.getSession_token(CorpNum),"", "", UserID )

    End Function

    Public Function GetFlatRatePopUpURL ( CorpNum, UserID )

        Dim result : Set result = m_PopbillBase.httpGET("/EasyFin/Bank?TG=CHRG", m_PopbillBase.getSession_token(CorpNum), UserID)
        GetFlatRatePopUpURL = result.url

    End Function

    Public Function GetFlatRateState ( CorpNum, BankCode, AccountNumber, UserID ) 

        If BankCode = "" Then
            Err.Raise -99999999, "POPBILL", "은행코드가 입력되지 않았습니다."
        End If
        If AccountNumber = "" Then
            Err.Raise -99999999, "POPBILL", "계좌번호가 입력되지 않았습니다."
        End If

        Dim responseObj : Set responseObj = m_PopbillBase.httpGET("/EasyFin/Bank/Contract/" & BankCode & "/" & AccountNumber, _
                        m_PopbillBase.getSession_token(CorpNum), UserID)

        Set flatRateObj = New EasyFinBankFlatRate
        flatRateObj.fromJsonInfo responseObj
        Set GetFlatRateState = flatrateObj
    End Function 

End Class

Class EasyFinBankFlatRate 
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
End Class


Class EasyFinBankSummaryResult
    Public count
    Public cntAccIn
    Public cntAccOut
    Public totalAccIn
    Public totalAccOut
    
    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
        count = jsonInfo.count
        cntAccIn = jsonInfo.cntAccIn
        cntAccOut = jsonInfo.cntAccOut
        totalAccIn = jsonInfo.totalAccIn
        totalAccOut = jsonInfo.totalAccOut
        On Error GoTo 0 
    End Sub 
End Class 


Class EasyFinBankSearchResult
    Public code
    Public message
    Public total
    Public perPage
    Public pageNum
    Public pageCount
    Public lastScrapDT
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
        lastScrapDT = jsonInfo.lastScrapDT
        
        ReDim list ( jsonInfo.list.length )
        Dim i
        For i = 0 To jsonInfo.list.length -1
            Dim tmpObj : Set tmpObj = New EasyFinSearchDetail
            tmpObj.fromJsonInfo jsonInfo.list.Get(i)
            Set list(i) = tmpObj
        next
        
        On Error GoTo 0 
    End Sub 
End Class 

Class EasyFinSearchDetail
    Public tid
    Public trdate
    Public trserial
    Public trdt
    Public accIn
    Public accOut
    Public balance
    Public remark1
    Public remark2
    Public remark3
    Public remark4
    Public regDT
    Public memo
    
    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
        tid = jsonInfo.tid
        trdate = jsonInfo.trdate
        trserial = jsonInfo.trserial
        trdt = jsonInfo.trdt
        accIn = jsonInfo.accIn
        accOut = jsonInfo.accOut
        balance = jsonInfo.balance
        remark1 = jsonInfo.remark1
        remark2 = jsonInfo.remark2
        remark3 = jsonInfo.remark3
        remark4 = jsonInfo.remark4
        regDT = jsonInfo.regDT
        memo = jsonInfo.memo
        On Error GoTo 0
    End Sub 
End class

Class EasyFinJobState

    Public jobID
    Public jobState
    Public startDate
    Public endDate
    Public errorCode
    Public errorReason
    Public jobStartDT
    Public jobEndDT
    Public regDT

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next
            jobID = jsonInfo.jobID
            jobState = jsonInfo.jobState
            startDate = jsonInfo.startDate
            endDate = jsonInfo.endDate
            errorCode = jsonInfo.errorCode
            errorReason = jsonInfo.errorReason
            jobStartDT = jsonInfo.jobStartDT
            jobEndDT = jsonInfo.jobEndDT
            regDT = jsonInfo.regDT
        On Error GoTo 0 
    End sub
End Class

Class EasyFinBankAccountForm
    public BankCode
    public AccountNumber
    public AccountPWD
    public AccountType
    public IdentityNumber
    public AccountName
    public BankID
    public FastID
    public FastPWD
    public UsePeriod
    public Memo
        
    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.set "BankCode", BankCode
        toJsonInfo.set "AccountNumber", AccountNumber
        toJsonInfo.set "AccountPWD", AccountPWD
        toJsonInfo.set "AccountType", AccountType
        toJsonInfo.set "IdentityNumber", IdentityNumber
        toJsonInfo.set "AccountName", AccountName
        toJsonInfo.set "BankID", BankID
        toJsonInfo.set "FastID", FastID
        toJsonInfo.set "FastPWD", FastPWD
        toJsonInfo.set "Memo", Memo
    End Function

End Class


Class EasyFinBankAccount

    Public accountNumber
    Public bankCode
    Public accountName
    Public accountType
    Public state
    Public regDT
    Public memo

    Public contractDT
    Public useEndDate
    Public baseDate
    Public contractState
    Public closeRequestYN
    Public useRestrictYN
    Public closeOnExpired
    Public unPaidYN

    Public Sub fromJsonInfo ( jsonInfo )
        On Error Resume Next

            accountNumber = jsonInfo.accountNumber
            bankCode = jsonInfo.bankCode
            accountName = jsonInfo.accountName
            accountType = jsonInfo.accountType
            state = jsonInfo.state
            regDT = jsonInfo.regDT
            memo = jsonInfo.memo

            contractDT = jsonInfo.contractDT
            useEndDate = jsonInfo.useEndDate
            baseDate = jsonInfo.baseDate
            contractState = jsonInfo.contractState
            closeRequestYN = jsonInfo.closeRequestYN
            useRestrictYN = jsonInfo.useRestrictYN
            closeOnExpired = jsonInfo.closeOnExpired
            unPaidYN = jsonInfo.unPaidYN


        On Error GoTo 0 
    End sub
End Class
%>