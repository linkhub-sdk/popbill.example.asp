<%

Class AccountCheckService

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
        m_PopbillBase.AddScope("182")
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

    '팝빌 연동회원 포인트 충전내역 URL
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
        Dim result : Set result = m_PopbillBase.httpGET("/EasyFin/AccountCheck/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)

        Dim chrgInfo : Set chrgInfo = New ChargeInfo
        chrgInfo.fromJsonInfo result
        
        Set GetChargeInfo = chrgInfo
    End Function
    '''''''''''''  End of PopbillBase

    '조회단가확인
    Public Function GetUnitCost(CorpNum)
        Dim result : Set result = m_PopbillBase.httpGET("/EasyFin/AccountCheck/UnitCost", m_PopbillBase.getSession_token(CorpNum),"")
        GetUnitCost = result.unitCost
    End Function

    Public Function CheckAccountInfo(CorpNum , BankCode, AccountNumber, UserID)

        If BankCode = "" Then
            Err.Raise -99999999, "POPBILL", "기관코드가 입력되지 않았습니다."
        End If

        If AccountNumber = "" Then
            Err.Raise -99999999, "POPBILL", "계좌번호가 입력되지 않았습니다."
        End If

        Dim uri
        uri = "/EasyFin/AccountCheck"
        uri = uri + "?c=" & BankCode
        uri = uri + "&n=" & AccountNumber
    

        Dim result : Set result = m_PopbillBase.httpPOST( uri, m_PopbillBase.getSession_token(CorpNum),"", "", UserID )

        Dim infoObj : Set infoObj = New AccountCheckInfo
        infoObj.fromJsonInfo result
        Set CheckAccountInfo = infoObj

    End Function


End Class

Class AccountCheckInfo

    Public resultCode
    Public resultMessage
    Public bankCode
    Public accountNumber
    Public accountName
    Public checkDate

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
            If Not isEmpty(jsonInfo.resultCode) Then
                resultCode = jsonInfo.resultCode
            End If 

            If Not isEmpty(jsonInfo.resultMessage) Then
                resultMessage = jsonInfo.resultMessage
            End If 

            If Not isEmpty(jsonInfo.bankCode) Then
                bankCode = jsonInfo.bankCode
            End If 

            If Not isEmpty(jsonInfo.accountNumber) Then
                accountNumber = jsonInfo.accountNumber
            End If 

            If Not isEmpty(jsonInfo.accountName) Then
                accountName = jsonInfo.accountName
            End If 

            If Not isEmpty(jsonInfo.checkDate) Then
                checkDate = jsonInfo.checkDate
            End If 
            
        On Error GoTo 0
    End Sub
End Class


%>