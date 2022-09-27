<%
Class BizInfoCheckService

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
    m_PopbillBase.AddScope("171")
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
    Dim result : Set result = m_PopbillBase.httpGET("/BizInfo/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)

    Dim chrgInfo : Set chrgInfo = New ChargeInfo
    chrgInfo.fromJsonInfo result
    
    Set GetChargeInfo = chrgInfo
End Function
'''''''''''''  End of PopbillBase

'조회단가확인
Public Function GetUnitCost(CorpNum)
    Dim result : Set result = m_PopbillBase.httpGET("/BizInfo/UnitCost", m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'기업정보조회 단건
Public Function CheckBizInfo(MemberCorpNum, CorpNum, UserID) 
    If MemberCorpNum = "" Or isEmpty(MemberCorpNum) Then 
        Err.Raise -99999999, "POPBILL", "팝빌회원 사업자번호가 입력되지 않았습니다."
    End If

    If CorpNum = "" Or isEmpty(CorpNum) Then 
        Err.Raise -99999999, "POPBILL", "조회할 사업자번호가 입력되지 않았습니다."
    End If
    
    Dim result : Set result = m_PopbillBase.httpGet("/BizInfo/Check?CN="+CorpNum, m_PopbillBase.getSession_token(MemberCorpNum), UserID)
    
    Dim stateObj : Set stateObj = New BizCheckInfo
    stateObj.fromJsonInfo result
    Set CheckBizInfo = stateObj

End Function 
End Class

Class BizCheckInfo
    Public corpNum
    Public checkDT
    Public corpName
    Public corpCode
    Public corpScaleCode
    Public personCorpCode
    Public headOfficeCode
    Public industryCode
    Public companyRegNum
    Public establishDate
    Public establishCode
    Public ceoname
    Public workPlaceCode
    Public addrCode
    Public zipCode
    Public addr
    Public addrDetail
    Public enAddr
    Public bizClass
    Public bizType
    Public result
    Public resultMessage
    Public closeDownTaxType
    Public closeDownTaxTypeDate
    Public closeDownState
    Public closeDownStateDate

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
            If Not isEmpty(jsonInfo.corpNum) Then
                corpNum = jsonInfo.corpNum
            End If 
            If Not isEmpty(jsonInfo.checkDT) Then
                checkDT = jsonInfo.checkDT
            End If 
            If Not isEmpty(jsonInfo.corpName) Then
                corpName = jsonInfo.corpName
            End If 
            If Not isEmpty(jsonInfo.corpCode) Then
                corpCode = jsonInfo.corpCode
            End If 
            If Not isEmpty(jsonInfo.corpScaleCode) Then
                corpScaleCode = jsonInfo.corpScaleCode
            End If 
            If Not isEmpty(jsonInfo.personCorpCode) Then
                personCorpCode = jsonInfo.personCorpCode
            End If 
            If Not isEmpty(jsonInfo.headOfficeCode) Then
                headOfficeCode = jsonInfo.headOfficeCode
            End If 
            If Not isEmpty(jsonInfo.industryCode) Then
                industryCode = jsonInfo.industryCode
            End If 
            If Not isEmpty(jsonInfo.companyRegNum) Then
                companyRegNum = jsonInfo.companyRegNum
            End If 
            If Not isEmpty(jsonInfo.establishDate) Then
                establishDate = jsonInfo.establishDate
            End If 
            If Not isEmpty(jsonInfo.establishCode) Then
                establishCode = jsonInfo.establishCode
            End If 
            If Not isEmpty(jsonInfo.ceoname) Then
                ceoname = jsonInfo.ceoname
            End If 
            If Not isEmpty(jsonInfo.workPlaceCode) Then
                workPlaceCode = jsonInfo.workPlaceCode
            End If 
            If Not isEmpty(jsonInfo.addrCode) Then
                addrCode = jsonInfo.addrCode
            End If 
            If Not isEmpty(jsonInfo.zipCode) Then
                zipCode = jsonInfo.zipCode
            End If 
            If Not isEmpty(jsonInfo.addr) Then
                addr = jsonInfo.addr
            End If 
            If Not isEmpty(jsonInfo.addrDetail) Then
                addrDetail = jsonInfo.addrDetail
            End If 
            If Not isEmpty(jsonInfo.enAddr) Then
                enAddr = jsonInfo.enAddr
            End If 
            If Not isEmpty(jsonInfo.bizClass) Then
                bizClass = jsonInfo.bizClass
            End If 
            If Not isEmpty(jsonInfo.bizType) Then
                bizType = jsonInfo.bizType
            End If 
            If Not isEmpty(jsonInfo.result) Then
                result = jsonInfo.result
            End If 
            If Not isEmpty(jsonInfo.resultMessage) Then
                resultMessage = jsonInfo.resultMessage
            End If 
            If Not isEmpty(jsonInfo.closeDownTaxType) Then
                closeDownTaxType = jsonInfo.closeDownTaxType
            End If 
            If Not isEmpty(jsonInfo.closeDownTaxTypeDate) Then
            closeDownTaxTypeDate = jsonInfo.closeDownTaxTypeDate
            End If 
            If Not isEmpty(jsonInfo.closeDownState) Then
                closeDownState = jsonInfo.closeDownState
            End If 
            If Not isEmpty(jsonInfo.closeDownStateDate) Then
            closeDownStateDate = jsonInfo.closeDownStateDate
            End If 
        On Error GoTo 0
    End Sub
End Class


%>