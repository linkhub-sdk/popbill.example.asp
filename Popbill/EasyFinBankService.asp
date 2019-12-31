<%

Class EasyFinBankSErvice
	
	Private m_PopbillBase
	
	'�׽�Ʈ �÷���
	Public Property Let IsTest(ByVal value)
		m_PopbillBase.IsTest = value
	End Property
	
	Public Property Let IPRestrictOnOff(ByVal value)
	    m_PopbillBase.IPRestrictOnOff = value
	End Property
	
	Public Sub Class_Initialize
		Set m_PopbillBase = New PopbillBase
		m_PopbillBase.AddScope("180")
	End Sub
	
	Public Sub Initialize(linkID, SecretKey )
		m_PopbillBase.Initialize linkID,SecretKey
	End Sub

	'ȸ�� ����Ʈ��ȸ
	Public Function GetBalance(CorpNum)
		GetBalance = m_PopbillBase.GetBalance(CorpNum)
	End Function

	'��Ʈ�� �ܾ���ȸ
	Public Function GetPartnerBalance(CorpNum)
		GetPartnerBalance = m_PopbillBase.GetPartnerBalance(CorpNum)
	End Function
	
	'��Ʈ�� ����Ʈ ���� �˾� URL - 2017/08/29 �߰�
	Public Function GetPartnerURL(CorpNum, TOGO)
		GetPartnerURL = m_PopbillBase.GetPartnerURL(CorpNum,TOGO)
	End Function

	'�˺� �⺻ URL
	Public Function GetPopbillURL(CorpNum , UserID , TOGO )
		GetPopbillURL = m_PopbillBase.GetPopbillURL(CorpNum , UserID , TOGO )
	End Function

	'�˺� �α��� URL
	Public Function GetAccessURL(CorpNum , UserID)
		GetAccessURL = m_PopbillBase.GetAccessURL(CorpNum , UserID )
	End Function

	'�˺� ����ȸ�� ����Ʈ ���� URL
	Public Function GetChargeURL(CorpNum , UserID)
		GetChargeURL = m_PopbillBase.GetChargeURL(CorpNum , UserID )
	End Function

	'ȸ������ ����
	Public Function CheckIsMember(CorpNum , linkID)
		Set CheckIsMember = m_PopbillBase.CheckIsMember(CorpNum,linkID)
	End Function

	'ȸ������
	Public Function JoinMember(JoinInfo)
		Set JoinMember = m_PopbillBase.JoinMember(JoinInfo)
	End Function

	'����� �����ȸ
	Public Function ListContact(CorpNum, UserID)
		Set ListContact = m_popbillBase.ListContact(CorpNum,UserID)
	End Function

	'����� ��������
	Public Function UpdateContact(CorpNum, contInfo, UserId)
		Set UpdateContact = m_popbillBase.UpdateContact(CorpNum, contInfo, UserId)
	End Function

	'����� �߰� 
	Public Function RegistContact(CorpNum, contInfo, UserId)
		Set RegistContact = m_popbillBase.RegistContact(CorpNum, contInfo, UserId)
	End Function

	'ȸ������ ����
	Public Function UpdateCorpInfo(CorpNum, corpInfo, UserId)
		Set UpdateCorpInfo = m_popbillBase.UpdateCorpInfo(CorpNum, corpInfo, UserId)
	End Function

	'ȸ������ Ȯ�� 
	Public Function GetCorpInfo(CorpNum, UserId)
		Set GetCorpInfo = m_popbillBase.GetCorpInfo(CorpNum, UserId)
	End Function

	Public Function CheckID(id)
		Set CheckID = m_popbillBase.CheckID(id)
	End Function

	'�������� Ȯ��
	Public Function GetChargeInfo ( CorpNum, UserID )
		Set result = m_PopbillBase.httpGET("/EasyFin/Bank/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)
		Set chrgInfo = New ChargeInfo
		chrgInfo.fromJsonInfo result
		
		Set GetChargeInfo = chrgInfo
	End Function 

	'''''''''''''  End of PopbillBase

	Public Function GetBankAccountMgtURL ( CorpNum, UserID )
		Set result = m_PopbillBase.httpGET("/EasyFin/Bank?TG=BankAccount", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)
		GetBankAccountMgtURL = result.url
	End Function

	Public Function ListBankAccount(CorpNum, UserID)
		Set result = m_PopbillBase.httpGET("/EasyFin/Bank/ListBankAccount", _
						m_PopbillBase.getSession_token(CorpNum), UserID)
		
		Set bankAccountList = CreateObject("Scripting.Dictionary")
		For i=0 To result.length-1
			Set tmpInfo = New EasyFinBankAccount
			tmpInfo.fromJsonInfo result.Get(i)
			bankAccountList.Add i, tmpInfo
		Next
		Set ListBankAccount = bankAccountList
	End Function

	Public Function RequestJob(CorpNum , BankCode, AccountNumber, SDate, EDate, UserID)
		uri = "/EasyFin/Bank/BankAccount"
		uri = uri + "?BankCode=" & BankCode
		uri = uri + "&AccountNumber=" & AccountNumber
		uri = uri + "&SDate=" & SDate
		uri = uri + "&EDate=" & EDate
		Set result = m_PopbillBase.httpPOST( uri, m_PopbillBase.getSession_token(CorpNum),"", "", UserID )

		RequestJob = result.jobID
	End Function

	Public Function GetJobState(CorpNum, JobID, UserID)
		If Len(JobID) <> 18  Then
			Err.Raise -99999999, "POPBILL", "�۾����̵� �ùٸ��� �ʽ��ϴ�."
		End If

		Set result = m_PopbillBase.httpGET("/EasyFin/Bank/" & JobID & "/State", _
						m_PopbillBase.getSession_token(CorpNum), UserID)

		Set jobInfo = New EasyFinJobState	
		jobInfo.fromJsonInfo result
		Set GetJobState = jobInfo
	End Function

	Public Function ListActiveJob(CorpNum, UserID)
		Set result = m_PopbillBase.httpGET("/EasyFin/Bank/JobList", _
						m_PopbillBase.getSession_token(CorpNum), UserID)
		
		Set jobList = CreateObject("Scripting.Dictionary")

		For i=0 To result.length-1
			Set jobInfo = New EasyFinJobState
			jobInfo.fromJsonInfo result.Get(i)
			jobList.Add i, jobInfo
		Next

		Set ListActiveJob = jobList
	End Function

	Public Function Search ( CorpNum, JobID, TradeType, SearchString, Page, PerPage, Order, UserID )

		If  Not ( Len ( JobID ) = 18 )  Then
			Err.Raise -99999999, "POPBILL", "�۾����̵� �ùٸ��� �ʽ��ϴ�."
		End If 

		uri = "/EasyFin/Bank/" & JobID
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

		Set result = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), UserID)

		Set searchResult = New EasyFinBankSearchResult
		searchResult.fromJsonInfo result
		Set Search = searchResult 

	End Function 

	Public Function Summary ( CorpNum, JobID, TradeType, SearchString, UserID)

		If Not ( Len ( JobID ) = 18 ) Then
			Err.Raise -99999999, "POPBILL", "�۾����̵� �ùٸ��� �ʽ��ϴ�."
		End If 

		uri = "/EasyFin/Bank/" & JobID & "/Summary"

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

		Set result = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), UserID)
	
		Set summaryResult = New EasyFinBankSummaryResult
		summaryResult.fromJsonInfo result
		Set Summary = summaryResult

	End Function

	Public Function SaveMemo(CorpNum , TID, Memo, UserID)

		If TID = "" Then
			Err.Raise -99999999, "POPBILL", "�ŷ����� ���̵� �Էµ��� �ʾҽ��ϴ�."
		End If

		uri = "/EasyFin/Bank/SaveMemo"
		uri = uri + "?TID=" & TID
		uri = uri + "&Memo=" & Memo
		Set SaveMemo = m_PopbillBase.httpPOST( uri, m_PopbillBase.getSession_token(CorpNum),"", "", UserID )

	End Function

	Public Function GetFlatRatePopUpURL ( CorpNum, UserID )

		Set result = m_PopbillBase.httpGET("/EasyFin/Bank?TG=CHRG", m_PopbillBase.getSession_token(CorpNum), UserID)
		GetFlatRatePopUpURL = result.url

	End Function

	Public Function GetFlatRateState ( CorpNum, BankCode, AccountNumber, UserID ) 

		If BankCode = "" Then
			Err.Raise -99999999, "POPBILL", "�����ڵ尡 �Էµ��� �ʾҽ��ϴ�."
		End If
		If AccountNumber = "" Then
			Err.Raise -99999999, "POPBILL", "���¹�ȣ�� �Էµ��� �ʾҽ��ϴ�."
		End If

		Set responseObj = m_PopbillBase.httpGET("/EasyFin/Bank/Contract/" & BankCode & "/" & AccountNumber, _
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
		For i = 0 To jsonInfo.list.length -1
			Set tmpObj = New EasyFinSearchDetail
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

Class EasyFinBankAccount

	Public accountNumber
	Public bankCode
	Public accountName
	Public accountType
	Public state
	Public regDT
	Public memo

	Public Sub fromJsonInfo ( jsonInfo )
		On Error Resume Next
			accountNumber = jsonInfo.accountNumber
			bankCode = jsonInfo.bankCode
			accountName = jsonInfo.accountName
			accountType = jsonInfo.accountType
			state = jsonInfo.state
			regDT = jsonInfo.regDT
			memo = jsonInfo.memo
		On Error GoTo 0 
	End sub
End Class
%>