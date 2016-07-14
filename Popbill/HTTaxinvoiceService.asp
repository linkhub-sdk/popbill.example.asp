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
		Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/ChargeInfo", m_PopbillBase.getSession_token(CorpNum), UserID)

		Set chrgInfo = New ChargeInfo
		chrgInfo.fromJsonInfo result
		
		Set GetChargeInfo = chrgInfo
	End Function 
	'''''''''''''  End of PopbillBase

	'수집요청
	Public Function RequestJob(CorpNum , KeyType, DType, SDate, Edate, UserID)
		uri = "/HomeTax/Taxinvoice/" & KeyType
		uri = uri + "?DType=" & DType
		uri = uri + "&SDate=" & SDate
		uri = uri + "&EDate=" & EDate
		Set result = m_PopbillBase.httpPOST( uri, m_PopbillBase.getSession_token(CorpNum),"", "", UserID )

		RequestJob = result.jobID
	End Function

	'수집 상태 확인
	Public Function GetJobState(CorpNum, JobID, UserID)
		If Len(JobID) <> 18  Then
			Err.Raise -99999999, "POPBILL", "작업아이디가 올바르지 않습니다."
		End If

		Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/" & JobID & "/State", _
						m_PopbillBase.getSession_token(CorpNum), UserID)

		Set jobInfo = New HTTIJobState	
		jobInfo.fromJsonInfo result
		Set GetJobState = jobInfo
	End Function

	'수집 상태 목록 확인
	Public Function ListActiveJob(CorpNum, UserID)
		Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/JobList", _
						m_PopbillBase.getSession_token(CorpNum), UserID)
		
		Set jobList = CreateObject("Scripting.Dictionary")

		For i=0 To result.length-1
			Set jobInfo = New HTTIJobState
			jobInfo.fromJsonInfo result.Get(i)
			jobList.Add i, jobInfo
		Next

		Set ListActiveJob = jobList
	End Function

	'수집 결과 조회
	Public Function Search ( CorpNum, JobID, TIType, TaxType, PurposeType, TaxRegIDYN, TaxRegIDType, TaxRegID, Page, PerPage, Order, UserID )
		If  Not ( Len ( JobID ) = 18 )  Then
			Err.Raise -99999999, "POPBILL", "작업아이디가 올바르지 않습니다."
		End If 

		uri = "/HomeTax/Taxinvoice/" & JobID
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
		uri = uri & "&Page=" & CStr(Page)
		uri = uri & "&PerPage=" & CStr(PerPage)
		uri = uri & "&Order=" & Order

		Set result = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), UserID)

		Set searchResult = New HTTaxinvoiceSerach
		searchResult.fromJsonInfo result
		Set Search = searchResult 

	End Function 

	Public Function Summary ( CorpNum, JobID, TIType, TaxType, PurposeType, TaxRegIDYN, TaxRegIDType, TaxRegID, UserID )
		If Not ( Len ( JobID ) = 18 ) Then
			Err.Raise -99999999, "POPBILL", "작업아이디가 올바르지 않습니다."
		End If 

		uri = "/HomeTax/Taxinvoice/" & JobID & "/Summary"
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

		Set result = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), UserID)
	
		Set summaryResult = New HTTaxinvoiceSummary
		summaryResult.fromJsonInfo result
		Set Summary = summaryResult

	End Function
	
	Public Function GetTaxinvoice ( CorpNum, NTSConfirmNum, UserID )
		If Not ( Len ( NTSConfirmNum ) = 24 ) Then
			Err.Raise -99999999, "POPBILL", "국세청승인번호가 올바르지 않습니다."
		End If

		Set responseObj = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/" & NTSConfirmNUm, _
								m_PopbillBase.getSession_token(CorpNum), UserID)

		Set taxinvoiceDetail = New HTTaxinvoice
		taxinvoiceDetail.fromJsonInfo responseObj
		Set GetTaxinvoice = taxinvoiceDetail
	End Function
	
	Public Function GetXML ( CorpNum, NTSConfirmNum, UserID )
		If Not ( Len ( NTSConfirmNum ) = 24 ) Then
			Err.Raise -99999999, "POPBILL", "국세청승인번호가 올바르지 않습니다."
		End If
		
		Set responseObj = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/" & NTSConfirmNum & "?T=xml", _
								m_PopbillBase.getSession_token(CorpNum), UserID)

		Set taxinvoiceXML = New HTTaxinvoiceXML 
		taxinvoiceXML.fromJsonInfo responseObj
		Set GetXML = taxinvoiceXML
	End Function
	
	Public Function GetFlatRatePopUpURL ( CorpNum, UserID )
		Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice?TG=CHRG", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)
		GetFlatRatePopUpURL = result.url
	End Function
	
	Public Function GetFlatRateState ( CorpNum, UserID ) 
		Set responseObj = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/Contract", _
						m_PopbillBase.getSession_token(CorpNum), UserID)

		Set flatRateObj = New HTTIFlatRate
		flatRateObj.fromJsonInfo responseObj
		Set GetFlatRateState = flatrateObj
	End Function 

	Public Function GetCertificatePopUpURL ( CorpNum, UserID )
		Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice?TG=CERT", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)
		GetCertificatePopUpURL = result.url
	End Function 

	Public Function GetCertificateExpireDate ( CorpNum, UserID )
		Set result = m_PopbillBase.httpGET("/HomeTax/Taxinvoice/CertInfo", _
					m_PopbillBase.getSession_token(CorpNum), UserID)
		GetCertificateExpireDate = result.certificateExpiration
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
			For i = 0 To jsonInfo.detailList.length -1 
				Set tmpDetail = New HTTaxinvoiceDetail
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
		For i = 0 To jsonInfo.list.length -1
			Set tmpObj = New HTTaxinvoiceAbbr
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