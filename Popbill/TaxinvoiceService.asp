<%
Const SELL = "SELL"
Const BUY = "BUY"
Const TRUSTEE = "TRUSTEE"

Class TaxinvoiceService

Private m_PopbillBase

'테스트 플래그
Public Property Let IsTest(ByVal value)
    m_PopbillBase.IsTest = value
End Property

Public Sub Class_Initialize
	Set m_PopbillBase = New PopbillBase
	m_PopbillBase.AddScope("110")
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
'''''''''''''  End of PopbillBase

'임시저장
Public Function Register(CorpNum ,byref TI , writeSpecification ,  UserID)
	
	If TI Is Nothing Then Err.raise -99999999,"POPBILL","등록할 세금계산서 정보가 입력되지 않았습니다."

    Set tmpDic = TI.toJsonInfo
	If writeSpecification Then
        tmpDic.Add "writeSpecification", True
    End If
    
    postdata = m_PopbillBase.toString(tmpDic)

    Set Register = m_PopbillBase.httpPOST("/Taxinvoice", m_PopbillBase.getSession_token(CorpNum),"", postdata, UserID)
End Function
'파일 첨부
Public Function AttachFile(CorpNum , KeyType , MgtKey , FilePath ,  UserID )
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Set AttachFile = m_PopbillBase.httpPOST_File("/Taxinvoice/" + KeyType + "/" + MgtKey + "/Files", _
                        m_PopbillBase.getSession_token(CorpNum), FilePath, UserID)
End Function
'상세정보 확인
Public Function GetDetailInfo(CorpNum , KeyType, MgtKey )
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Set GetDetailInfo = m_PopbillBase.httpGET("/Taxinvoice/" + KeyType + "/" + MgtKey + "?Detail", _
                                m_PopbillBase.getSession_token(CorpNum), "")
End Function

'삭제
Public Function Delete(CorpNum , KeyType , MgtKey ,  UserID )
    If MgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "관리번호가 입력되지 않았습니다."
    End If
    
    Set Delete = m_PopbillBase.httpPOST("/Taxinvoice/" + KeyType + "/" + MgtKey, m_PopbillBase.getSession_token(CorpNum), "DELETE", "", UserID)
End Function

End Class

''Taxinvoice Class
Class Taxinvoice

Public writeDate            
Public chargeDirection      
Public issueType            
Public issueTiming          
Public taxType              

Public invoicerCorpNum      
Public invoicerMgtKey       
Public invoicerTaxRegID     
Public invoicerCorpName     
Public invoicerCEOName      
Public invoicerAddr         
Public invoicerBizClass     
Public invoicerBizType      
Public invoicerContactName  
Public invoicerDeptName     
Public invoicerTEL          
Public invoicerHP           
Public invoicerEmail        
Public invoicerSMSSendYN    

Public invoiceeType         
Public invoiceeCorpNum      
Public invoiceeMgtKey       
Public invoiceeTaxRegID     
Public invoiceeCorpName     
Public invoiceeCEOName      
Public invoiceeAddr         
Public invoiceeBizClass     
Public invoiceeBizType      
Public invoiceeContactName1 
Public invoiceeDeptName1    
Public invoiceeTEL1         
Public invoiceeHP1          
Public invoiceeEmail1       
Public invoiceeContactName2 
Public invoiceeDeptName2    
Public invoiceeTEL2         
Public invoiceeHP2          
Public invoiceeEmail2       
Public invoiceeSMSSendYN   

Public trusteeCorpNum       
Public trusteeMgtKey        
Public trusteeTaxRegID      
Public trusteeCorpName      
Public trusteeCEOName       
Public trusteeAddr          
Public trusteeBizClass      
Public trusteeBizType       
Public trusteeContactName   
Public trusteeDeptName      
Public trusteeTEL           
Public trusteeHP            
Public trusteeEmail         
Public trusteeSMSSendYN

Public taxTotal             
Public supplyCostTotal      
Public totalAmount          
Public modifyCode           
Public orgNTSConfirmNum     
Public purposeType          
Public serialNum            
Public cash                 
Public chkBill              
Public credit               
Public note                 
Public remark1              
Public remark2              
Public remark3              
Public kwon                 
Public ho                   
Public businessLicenseYN    
Public bankBookYN                 
Public ntsconfirmNum        
Public originalTaxinvoiceKey 
Public detailList()
Public addContactList()

Public Sub Class_Initialize
	ReDim detailList(-1)
	ReDim addContactList(-1)
End Sub

Public Function toJsonInfo()
    Set toJsonInfo = JSON.parse("{}")
    
    toJsonInfo.set "writeDate", writeDate
    
    toJsonInfo.set "chargeDirection", chargeDirection
    toJsonInfo.set "issueType", issueType
    toJsonInfo.set "issueTiming", issueTiming
    toJsonInfo.set "taxType", taxType
    toJsonInfo.set "invoicerCorpNum", invoicerCorpNum
    toJsonInfo.set "invoicerMgtKey", invoicerMgtKey
    toJsonInfo.set "invoicerTaxRegID", invoicerTaxRegID
    toJsonInfo.set "invoicerCorpName", invoicerCorpName
    toJsonInfo.set "invoicerCEOName", invoicerCEOName
    toJsonInfo.set "invoicerAddr", invoicerAddr
    toJsonInfo.set "invoicerBizClass", invoicerBizClass
    toJsonInfo.set "invoicerBizType", invoicerBizType
    toJsonInfo.set "invoicerContactName", invoicerContactName
    toJsonInfo.set "invoicerDeptName", invoicerDeptName
    toJsonInfo.set "invoicerTEL", invoicerTEL
    toJsonInfo.set "invoicerHP", invoicerHP
    toJsonInfo.set "invoicerEmail", invoicerEmail
    toJsonInfo.set "invoicerSMSSendYN", invoicerSMSSendYN
    
    toJsonInfo.set "invoiceeCorpNum", invoiceeCorpNum
    toJsonInfo.set "invoiceeType", invoiceeType
    toJsonInfo.set "invoiceeMgtKey", invoiceeMgtKey
    toJsonInfo.set "invoiceeTaxRegID", invoiceeTaxRegID
    toJsonInfo.set "invoiceeCorpName", invoiceeCorpName
    toJsonInfo.set "invoiceeCEOName", invoiceeCEOName
    toJsonInfo.set "invoiceeAddr", invoiceeAddr
    toJsonInfo.set "invoiceeBizType", invoiceeBizType
    toJsonInfo.set "invoiceeBizClass", invoiceeBizClass
    toJsonInfo.set "invoiceeContactName1", invoiceeContactName1
    toJsonInfo.set "invoiceeDeptName1", invoiceeDeptName1
    toJsonInfo.set "invoiceeTEL1", invoiceeTEL1
    toJsonInfo.set "invoiceeHP1", invoiceeHP1
    toJsonInfo.set "invoiceeEmail1", invoiceeEmail1
    toJsonInfo.set "invoiceeContactName2", invoiceeContactName2
    toJsonInfo.set "invoiceeDeptName2", invoiceeDeptName2
    toJsonInfo.set "invoiceeTEL2", invoiceeTEL2
    toJsonInfo.set "invoiceeEmail2", invoiceeEmail2
    toJsonInfo.set "invoiceeSMSSendYN", invoiceeSMSSendYN
    
    toJsonInfo.set "trusteeCorpNum", trusteeCorpNum
    toJsonInfo.set "trusteeMgtKey", trusteeMgtKey
    toJsonInfo.set "trusteeTaxRegID", trusteeTaxRegID
    toJsonInfo.set "trusteeCorpName", trusteeCorpName
    toJsonInfo.set "trusteeCEOName", trusteeCEOName
    toJsonInfo.set "trusteeAddr", trusteeAddr
    toJsonInfo.set "trusteeBizType", trusteeBizType
    toJsonInfo.set "trusteeBizClass", trusteeBizClass
    toJsonInfo.set "trusteeContactName", trusteeContactName
    toJsonInfo.set "trusteeDeptName", trusteeDeptName
    toJsonInfo.set "trusteeTEL", trusteeTEL
    toJsonInfo.set "trusteeHP", trusteeHP
    toJsonInfo.set "trusteeEmail", trusteeEmail
    toJsonInfo.set "trusteeSMSSendYN", trusteeSMSSendYN
    
    toJsonInfo.set "taxTotal", taxTotal
    toJsonInfo.set "supplyCostTotal", supplyCostTotal
    toJsonInfo.set "totalAmount", totalAmount
    If modifyCode <> "" Then
        toJsonInfo.set "modifyCode", CInt(modifyCode)
    End If
    
    toJsonInfo.set "orgNTSConfirmNum", orgNTSConfirmNum
    toJsonInfo.set "purposeType", purposeType
    toJsonInfo.set "serialNum", serialNum
    toJsonInfo.set "cash", cash
    toJsonInfo.set "chkBill", chkBill
    toJsonInfo.set "credit", credit
    toJsonInfo.set "note", note
    If kwon <> "" Then
        toJsonInfo.set "kwon", CInt(kwon)
    End If
    If ho <> "" Then
        toJsonInfo.set "ho", CInt(ho)
    End If
    
    toJsonInfo.set "businessLicenseYN", businessLicenseYN
    toJsonInfo.set "bankBookYN", bankBookYN
    
    toJsonInfo.set "remark1", remark1
    toJsonInfo.set "remark2", remark2
    toJsonInfo.set "remark3", remark3
    
    toJsonInfo.set "ntsconfirmNum", ntsconfirmNum
    toJsonInfo.set "originalTaxinvoiceKey", originalTaxinvoiceKey
    
	Dim detailJsonInfo()
	ReDim detailJsonInfo(UBound(detailList))
	i = 0
	For Each detail In detailList
		Set detailJsonInfo(i) = detailList(i).toJsonInfo
		i = i + 1
	next
	toJsonInfo.set "detailList", detailJsonInfo


	Dim addContactListJson()
	ReDim addContactListJson(UBound(addContactList))
	i = 0
	For Each detail In addContactList
		Set addContactListJson(i) = addContactList(i).toJsonInfo
		i = i + 1
	next
	toJsonInfo.set "addContactList", addContactListJson
    
End Function

Public Sub AddDetail(detail)
	ReDim Preserve detailList(UBound(detailList) + 1)
	Set detailList(Ubound(detailList)) = detail
End Sub

Public Sub AddContact(contact)
	ReDim Preserve addContactList(UBound(addContactList) + 1)
	Set addContactList(Ubound(addContactList)) = contact
End Sub

End Class

Class TaxinvoiceDetail
Public serialNum       
Public purchaseDT      
Public itemName        
Public spec            
Public qty             
Public unitCost        
Public supplyCost      
Public tax             
Public remark          

Public Function toJsonInfo() 
    Set toJsonInfo = JSON.parse("{}")
    
    toJsonInfo.set "serialNum", CInt(serialNum)
    toJsonInfo.set "purchaseDT", purchaseDT
    toJsonInfo.set "itemName", itemName
    toJsonInfo.set "spec", spec
    toJsonInfo.set "qty", qty
    toJsonInfo.set "unitCost", unitCost
    toJsonInfo.set "supplyCost", supplyCost
    toJsonInfo.set "tax", tax
    toJsonInfo.set "remark", remark
End Function
End Class

Class Contact
Public serialNum
Public email    
Public contactName

Public Function toJsonInfo() 
    Set toJsonInfo = JSON.parse("{}")
    
    toJsonInfo.set "serialNum", CInt(serialNum)
    toJsonInfo.set "email", email
    toJsonInfo.set "contactName", contactName
End Function
End Class
%>