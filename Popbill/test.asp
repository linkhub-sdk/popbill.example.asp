<!--#include file="Popbill.asp"--> 
<!--#include file="TaxinvoiceService.asp"--> 
<html>
<head>
	<title>ASP 참 그지같다.</title>
	<meta http-equiv="Content-Type" content="text/html;charset=euc-kr" />
</head>
<body>
<div>
<%

	set m_TaxinvoiceService = new TaxinvoiceService
	m_TaxinvoiceService.Initialize "TESTER", "t4B19Ph5K2aIh9oNd91Q99Vwe9jST2/2IJbWjxhCgsA="
	m_TaxinvoiceService.IsTest = True
	
	On Error Resume Next

	remainPoint = m_TaxinvoiceService.getBalance("1231212312")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "getBalance : " + CStr(remainpoint)
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	remainPoint = m_TaxinvoiceService.getPartnerBalance("1231212312")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "getPartnerBalance : " + CStr(remainpoint)
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	url = m_TaxinvoiceService.GetPopbillURL("1231212312","userid","CHRG")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "GetPopbillURL : " + url
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.CheckIsMember("1231212312","TESTER")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "CheckIsMember : [" + CStr(Presponse.code) & " ] "  & Presponse.message
	End If

	On Error GoTo 0


	Response.write "<br/>"
	
	Set newTaxinvoice = New Taxinvoice
	newTaxinvoice.writeDate = "20150105"             '필수, 기재상 작성일자
    newTaxinvoice.chargeDirection = "정과금"         '필수, {정과금, 역과금}
    newTaxinvoice.issueType = "정발행"               '필수, {정발행, 역발행, 위수탁}
    newTaxinvoice.purposeType = "영수"               '필수, {영수, 청구}
    newTaxinvoice.issueTiming = "직접발행"           '필수, {직접발행, 승인시자동발행}
    newTaxinvoice.taxType = "과세"                   '필수, {과세, 영세, 면세}
    
    
    newTaxinvoice.invoicerCorpNum = "1231212312"
    newTaxinvoice.invoicerTaxRegID = "" '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    newTaxinvoice.invoicerCorpName = "공급자 상호&%$@<>^^"
    newTaxinvoice.invoicerMgtKey = "1234567890"    '공급자 파트너 관리번호
    newTaxinvoice.invoicerCEOName = "공급자"" 대표자 성명"
    newTaxinvoice.invoicerAddr = "공급자 주소"
    newTaxinvoice.invoicerBizClass = "공급자 업종"
    newTaxinvoice.invoicerBizType = "공급자 업태,업태2"
    newTaxinvoice.invoicerContactName = "공급자 담당자명"
    newTaxinvoice.invoicerEmail = "test@test.com"
    newTaxinvoice.invoicerTEL = "070-7070-0707"
    newTaxinvoice.invoicerHP = "010-000-2222"
    newTaxinvoice.invoicerSMSSendYN = True '발행시 문자발송기능 사용시 활용
    
    newTaxinvoice.invoiceeType = "사업자"
    newTaxinvoice.invoiceeCorpNum = "8888888888"
    newTaxinvoice.invoiceeCorpName = "공급받는자 상호"
    newTaxinvoice.invoiceeMgtKey = ""
    newTaxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    newTaxinvoice.invoiceeAddr = "공급받는자 주소"
    newTaxinvoice.invoiceeBizClass = "공급받는자 업종"
    newTaxinvoice.invoiceeBizType = "공급받는자 업태"
    newTaxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    newTaxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    newTaxinvoice.supplyCostTotal = "100000"         '필수 공급가액 합계
    newTaxinvoice.taxTotal = "10000"                 '필수 세액 합계
    newTaxinvoice.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    newTaxinvoice.modifyCode = "" '수정세금계산서 작성시 1~6까지 선택기재.
    newTaxinvoice.originalTaxinvoiceKey = "" '수정세금계산서 작성시 원본세금계산서의 ItemKey기재. ItemKey는 문서확인.
    newTaxinvoice.serialNum = "123"
    newTaxinvoice.cash = ""          '현금
    newTaxinvoice.chkBill = ""       '수표
    newTaxinvoice.note = ""          '어음
    newTaxinvoice.credit = ""        '외상미수금
    newTaxinvoice.remark1 = "비고1"
    newTaxinvoice.remark2 = "비고2"
    newTaxinvoice.remark3 = "비고3"
    newTaxinvoice.kwon = "1"
    newTaxinvoice.ho = "1"
    
    newTaxinvoice.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    newTaxinvoice.bankBookYN = False         '통장사본 이미지 첨부시 설정.
  

	'상세항목 추가.
    
    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20140410"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.unitCost = "100000"       ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "비고"

    newTaxinvoice.AddDetail newDetail

    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 2
    newDetail.itemName = "품명2"
    
    newTaxinvoice.AddDetail newDetail
 

	'추가담당자 추가. 옵션.
    set newContact = New Contact
    newContact.contactName = "담당자 성명"
    newContact.email = "test2@test.com"
    
    newTaxinvoice.AddContact newContact
    
	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.Register("1231212312",newTaxinvoice,false,"")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "Register : [" + CStr(Presponse.code) & " ] "  & Presponse.message
	End If

	On Error GoTo 0


	Response.write "<br/>"
	On Error Resume Next

	Set taxinvoiceInfo = m_TaxinvoiceService.GetDetailInfo("1231212312",SELL,"1234567890")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "GetDetailInfo : " & taxinvoiceInfo.InvoicerCorpName & "|" &  (taxinvoiceInfo.detailList.Get(0).itemName)
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	Set Presonse = m_TaxinvoiceService.AttachFile("1231212312",SELL,"1234567890","C:\Inetpub\wwwroot\Popbill\로고.gif","userid")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "AttachFile : [" + CStr(Presponse.code) & " ] "  & Presponse.message
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.Delete("1231212312",SELL,"1234567890","")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "Delete : [" + CStr(Presponse.code) & " ] "  & Presponse.message
	End If

	On Error GoTo 0

%>
</div>
</body>
</html>