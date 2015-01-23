<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	Set newTaxinvoice = New Taxinvoice

	newTaxinvoice.writeDate = "20140122"             '필수, 기재상 작성일자
    newTaxinvoice.chargeDirection = "정과금"         '필수, {정과금, 역과금}
    newTaxinvoice.issueType = "정발행"               '필수, {정발행, 역발행, 위수탁}
    newTaxinvoice.purposeType = "영수"               '필수, {영수, 청구}
    newTaxinvoice.issueTiming = "직접발행"           '필수, {직접발행, 승인시자동발행}
    newTaxinvoice.taxType = "과세"                   '필수, {과세, 영세, 면세}
    
    
    newTaxinvoice.invoicerCorpNum = "1234567890"
    newTaxinvoice.invoicerTaxRegID = ""					'종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    newTaxinvoice.invoicerCorpName = "공급자 상호"
    newTaxinvoice.invoicerMgtKey = "20150122-29"		'공급자 파트너 관리번호
    newTaxinvoice.invoicerCEOName = "공급자 대표자 성명"
    newTaxinvoice.invoicerAddr = "공급자 주소"
    newTaxinvoice.invoicerBizClass = "공급자 업종"
    newTaxinvoice.invoicerBizType = "공급자 업태,업태2"
    newTaxinvoice.invoicerContactName = "공급자 담당자명"
    newTaxinvoice.invoicerEmail = "test@test.com"
    newTaxinvoice.invoicerTEL = "070-7070-0707"
    newTaxinvoice.invoicerHP = "010-000-2222"
    newTaxinvoice.invoicerSMSSendYN = False			'발행시 문자발송기능 사용시 활용
    
    newTaxinvoice.invoiceeType = "사업자"
    newTaxinvoice.invoiceeCorpNum = "1231212312"
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
    
    newTaxinvoice.modifyCode = ""				'수정세금계산서 작성시 1~6까지 선택기재.
    newTaxinvoice.originalTaxinvoiceKey = ""	'수정세금계산서 작성시 원본세금계산서의 ItemKey기재. ItemKey는 문서확인(getInfo.asp) API 통해 확인.
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

	testCorpNum = "1234567890"		'팝빌회원 사업자번호
	writeSpecificationYN = False	'거래명세서 동시작성여부
	userID = "testkorea"			'회원 아이디

	Set Presponse = m_TaxinvoiceService.Register(testCorpNum, newTaxinvoice, writeSpecificationYN, userID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = Presponse.code
		message =Presponse.message
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>세금계산서 임시저장</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>