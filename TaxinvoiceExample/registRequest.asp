<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' [공급받는자]가 공급자에게 1건의 역발행 세금계산서를 [즉시 요청]합니다.
	' - 역발행 세금계산서 프로세스를 구현하기 위해서는 공급자/공급받는자가 모두 팝빌에 회원이여야 합니다.
	' - 역발행 즉시요청후 공급자가 [발행] 처리시 포인트가 차감되며 역발행 세금계산서 항목중 과금방향(ChargeDirection)에 기재한 값에 따라
	'    정과금(공급자과금) 또는 역과금(공급받는자 과금) 처리됩니다.
	' - https://docs.popbill.com/taxinvoice/asp/api#RegistRequest
	'**************************************************************
	
	' 팝빌회원 사업자번호
	testCorpNum = "1234567890"

	' 팝빌회원 아이디
	userID = "testkorea"

	
	' 세금계산서 정보 객체 생성
	Set newTaxinvoice = New Taxinvoice

	' [필수] 작성일자, 날짜형식(yyyyMMdd)
	newTaxinvoice.writeDate = "20190103"

	' [필수] {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
    newTaxinvoice.chargeDirection = "정과금"
	
	' [필수] 발행형태, {정발행, 역발행, 위수탁} 중 기재
    newTaxinvoice.issueType = "역발행"

	' [필수] {영수, 청구} 중 기재 
    newTaxinvoice.purposeType = "영수"
	
	' [필수] 과세형태,  {과세, 영세, 면세} 중 기재 
    newTaxinvoice.taxType = "과세"
    

	'**************************************************************
    '						                       공급자 정보
	'**************************************************************

    '[필수] 공급자 사업자번호, '-' 제외 10자리
    newTaxinvoice.invoicerCorpNum = "8888888888"

	'[필수] 공급자 종사업자 식별번호. 필요시 숫자 4자리 기재
    newTaxinvoice.invoicerTaxRegID = ""

    '[필수] 공급자 상호
	newTaxinvoice.invoicerCorpName = "공급자 상호"

    '공급자 문서번호, 1~24자리 (숫자, 영문, '-', '_') 조합으로
    '사업자 별로 중복되지 않도록 구성
    newTaxinvoice.invoicerMgtKey = ""

	'[필수] 공급자 대표자 성명
    newTaxinvoice.invoicerCEOName = "공급자 대표자 성명"
    
	' 공급자 주소
	newTaxinvoice.invoicerAddr = "공급자 주소"
    
	' 공급자 종목
	newTaxinvoice.invoicerBizClass = "공급자 종목"
    
	' 공급자 업태
	newTaxinvoice.invoicerBizType = "공급자 업태,업태2"
    
	' 공급자 담당자명
	newTaxinvoice.invoicerContactName = "공급자 담당자명"
    
	' 공급자 담당자 메일주소 
	newTaxinvoice.invoicerEmail = "test@test.com"
    
	' 공급자 담당자 연락처 
	newTaxinvoice.invoicerTEL = "070-7070-0707"
    
	' 공급자 담당자 휴대폰번호
	newTaxinvoice.invoicerHP = "010-000-2222"

    '정발행시 공급받는자에게 발행안내문자 전송여부
    '- 안내문자 전송기능 이용시 포인트가 차감됩니다.	
	newTaxinvoice.invoicerSMSSendYN = False
    


	'**************************************************************
    '				                            공급받는자 정보
	'**************************************************************

	'[필수] 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
    newTaxinvoice.invoiceeType = "사업자"

    '[필수] 공급받는자 사업자번호, '-' 제외 10자리
    newTaxinvoice.invoiceeCorpNum = testCorpNum

    '[필수] 공급받는자 종사업자 식별번호. 필요시 숫자 4자리 기재	
	newTaxinvoice.invoiceeTaxRegID = ""
    
	'[필수] 공급자받는자 상호
	newTaxinvoice.invoiceeCorpName = "공급받는자 상호"

    '[역발행시 필수] 공급받는자 문서번호(역발행시 필수)
    newTaxinvoice.invoiceeMgtKey = "20190103-001"

	'[필수] 공급받는자 대표자 성명
	newTaxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
    
	'공급받는자 주소
	newTaxinvoice.invoiceeAddr = "공급받는자 주소"
    
	'공급받는자 종목
	newTaxinvoice.invoiceeBizClass = "공급받는자 종목"
    
	'공급받는자 업태
	newTaxinvoice.invoiceeBizType = "공급받는자 업태"
    
	'공급받는자 담당자명
	newTaxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
    
	'공급받는자 담당자 메일주소
	'팝빌 개발환경에서 테스트하는 경우에도 안내 메일이 전송되므로,
	'실제 거래처의 메일주소가 기재되지 않도록 주의
	newTaxinvoice.invoiceeEmail1 = "test@invoicee.com"
	
	'공급받는자 연락처
	newTaxinvoice.invoiceeTEL1 = "070-111-222"
	
	'공급받는자 휴대폰번호
	newTaxinvoice.invoiceeHP1 = "010-111-222"

    '역발행시 공급자에게 발행안내문자 전송여부
    newTaxinvoice.invoiceeSMSSendYN = False



	'**************************************************************
    '				                            세금계산서 정보
	'**************************************************************

    '[필수] 공급가액 합계
    newTaxinvoice.supplyCostTotal = "100000"

    '[필수] 세액 합계
    newTaxinvoice.taxTotal = "10000"

    '[필수] 합계금액, 공급가액 합계 + 세액합계
	newTaxinvoice.totalAmount = "110000"
    
    '기재 상 '일련번호' 항목
    newTaxinvoice.serialNum = "123"

	'기재 상 '권' 항목, 최대값 32767
    newTaxinvoice.kwon = "1"

	'기재 상 '호' 항목, 최대값 32767
    newTaxinvoice.ho = "1"

	'기재 상 '현금' 항목
    newTaxinvoice.cash = ""
    
	'기재 상 '수표' 항목
    newTaxinvoice.chkBill = ""

	'기재 상 '어음' 항목
    newTaxinvoice.note = ""
	
	'기재 상 '외상미수금' 항목
    newTaxinvoice.credit = ""

	'기재 상 '비고'항목
    newTaxinvoice.remark1 = "비고1"
    newTaxinvoice.remark2 = "비고2"
    newTaxinvoice.remark3 = "비고3"

	'사업자등록증 이미지 첨부여부
    newTaxinvoice.businessLicenseYN = False 

	' 통장사본 이미지 첨부여부
    newTaxinvoice.bankBookYN = False         
  
	
	
	'**************************************************************
    '         수정세금계산서 정보 (수정세금계산서 작성시에만 기재
    ' - 수정세금계산서 관련 정보는 연동매뉴얼 또는 개발가이드 링크 참조
    ' - [참고] 수정세금계산서 작성방법 안내 - https://docs.popbill.com/taxinvoice/modify?lang=asp
	'**************************************************************

	' [수정세금계산서 발행시 필수] 수정사유코드, 수정사유에 따라 1~6중 선택기재
    newTaxinvoice.modifyCode = ""

	' [수정세금계산서 발행시 필수] 원본세금계산서의 국세청 승인번호 기재
    newTaxinvoice.orgNTSConfirmNum = ""


	'**************************************************************
	'										상세항목(품목) 정보
	'**************************************************************
    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 1             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20190103"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명1번"
    newDetail.spec = "규격"
    newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.unitCost = "50000"       ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.supplyCost = "50000"
    newDetail.tax = "5000"
    newDetail.remark = "비고"

    newTaxinvoice.AddDetail newDetail

    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 2             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20190103"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명2번"
    newDetail.spec = "규격"
    newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.unitCost = "50000"       ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.supplyCost = "50000"
    newDetail.tax = "5000"
    newDetail.remark = "비고"
    
    newTaxinvoice.AddDetail newDetail
 
	' 즉시요청 메모
	memo = "즉시요청 메모"

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.RegistRequest(testCorpNum, newTaxinvoice, memo, userID)

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
				<legend>세금계산서 즉시요청</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>