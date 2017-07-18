<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	' 팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 팝빌 회원 아이디
	userID = "testkorea"				 


	' 문서관리번호, 발행자별 고유번호 할당, 1~24자리 영문,숫자조합으로 중복없이 구성.
	mgtKey = "20170718-04"

	' 메모
	memo = "즉시발행 메모"

	' 현금영수증 객체 생성
	Set CashbillObj = New CashBill
	
    CashbillObj.mgtKey = mgtKey				   

    '현금영수증 형태, [승인거래, 취소거래] 중 기재
    CashbillObj.tradeType = "취소거래"

	'[취소거래시 필수] 원본 현금영수증 국세청승인번호
	CashbillObj.orgConfirmNum = "820116333"

	'[취소거래시 필수] 원본 현금영수증 거래일자
	CashbillObj.orgTradeDate = ""

	'발행자 사업자번호, "-" 제외 10자리
	CashbillObj.franchiseCorpNum = testCorpNum	

    '발행자 상호명
    CashbillObj.franchiseCorpName = "발행자 상호"

    '발행자 대표자 성명
	CashbillObj.franchiseCEOName = "발행자 대표자"
    
    '발행자 주소
	CashbillObj.franchiseAddr = "발행자 주소"
    
    '발행자 연락처
	CashbillObj.franchiseTEL = "070-1234-1234"
    
    '거래처 식별번호, 거래유형에 따라 작성
    '소득공제용 - 주민등록/휴대폰/카드번호 기재가능
    '지출증빙용 - 사업자번호/주민등록/휴대폰/카드번호 기재가능	
	CashbillObj.identityNum = "0101112222"

    '거래유형, [소득공제용, 지출증빙용] 중 기재
	CashbillObj.tradeUsage = "소득공제용"
    
    '과세형태, [과세, 비과세] 중 기재
	CashbillObj.taxationType = "과세" 
	
	'공급가액
	CashbillObj.supplyCost = "10000"			

	'세액
	CashbillObj.tax = "1000"					

	'봉사료
	CashbillObj.serviceFee = "0"				
    
	'합계금액, 공급가액 + 봉사료 + 세액
	CashbillObj.totalAmount = "11000"			
    
	
    '주문고객명
	CashbillObj.customerName = "고객명"
    
	'상품명	
	CashbillObj.itemName = "상품명"
    
	'주문번호
	CashbillObj.orderNumber = "주문번호"
    
	'고객 메일주소
	CashbillObj.email = "test@test.com"
    
	'고객 휴대폰번호
	CashbillObj.hp = "111-1234-1234"
    
	'고객 팩스번호
	CashbillObj.fax = "777-444-3333"			
    
	'발행안내문자 전송여부
	'안내문자 전송시 포인트가 차감되며, 전송실패시 환불처리됩니다.
	CashbillObj.smssendYN = False

	On Error Resume Next

	Set Presponse = m_CashbillService.RegistIssue(testCorpNum, CashbillObj, memo, userID)

	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
	End If

	On Error GoTo 0 

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>현금영수증 즉시발행</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>