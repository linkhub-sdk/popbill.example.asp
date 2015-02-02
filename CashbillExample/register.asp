<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	 '팝빌 회원 사업자번호, "-" 제외
	userID = "testkorea"		 '팝빌 회원 아이디
	mgtKey = "20150201-01"		 '연동관리번호, 발행자별 고유번호 할당, 1~24자리 영문,숫자조합으로 중복없이 구성.

	Set CashbillObj = New CashBill

    CashbillObj.mgtKey = mgtKey				   
    CashbillObj.tradeType = "승인거래"				'승인거래 or 취소거래
    CashbillObj.franchiseCorpNum = testCorpNum		'발행자 사업자번호
    CashbillObj.franchiseCorpName = "발행자 상호"
    CashbillObj.franchiseCEOName = "발행자 대표자"
    CashbillObj.franchiseAddr = "발행자 주소"
    CashbillObj.franchiseTEL = "070-1234-1234"
    CashbillObj.identityNum = "01041680206"
    CashbillObj.customerName = "고객명"
    CashbillObj.itemName = "상품명"
    CashbillObj.orderNumber = "주문번호"
    CashbillObj.email = "test@test.com"
    CashbillObj.hp = "111-1234-1234"
    CashbillObj.fax = "777-444-3333"			
    CashbillObj.serviceFee = "0"				'봉사료
    CashbillObj.supplyCost = "10000"			'공급가액
    CashbillObj.tax = "1000"					'부가세
    CashbillObj.totalAmount = "11000"			'거래금액
    CashbillObj.tradeUsage = "소득공제용"       '소득공제용 or 지출증빙용
    CashbillObj.taxationType = "과세"           '과세 or 비과세
    
	CashbillObj.smssendYN = False				'발행시 안내문자 자동전송여부

	On Error Resume Next

	Set Presponse = m_CashbillService.Register(testCorpNum, CashbillObj, UserID)

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
				<legend>팝빌 현금영수증 등록</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>