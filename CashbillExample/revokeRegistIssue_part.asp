<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 1건의 (부분) 취소현금영수증을 즉시발행합니다.
	' - 현금영수증 국세청 전송 정책 : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=asp
	' - https://docs.popbill.com/cashbill/asp/api#RevokeRegistIssue_Part
	'**************************************************************

	' 팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 팝빌 회원 아이디
	userID = "testkorea"				 

	' 문서번호, 가맹점 사업자번호 단위 고유번호 할당, 1~24자리 영문,숫자조합으로 중복없이 구성.
	mgtKey = "20171115-01"

	' 원본 현금영수증 국세청승인번호
	orgConfirmNum = "820116333"

	' 원본 현금영수증 거래일자
	orgTradeDate = "20170711"

	' 발행안내 문자 전송여부
	smssendYN = False

	' 메모
	memo = "즉시발행 메모"
	
	'부분취소여부, True-부분취소, False-전체취소
	isPartCancel = True

	'취소사유, 1-거래취소, 2-오류발급 취소, 3-기타
	cancelType = 1

	'[취소] 공급가액
	supplyCost = "5000"

	'[취소] 세액
	tax = "500"

	'[취소] 봉사료
	serviceFee = "0"
	
	'[취소] 합계금액
	totalAmount = "5500"

	On Error Resume Next

	Set Presponse = m_CashbillService.RevokeRegistIssue_Part(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID, _
		isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount)

	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
		confirmNum = Presponse.confirmNum
		tradeDate = Presponse.tradeDate		
	End If

	On Error GoTo 0 

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>(부분) 취소현금영수증 즉시발행</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
					<li> Response.confirmNum : <%=confirmNum%> </li>
					<li> Response.tradeDate : <%=tradeDate%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>