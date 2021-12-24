<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 1건의 취소현금영수증을 즉시발행합니다.
	' - 현금영수증 국세청 전송 정책 : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=asp
	' - https://docs.popbill.com/cashbill/asp/api#RevokeRegistIssue
	'**************************************************************

	' 팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 팝빌 회원 아이디
	userID = "testkorea"				 

	' 문서번호, 가맹점 사업자번호 단위 고유번호 할당, 1~24자리 영문,숫자조합으로 중복없이 구성.
	mgtKey = "20211201-001"

	' 원본 현금영수증 국세청승인번호
	orgConfirmNum = "820116333"

	' 원본 현금영수증 거래일자
	orgTradeDate = "20211101"

	' 발행안내 문자 전송여부
	smssendYN = False

	' 메모
	memo = "즉시발행 메모"

	On Error Resume Next

	Set Presponse = m_CashbillService.RevokeRegistIssue(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID)

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
				<legend>취소현금영수증 즉시발행</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>