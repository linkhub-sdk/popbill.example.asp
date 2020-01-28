<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
    ' 1건의 취소현금영수증을 임시저장 합니다.
    ' - https://docs.popbill.com/cashbill/asp/api#RevokeRegister
	'**************************************************************

	' 팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 팝빌 회원 아이디
	userID = "testkorea"				 

	' 문서번호, 가맹점 사업자단위 고유번호 할당, 1~24자리 영문,숫자조합으로 중복없이 구성.
	mgtKey = "20190103-001"

	' 원본 현금영수증 국세청승인번호
	orgConfirmNum = "820116333"

	' 원본 현금영수증 거래일자
	orgTradeDate = "20181231"

	' 발행안내 문자 전송여부
	smssendYN = False

	On Error Resume Next

	Set Presponse = m_CashbillService.RevokeRegister(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, userID)

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
				<legend>취소현금영수증 임시저장</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>