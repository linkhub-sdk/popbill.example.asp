<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 세금계산서 문서번호 중복여부를 확인합니다.
	' - 1~24자리로 숫자, 영문 '-', '_' 조합으로 사업자별로 중복되지 않도록 구성합니다.
	' - https://docs.popbill.com/taxinvoice/asp/api#CheckMgtKeyInUse
	'**************************************************************
	
	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	' 팝빌회원 아이디
	testUserID = "testkorea"

	' 문서번호
	mgtKey = "20190103-001"

	' 발행형태, (SELL-매출) (BUY-매입) (TRUSTEE-위수탁)
	keyType = "SELL"

	On Error Resume Next
	checkMgtKeyInUse = m_TaxinvoiceService.CheckMgtKeyInUse(testCorpNum, keyType, mgtKey)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
	Else	
		If checkMgtKeyInUse = True Then
			code = 1
			message = "사용중"
		Else
			code = 0 
			message = "미사용중"
		End If
	End If 

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>문서번호 사용여부 확인</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>