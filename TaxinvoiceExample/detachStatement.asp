<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'팝빌회원 사업자번호, "-" 제외
	KeyType= "SELL"						'발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	MgtKey = "20160126-55"			'세금계산서 연동관리번호 
	SubItemCode = 121					'첨부해제할 전자명세서 종류 코드
	SubMgtKey = "20160126-54"		'첨부해제할 전자명세서 관리번호

	On Error Resume Next
	Set Presponse = m_TaxinvoiceService.DetachStatement(testCorpNum, KeyType, MgtKey, SubItemCode, SubMgtKey)
	
	If Err.Number <> 0 Then
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
				<legend>전자명세서 첨부해제</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>