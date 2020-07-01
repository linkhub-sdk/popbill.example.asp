<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
		<title>예금주조회 API SDK ASP Example.</title>
	</head>
	<!--#include file="common.asp"--> 
	<%
		'**************************************************************
		' 1건의 계좌정보에 대한 예금주정보를 조회합니다.
		'**************************************************************
		'팝빌회원 사업자번호
		CorpNum = "1234567890"	

		'팝빌회원 아이디
		UserID = "testkorea"
		
		'기관코드
		BankCode = "0004"

		'계좌번호
		AccountNumber = "94324511758125"

		On Error Resume Next
			Set result = m_AccountCheckService.checkAccountInfo(CorpNum, BankCode, AccountNumber, UserID)
			
			If Err.Number <> 0 Then
				code = Err.Number
				message = Err.Description
				Err.Clears
			End If
		On Error GoTo 0


	%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>예금주조회</legend>
			<%
				If Not IsEmpty(result) Then  

			%>

				<ul>
					<li>bankCode (기관코드) : <%= result.bankCode%></li>	
					<li>accountNumber (계좌번호) : <%= result.accountNumber%></li>	
					<li>accountName (예금주 성명) : <%= result.accountName%></li>	
					<li>checkDate (확인일시) : <%= result.checkDate%></li>	
					<li>resultCode (응답코드) : <%= result.resultCode%></li>	
					<li>resultMessage (응답메시지) : <%= result.resultMessage%></li>	
				</ul>

			<%
				End If 
				If Not IsEmpty(code) then
			%>

			<ul>	
				<li>Response.code : <%= code %> </li>
				<li>Response.message : <%= message %></li>
			</ul>
			<%
				End If
			%>		

			</fieldset>
	</body>
</html>