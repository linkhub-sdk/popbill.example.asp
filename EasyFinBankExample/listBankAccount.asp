<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 팝빌에 등록된 은행계좌 목록을 확인합니다.
	' - https://docs.popbill.com/easyfinbank/asp/api#ListBankAccount
	'**************************************************************

	''팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	'팝빌회원 아이디
	UserID = "testkorea"
	
	On Error Resume Next

	Set result = m_EasyFinBankService.ListBankAccount(testCorpNum, UserID)
	
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
				<legend>계좌 목록</legend>
				<%
					If code = 0 Then
						For i=0 To result.Count-1
				%>
							<fieldset class="fieldset2">					
								<legend>ListBankAccount [ <%=i+1%> / <%=result.Count%> ] </legend>
									<ul>
										<li> accountNumber (계좌번호) : <%=result.Item(i).accountNumber%></li>
										<li> bankCode (은행코드) : <%=result.Item(i).bankCode%></li>
										<li> accountName (계좌 별칭) : <%=result.Item(i).accountName%></li>
										<li> accountType (계좌유형) : <%=result.Item(i).accountType%></li>
										<li> state (정액제 상태) : <%=result.Item(i).state%></li>
										<li> regDT (등록일시) : <%=result.Item(i).regDT%></li>
										<li> memo (메모) : <%=result.Item(i).memo%></li>
									</ul>
								</fieldset>
				<%
						Next
					Else
				%>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	
					End If
				%>
			</fieldset>
		 </div>
	</body>
</html>
