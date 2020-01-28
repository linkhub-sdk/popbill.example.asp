<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 파트너의 잔여포인트를 확인합니다.
	' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를
	'   이용하시기 바랍니다.
	' - https://docs.popbill.com/htcashbill/asp/api#GetPartnerBalance
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		 
	
	On Error Resume Next

	remainPoint = m_HTCashbillService.getPartnerBalance(testCorpNum)

	If Err.Number <> 0 then
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
				<legend>파트너 잔여포인트 확인결과</legend>
				<% If code = 0 Then %>
					<ul>
						<li>잔여포인트 : <%=CStr(remainpoint)%> </li>
					</ul>
				<%	Else  %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>