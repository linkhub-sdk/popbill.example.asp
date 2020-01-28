<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'***************************************************
	' 팝빌에 등록된 전자세금계산서 부서사용자 계정정보를 삭제합니다.
	' - https://docs.popbill.com/httaxinvoice/asp/api#DeleteDeptUser
	'***************************************************

	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	' 팝빌회원 아이디
	userID = "testkorea"

	On Error Resume Next

	Set result = m_HTTaxinvoiceService.DeleteDeptUser(testCorpNum, userID)

	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = result.code
		message = result.message
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>부서사용자 등록정보 삭제</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>