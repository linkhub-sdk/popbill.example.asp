<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 해당 사업자의 파트너 연동회원 가입여부를 확인합니다.
	' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
	'**************************************************************

	' 사업자번호 ("-"제외)
	testCorpNum = "1231212312"		
		
	On Error Resume Next

	Set Presponse = m_HTCashbillService.CheckIsMember(testCorpNum,LinkID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = Presponse.code
		message =Presponse.message
	End If

	On Error GoTo 0
%>

	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>연동회원사 가입여부 확인</legend>
				<ul>
					<li>Response.code : <%=CStr(code)%></li>
					<li>Response.message : <%=message%></li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>