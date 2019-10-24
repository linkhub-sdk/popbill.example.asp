<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 현금영수증 관리번호 중복여부를 확인합니다.
	' - 관리번호는 1~24자리로 숫자, 영문 '-', '_' 조합으로 구성할 수 있습니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외 10자리
	testCorpNum = "1234567890"	
	
	'문서번호
	mgtKey = "20190103-001"		
		
	On Error Resume Next

	Presponse = m_CashbillService.CheckMgtKeyInUse(testCorpNum, mgtKey)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
	Else	
		If Presponse = True Then
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