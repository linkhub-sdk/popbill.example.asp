<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 팝빌에 등록되어 있는 공인인증서의 만료일자를 확인합니다.
	' - 공인인증서가 갱신/재발급/비밀번호 변경이 되는 경우 해당 인증서를
	'   재등록 하셔야 정상적으로 API를 이용하실 수 있습니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		 
	
	On Error Resume Next

	certificateExpiration = m_TaxinvoiceService.GetCertificateExpireDate(testCorpNum)

	If Err.Number <> 0 then
		code = Err.Number
		message =  Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>공인인증서 만료일시 확인</legend>
					<ul>
						<% If code = 0 Then %>
							<li>만료일시 : <%=certificateExpiration%> </li>
						<% Else %>
							<li>Response.code : <%=code%> </li>
							<li>Response.message : <%=message%> </li>
						<% End If%>
					</ul>
			</fieldset>
		 </div>
	</body>
</html>