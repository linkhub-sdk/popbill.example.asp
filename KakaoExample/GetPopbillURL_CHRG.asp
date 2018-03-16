<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 연동회원 포인트충전 URL을 반환합니다.
	' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
	'**************************************************************
	
	' 팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	' 팝빌 회원 아이디
	userID = "testkorea"

	' CHRG : 포인트충전 팝업 
	TOGO = "CHRG"

	On Error Resume Next

	url = m_KakaoService.GetPopbillURL(testCorpNum, userID, TOGO)

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
				<legend>연동회원 포인트충전 팝업 URL 요청</legend>
				<% If code = 0 Then %>
					<ul>
						<li>URL : <%=CStr(url)%> </li>
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