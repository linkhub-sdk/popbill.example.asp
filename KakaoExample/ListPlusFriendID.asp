<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 카카오톡 플러스친구 계정 목록을 확인합니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		 
	
	On Error Resume Next

	Set Presponse = m_KakaoService.ListPlusFriendID(testCorpNum)

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
				<legend>카카오톡 플러스친구 계정 목록 확인</legend>
				<%
					For i=0 To Presponse.length -1
				%>
				<fieldset class="fieldset2">
				<ul>
					<li>플러스친구 아이디 (plusFriendID) : <%=Presponse.Get(i).plusFriendID%> </li>
					<li>플러스친구 이름(plusFriendName) : <%=Presponse.Get(i).plusFriendName%> </li>
					<li>등록일시(regDT) : <%=Presponse.Get(i).regDT%> </li>
				</ul>
				</fieldset>
				<%
					Next
				%>

		 </div>
	</body>
</html>