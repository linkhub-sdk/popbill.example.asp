<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 전자명세서를 팩스전송합니다.
	' - 팩스 전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
	' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [팩스] > [전송내역]
	'   메뉴에서 전송결과를 확인할 수 있습니다.
	' - https://docs.popbill.com/statement/asp/api#SendFAX
	'**************************************************************

	'팝빌 회원 사업자번호, "-"제외 10자리
	testCorpNum = "1234567890"
	
	'팝빌 회원 아이디
	userID = "testkorea"

	'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
	itemCode = "121"

	'문서번호
	mgtKey = "20190103-001"

	'발신번호
	sender = "07043042991"

	'수신팩스번호
	receiver = "070111222"

	On Error Resume Next

	Set result = m_StatementService.SendFAX(testCorpNum, itemCode, mgtKey, sender, receiver, userID)

	If Err.Number <> 0 Then
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
				<legend>전자명세서 팩스 전송</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>