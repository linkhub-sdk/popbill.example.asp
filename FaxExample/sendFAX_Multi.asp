<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'팝빌 회원 사업자번호, "-" 제외
	userID = "testkorea"			'팝빌 회원 아이디
	sendNum = "07075103710"			'발신자 번호
	senderName = "발신자명"			'발신자명
'	reserveDT = "20150123200000"	'전송예약시간 yyyyMMddHHmmss, reserveDT값이 null 경우 즉시전송
	
	'수신자목록, 최대 1000건
	Dim receivers(1)
	Set receivers(0) = New FaxReceiver
	receivers(0).receiverNum = "010111222"
	receivers(0).receiverName = "수신자 명칭"

	Set receivers(1) = New FaxReceiver
	receivers(1).receiverNum = "00011112222"
	receivers(1).receiverName = "수신자 명칭"

	FilePaths = Array("C:\popbill.example.asp\대한민국헌법.doc")

	On Error Resume Next

	url = m_FaxService.SendFAX(testCorpNum, sendNum, senderName, receivers, FilePaths,  reserveDT , userID )

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
				<legend>팩스 전송</legend>
				<ul>
					<% If code = 0 Then %>
						<li>recepitNum : <%=url%> </li>
					<% Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>