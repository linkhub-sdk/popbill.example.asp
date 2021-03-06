<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 팩스를 재전송합니다.
	' - 접수일로부터 60일이 경과되지 않은 건만 재전송 가능합니다.
	' - 팩스 재전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
	' - https://docs.popbill.com/fax/asp/api#ResendFAX
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		

	'팝빌 회원 아이디
	userID = "testkorea"			
	
	'팩스 접수번호 
	receiptNum = "019010315075200001"
	
	'발신자 번호
	sendNum = "07043042991"
	
	'발신자명
	sendName = "발신자명"

	'전송예약시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
	reserveDT = ""	

	'팩스 제목
	title = "팩스 재전송"
	
	'수신정보가 기존전송정보와 동일한 경우
	ReDim receivers(-1)


	'수신정보가 기존전송정보 다를 경우 아래 코드 참조	
'	Dim receivers(0)
'	Set receivers(0) = New FaxReceiver
	
	'수신번호
'	receivers(0).receiverNum = "07066666"

	'수신자명
'	receivers(0).receiverName = "수신자 명칭"

	'전송요청번호 (팝빌 회원별 비중복 번호 할당)
	'영문,숫자,'-','_' 조합, 최대 36자
	requestNum = ""		

	On Error Resume Next

	url = m_FaxService.ResendFAX(testCorpNum, receiptNum, sendNum, senderName, receivers, reserveDT , userID, title, requestNum)

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
				<legend>팩스 재전송</legend>
				<ul>
					<% If code = 0 Then %>
						<li>recepitNum (접수번호) : <%=url%> </li>
					<% Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>