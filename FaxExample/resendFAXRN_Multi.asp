<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 전송요청번호(requestNum)을 할당한 팩스를 재전송합니다.
    ' - 접수일로부터 60일이 경과된 경우 재전송할 수 없습니다.
	' - 팩스 재전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
	' - https://docs.popbill.com/fax/asp/api#ResendFAXRN
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		

	'팝빌 회원 아이디
	userID = "testkorea"			
	
	'원본 팩스 전송시 할당한 전송요청번호(requestNum)
	orgRequestNum = "1"
	
	'발신자 번호
	sendNum = "07043042991"		
	
	'발신자명
	sendName = "발신자명"

	'전송예약시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
	reserveDT = ""	

	'팩스제목
	title = "팩스 동보 재전송"
	
	'수신정보가 기존전송정보와 동일한 경우
'	ReDim receivers(-1)
	

	'수신정보가 기존전송정보 다를 경우 아래 코드 참조	
	Dim receivers(1)
	Set receivers(0) = New FaxReceiver
	receivers(0).receiverNum = "010111222"
	receivers(0).receiverName = "수신자 명칭"

	Set receivers(1) = New FaxReceiver
	receivers(1).receiverNum = "000111222"
	receivers(1).receiverName = "수신자 명칭"
	
	'전송요청번호 (팝빌 회원별 비중복 번호 할당)
	'영문,숫자,'-','_' 조합, 최대 36자
	requestNum = ""		

	On Error Resume Next

	url = m_FaxService.ResendFAXRN(testCorpNum, orgRequestNum, sendNum, sendName, receivers, reserveDT, userID, title, requestNum)

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