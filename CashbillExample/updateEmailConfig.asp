<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 현금영수증 관련 메일전송 항목에 대한 전송여부를 수정합니다.
	' - https://docs.popbill.com/cashbill/asp/api#UpdateEmailConfig
    ' 메일전송유형
    ' CSH_ISSUE : 고객에게 현금영수증이 발행 되었음을 알려주는 메일 입니다.
    ' CSH_CANCEL : 고객에게 현금영수증이 발행취소 되었음을 알려주는 메일 입니다.
	'**************************************************************
	
	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	'팝빌회원 아이디
	userID = "testkorea"		 

	'메일 전송 유형
	emailType = "CSH_ISSUE"

	'전송 여부 (true = 전송, false = 미전송)
	sendYN = true

	On Error Resume Next

	Set Presponse = m_CashbillService.updateEmailConfig(testCorpNum, emailType, sendYN, UserID)

	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
	End If

	On Error GoTo 0 

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>알림메일 전송설정 수정</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>