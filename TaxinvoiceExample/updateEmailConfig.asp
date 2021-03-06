<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************'
	' 전자세금계산서 관련 메일전송 항목에 대한 전송여부를 수정한다.
	' - https://docs.popbill.com/taxinvoice/asp/api#UpdateEmailConfig
	'
	' 메일전송유형
	' [정발행]
	' TAX_ISSUE : 공급받는자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
	' TAX_ISSUE_INVOICER : 공급자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
	' TAX_CHECK : 공급자에게 전자세금계산서가 수신확인 되었음을 알려주는 메일입니다.
	' TAX_CANCEL_ISSUE : 공급받는자에게 전자세금계산서가 발행취소 되었음을 알려주는 메일입니다.
	'
	' [발행예정]
	' TAX_SEND : 공급받는자에게 [발행예정] 세금계산서가 발송 되었음을 알려주는 메일입니다.
	' TAX_ACCEPT : 공급자에게 [발행예정] 세금계산서가 승인 되었음을 알려주는 메일입니다.
	' TAX_ACCEPT_ISSUE : 공급자에게 [발행예정] 세금계산서가 자동발행 되었음을 알려주는 메일입니다.
	' TAX_DENY : 공급자에게 [발행예정] 세금계산서가 거부 되었음을 알려주는 메일입니다.
	' TAX_CANCEL_SEND : 공급받는자에게 [발행예정] 세금계산서가 취소 되었음을 알려주는 메일입니다.
	'
	' [역발행]
	' TAX_REQUEST : 공급자에게 세금계산서를 전자서명 하여 발행을 요청하는 메일입니다.
	' TAX_CANCEL_REQUEST : 공급받는자에게 세금계산서가 취소 되었음을 알려주는 메일입니다.
	' TAX_REFUSE : 공급받는자에게 세금계산서가 거부 되었음을 알려주는 메일입니다.
	'
	' [위수탁발행]
	' TAX_TRUST_ISSUE : 공급받는자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
	' TAX_TRUST_ISSUE_TRUSTEE : 수탁자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
	' TAX_TRUST_ISSUE_INVOICER : 공급자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
	' TAX_TRUST_CANCEL_ISSUE : 공급받는자에게 전자세금계산서가 발행취소 되었음을 알려주는 메일입니다.
	' TAX_TRUST_CANCEL_ISSUE_INVOICER : 공급자에게 전자세금계산서가 발행취소 되었음을 알려주는 메일입니다.
	'
	' [위수탁 발행예정]
	' TAX_TRUST_SEND : 공급받는자에게 [발행예정] 세금계산서가 발송 되었음을 알려주는 메일입니다.
	' TAX_TRUST_ACCEPT : 수탁자에게 [발행예정] 세금계산서가 승인 되었음을 알려주는 메일입니다.
	' TAX_TRUST_ACCEPT_ISSUE : 수탁자에게 [발행예정] 세금계산서가 자동발행 되었음을 알려주는 메일입니다.
	' TAX_TRUST_DENY : 수탁자에게 [발행예정] 세금계산서가 거부 되었음을 알려주는 메일입니다.
	' TAX_TRUST_CANCEL_SEND : 공급받는자에게 [발행예정] 세금계산서가 취소 되었음을 알려주는 메일입니다.
	'
	' [처리결과]
	' TAX_CLOSEDOWN : 거래처의 휴폐업 여부를 확인하여 안내하는 메일입니다.
	' TAX_NTSFAIL_INVOICER : 전자세금계산서 국세청 전송실패를 안내하는 메일입니다.
	'
	' [정기발송]
	' TAX_SEND_INFO : 전월 귀속분 [매출 발행 대기] 세금계산서의 발행을 안내하는 메일입니다.
	' ETC_CERT_EXPIRATION : 팝빌에서 이용중인 공인인증서의 갱신을 안내하는 메일입니다.
	'**************************************************************
	
	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	'팝빌회원 아이디
	userID = "testkorea"		 

	'메일 전송 유형
	emailType = "TAX_ISSUE"

	'전송 여부 (true = 전송, false = 미전송)
	sendYN = true

	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.updateEmailConfig(testCorpNum, emailType, sendYN, UserID)

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