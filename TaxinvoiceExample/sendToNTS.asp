<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' [발행완료] 상태의 세금계산서를 국세청으로 즉시전송합니다.
	' - 국세청 즉시전송을 호출하지 않은 세금계산서는 발행일 기준 익일 오후 3시에
	'   팝빌 시스템에서 일괄적으로 국세청으로 전송합니다.
	' - 익일전송시 전송일이 법정공휴일인 경우 다음 영업일에 전송됩니다.
	' - 국세청 전송에 관한 사항은 "[전자세금계산서 API 연동매뉴얼] > 1.3 국세청
	'   전송 정책" 을 참조하시기 바랍니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 팝빌회원 아이디
	testUserID = "testkorea"   
	 
	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType = "SELL"             

	' 문서번호 
	MgtKey = "20190103-001"      

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.SendToNTS(testCorpNum, KeyType ,MgtKey, testUserID)
	
	If Err.Number <> 0 Then
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
				<legend>국세청 즉시전송</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>