<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' [임시저장] 상태의 세금계산서를 [발행]처리 합니다.
	' - 발행(Issue API)를 호출하는 시점에서 포인트가 차감됩니다.
	' - [발행완료] 세금계산서는 연동회원의 국세청 전송설정에 따라
	'   익일/즉시전송 처리됩니다. 기본설정(익일전송)
	' - 국세청 전송설정은 "팝빌 로그인" > [전자세금계산서] > [환경설정] >
	'   [전자세금계산서 관리] > [국세청 전송 및 지연발행 설정] 탭에서
	'   확인할 수 있습니다.
	' - 국세청 전송정책에 대한 사항은 "[전자세금계산서 API 연동매뉴얼] >
	'   1.4. 국세청 전송 정책" 을 참조하시기 바랍니다
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	' 팝빌회원 아이디
	testUserID = "testkorea"

	' 세금계산서 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"

	' 문서관리번호 
	MgtKey = "20190227-023"
	
	' 메모
	Memo = "발행 메모"

	' 발행 안내메일 제목, 미기재시 기본양식으로 전송
	EmailSubject = ""
	
	' 지연발행 강제여부, 기본값 - False
    ' 발행마감일이 지난 세금계산서를 발행하는 경우, 가산세가 부과될 수 있습니다.
    ' 지연발행 세금계산서를 신고해야 하는 경우 forceIssue 값을 True로 
	' 선언하여 발행(Issue API)을 호출할 수 있습니다.
	ForceIssue = False

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.Issue(testCorpNum, KeyType ,MgtKey, Memo ,EmailSubject, ForceIssue, testUserID)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		ntsConfirmNum = ""
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
		ntsConfirmNum = Presponse.ntsConfirmNum
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>세금계산서 발행</legend>
				<ul>
					<li>응답코드 (Response.code) : <%=code%> </li>
					<li>응답메시지 (Response.message) : <%=message%> </li>
					<% If ntsConfirmNum <> "" Then %>
					<li>국세청승인번호 (Response.ntsConfirmNum) : <%=ntsConfirmNum%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>