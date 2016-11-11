<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 1건의 현금영수증을 삭제합니다.
	' - 현금영수증을 삭제하면 사용된 문서관리번호(mgtKey)를 재사용할 수 있습니다.
	' - 삭제가능한 문서 상태 : [임시저장], [발행취소]
	'**************************************************************
	
	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	'팝빌회원 아이디
	userID = "testkorea"		 

	'문서관리번호
	mgtKey = "20150201-01"		 

	On Error Resume Next

	Set Presponse = m_CashbillService.Delete(testCorpNum, mgtKey, UserID)

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
				<legend>현금영수증 삭제</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>