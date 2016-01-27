<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 

<%
	CorpNum = "1234567890"		 ' 연동회원 사업자번호
	UserID = "testkorea"				 ' 연동회원 아이디 

	Set infoObj = New CorpInfo

	infoObj.ceoname = "링크허브 대표자"
	infoObj.corpName = "링크허브"
	infoObj.addr	= "주소수정"
	infoObj.bizType = "업태정보"
	infoObj.bizClass = "업종정보"
	
	On Error Resume Next

	Set Presponse = m_FaxService.UpdateCorpInfo(CorpNum, infoObj, UserID)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = Presponse.code
		message =Presponse.message
	End If

	On Error GoTo 0

%>

	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>회사정보 수정</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>