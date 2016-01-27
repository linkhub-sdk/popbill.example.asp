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

	Set contInfo = New ContactInfo
	
	contInfo.personName = "ASPTest"
	contInfo.tel = "010-1234-1234"
	contInfo.hp = "010-1234-1234"
	contInfo.email = "code@linkhub.co.kr"
	contInfo.fax = "02-6442-9700"
	contInfo.searchAllAllowYN = True
	contInfo.mgrYN = True

	On Error Resume Next

	Set Presponse = m_FaxService.UpdateContact(CorpNum, contInfo, UserID)
	
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
				<legend>담당자 정보수정</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>