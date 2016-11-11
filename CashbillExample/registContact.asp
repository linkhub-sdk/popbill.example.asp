<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 담당자를 신규로 등록합니다.
	'**************************************************************

	' 팝빌회원 사업자번호
	CorpNum = "1234567890"		 

	' 팝빌회원 아이디 
	UserID = "testkorea"				 

	' 담당자 정보 객체 생성
	Set contInfo = New ContactInfo

	' 담당자 아이디, 6자이상 20자미만
	contInfo.id = "testkorea00000"

	' 비밀번호
	contInfo.pwd = "testkorea1234"

	' 담당자명
	contInfo.personName = "ASPTest"

	' 연락처
	contInfo.tel = "010-1234-1234"

	' 휴대폰번호
	contInfo.hp = "010-1234-1234"

	' 메일주소
	contInfo.email = "dev@linkhub.co.kr"

	' 팩스 번호
	contInfo.fax = "02-6442-9700"

	' 회사조회 권한여부, True-회사조회 / False -개인조회
	contInfo.searchAllAllowYN = True
	
	' 관리자 권한여부
	contInfo.mgrYN = False

	On Error Resume Next

	Set Presponse = m_CashbillService.RegistContact(CorpNum, contInfo, UserID)
	
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
				<legend>담당자 추가등록</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>