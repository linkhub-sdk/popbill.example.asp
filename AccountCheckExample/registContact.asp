<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"-->
<%
	'**************************************************************
	' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
	' - https://developers.popbill.com/reference/accountcheck/asp/api/member#RegistContact
	'**************************************************************

	' 팝빌회원 사업자번호
	CorpNum = "1234567890"

	' 팝빌회원 아이디
	UserID = "testkorea"

	' 담당자 정보 객체 생성
	Set contInfo = New ContactInfo

	' 담당자 아이디, 6자이상 20자미만
	contInfo.id = "testkorea00000"

	' 비밀번호 (8자이상 20자 이하) 영문, 숫자, 특수문자 조합
	contInfo.Password = "asdf1234!@#$"

	' 담당자명
	contInfo.personName = "ASPTest"

	' 연락처
	contInfo.tel = ""

	' 메일주소
	contInfo.email = ""

	' 담당자 조회권한 1 - 개인권한 / 2 - 읽기권한  / 3 - 회사권한
	contInfo.searchRole = 3

	On Error Resume Next

	Set Presponse = m_AccountCheckService.RegistContact(CorpNum, contInfo, UserID)

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
        		<legend>담당자 추가</legend>
        		<ul>
            		<li>Response.code : <%=code%> </li>
            		<li>Response.message: <%=message%> </li>
        		</ul>
    		</fieldset>
		</div>
	</body>
</html>