<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"-->
<%
	'**************************************************************
	' 가입된 연동회원의 탈퇴를 요청합니다.
	' - 회원탈퇴 신청과 동시에 팝빌의 모든 서비스 이용이 불가하며, 관리자를 포함한 모든 담당자 계정도 일괄탈퇴 됩니다.
	' - 회원탈퇴로 삭제된 데이터는 복원이 불가능합니다.
	' - 관리자 계정만 회원탈퇴가 가능합니다.
	' - https://developers.popbill.com/reference/statement/asp/api/member#QuitMember
	'**************************************************************

	'팝빌회원 사업자번호, "-" 제외
	CorpNum = "1234567890"

	'탈퇴 사유
	QuitReason = "탈퇴사유"

	'팝빌회원 아이디
	UserID = "testkorea"

	On Error Resume Next

	Set result = m_StatementService.QuitMember(CorpNum, QuitReason, UserID)

	If Err.Number <> 0 Then
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
        		<legend>팝빌 회원 탈퇴</legend>
        		<%
            		If code = 0 Then
        		%>
            		<fieldset class="fieldset2">
                		<legend>팝빌 회원 탈퇴</legend>
                    		<ul>
                        		<li> code (응답 코드) : <%=result.code%></li>
                        		<li> message (응답 메시지) : <%=result.message%></li>
                    		</ul>
                		</fieldset>
        		<%
            		Else
        		%>
            		<ul>
                		<li>Response.code: <%=code%> </li>
                		<li>Response.message: <%=message%> </li>
            		</ul>
        		<%
            		End If
        		%>
    		</fieldset>
		</div>
	</body>
</html>