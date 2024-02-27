<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
    	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
    	<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
    	<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"-->
<%
	'**************************************************************
	' 예금주 성명 조회시 과금되는 포인트 단가를 확인합니다.
	' - https://developers.popbill.com/reference/accountcheck/asp/api/point#GetUnitCost
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	CorpNum = "1234567890"

	' 팝빌회원 아이디
	UserID = "testkorea"

	' 서비스 유형 ("성명" / "실명" 중 택 1 , 성명 = 예금주성명조회, 실명 = 예금주실명조회)
	serviceType = ""

	On Error Resume Next

	unitCost = m_AccountCheckService.GetUnitCost(CorpNum, serviceType, UserID)

	If Err.Number <> 0 then
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
            	<legend>예금주조회 단가 확인 </legend>
            	<ul>
                	<% If code = 0 Then %>
                    	<li>조회단가 : <%=unitCost%> </li>
                	<% Else %>
                    	<li> Response.code : <%=code%> </li>
                    	<li> Response.message : <%=message%> </li>
                	<% End If %>
            	</ul>
        	</fieldset>
    	</div>
	</body>
</html>
