<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 홈택스연동 인증관리를 위한 URL을 반환합니다.
    ' - 인증방식에는 부서사용자/공인인증서 인증 방식이 있습니다.
    ' - 반환된 URL은 보안정책에 따라 30초의 유효시간을 갖습니다.
    ' - https://docs.popbill.com/htcashbill/asp/api#GetCertificatePopUpURL
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		

    ' 팝빌회원 아이디
    userID = "testkorea"					

    On Error Resume Next

    url = m_HTCashbillService.GetCertificatePopUpURL(testCorpNum, userID)

    If Err.Number <> 0 then
        code = Err.Number
        message =  Err.Description
        Err.Clears
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>홈택스 인증정보 관리 팝업 URL</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>URL : <%=url%> </li>
                    </ul>
                <%	Else  %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>	
                <%	End If	%>
            </fieldset>
         </div>
    </body>
</html>