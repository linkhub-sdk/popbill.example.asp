<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp" -->
<%
    '**************************************************************
    ' 환불 가능한 포인트를 확인합니다. (보너스 포인트는 환불가능포인트에서 제외됩니다.)
    ' - https://developers.popbill.com/reference/bizinfocheck/asp/api/point#GetRefundableBalance
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

	refundableBalance = m_BizInfoCheckService.GetRefundableBalance(CorpNum, UserID)

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
                <legend>환불 가능 포인트 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> refundableBalance (환불 가능 포인트) : <%=refundableBalance%></li>
                    </ul>
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