<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팝빌에 등록된 인증서 만료일자를 확인합니다.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/cert#GetCertificateExpireDate
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    expireDate = m_HTTaxinvoiceService.getCertificateExpireDate(CorpNum, UserID)

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
                <legend>공인인증서 만료일시 확인</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>공인인증서 만료일시 : <%=expireDate%> </li>
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