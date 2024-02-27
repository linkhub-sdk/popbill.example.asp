<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팝빌 인증서버에 등록된 인증서의 만료일을 확인합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/cert#GetCertificateExpireDate
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    On Error Resume Next

    certificateExpiration = m_TaxinvoiceService.GetCertificateExpireDate(CorpNum)

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
                    <ul>
                        <% If code = 0 Then %>
                            <li>만료일시 : <%=certificateExpiration%> </li>
                        <% Else %>
                            <li>Response.code : <%=code%> </li>
                            <li>Response.message : <%=message%> </li>
                        <% End If%>
                    </ul>
            </fieldset>
        </div>
    </body>
</html>
