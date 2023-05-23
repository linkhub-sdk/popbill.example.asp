<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 전자세금계산서 유통사업자의 메일 목록을 확인합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#GetEmailPublicKeys
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외 10자리
    CorpNum = "1234567890"

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.GetEmailPublicKeys(CorpNum)

    If Err.Number <> 0 then
        Response.Write("Error Number -> " & Err.Number)
        Response.write("<BR>Error Source -> " & Err.Source)
        Response.Write("<BR>Error Desc   -> " & Err.Description)
        Err.Clears
        Response.end
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>유통사업자 이메일 목록 확인 </legend>
                <ul>
                <%
                    For i=0 To Presponse.length -1
                %>
                        <li> <%=Presponse.Get(i).email%></li>
                <%
                    Next
                %>
                </ul>
        </div>
    </body>
</html>