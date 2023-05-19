<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원의 국세청 전송 옵션 설정 상태를 확인합니다.
    ' - 팝빌 국세청 전송 정책 [https://developers.popbill.com/guide/taxinvoice/asp/introduction/policy-of-send-to-nts]
    ' - 국세청 전송 옵션 설정은 팝빌 사이트 [전자세금계산서] > [환경설정] > [세금계산서 관리] 메뉴에서 설정할 수 있으며, API로 설정은 불가능 합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#GetSendToNTSConfig
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    On Error Resume Next

    ntsConfig = m_TaxinvoiceService.GetSendToNTSConfig(testCorpNum)

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
                <legend> 국세청 전송 설정 확인</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>ntsConfig : <%=ntsConfig%></li>
                        <li>(True)-발행 즉시 전송 (False)-익일 자동 전송</li>
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