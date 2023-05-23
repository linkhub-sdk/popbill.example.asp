<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팝빌 사이트와 동일한 세금계산서 1건의 상세정보 페이지(사이트 상단, 좌측 메뉴 및 버튼 제외)의 팝업 URL을 반환합니다.
    ' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/view#GetViewURL
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외 10자리
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType = "SELL"

    ' 문서번호
    MgtKey = "20220720-ASP-001"

    On Error Resume Next

    url = m_TaxinvoiceService.GetViewURL(CorpNum, KeyType, MgtKey, UserID)

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
                <legend>세금계산서 보기 팝업 URL</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li>URL : <%=url%> </li>
                    </ul>
                <% Else %>
                    <ul>
                        <li> Response.code : <%=code%> </li>
                        <li> Response.message : <%=message%> </li>
                    </ul>
                <% End If %>
            </fieldset>
        </div>
    </body>
</html>