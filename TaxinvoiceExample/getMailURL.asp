<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 전자세금계산서 안내메일의 상세보기 링크 URL을 반환합니다.
    ' 함수 호출로 반환 받은 URL에는 유효시간이 없습니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/view#GetMailURL
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType= "SELL"

    ' 문서번호
    MgtKey = "20220720-ASP-001"

    On Error Resume Next

    url = m_TaxinvoiceService.GetMailURL(CorpNum, KeyType, MgtKey, UserID)

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
                <legend>세금계산서 메일링크 URL</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>URL : <%=url%> </li>
                    <% Else %>
                        <li> Response.code : <%=code%> </li>
                        <li> Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>
