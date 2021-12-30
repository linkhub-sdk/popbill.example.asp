<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 팝빌 현금영수증 문서함 관련 팝업 URL을 반환합니다.
    ' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
    ' - https://docs.popbill.com/cashbill/asp/api#GetURL
    '**************************************************************

    '팝빌 회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '팝빌 회원 아이디
    userID = "testkorea"

    'TBOX(임시문서함), PBOX(발행문서함), WRITE(현금영수증 작성)
    TOGO = "PBOX"

    On Error Resume Next

    url = m_CashbillService.GetURL(testCorpNum, userID, TOGO)

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
                <legend>팝빌 현금영수증 문서함 URL</legend>
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