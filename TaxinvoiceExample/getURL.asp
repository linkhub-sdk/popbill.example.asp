<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 로그인 상태로 팝빌 사이트의 전자세금계산서 문서함 메뉴에 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
    ' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/info#GetURL
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' TBOX(임시문서함), SBOX(매출문서함), PBOX(매입문서함), SWBOX(매출 발행 대기함), PWBOX(매입 발행 대기함), WRITE(정발행 작성)
    TOGO = "SBOX"

    On Error Resume Next

    url = m_TaxinvoiceService.GetURL(CorpNum, UserID, TOGO)

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
                <legend>팝빌 전자세금계산서 문서함 URL</legend>
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
