<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 1건의 전자세금계산서 (구)양식 인쇄팝업 URL을 반환합니다.
    ' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
    ' - https://docs.popbill.com/taxinvoice/asp/api#GetOldPrintURL
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외 10자리
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    userID = "testkorea"

    ' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType = "SELL"

    ' 문서번호 
    MgtKey = "20210601ASP001"
    
    On Error Resume Next

    url = m_TaxinvoiceService.GetOldPrintURL(testCorpNum, KeyType, MgtKey, userID)

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
                <legend>세금계산서 (구)양식 인쇄 URL </legend>
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