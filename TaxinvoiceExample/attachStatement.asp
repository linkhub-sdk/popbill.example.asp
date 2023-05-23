<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    '  팝빌 전자명세서 API를 통해 발행한 전자명세서를 세금계산서에 첨부합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#AttachStatement
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 세금계산서 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType= "SELL"

    ' 세금계산서 문서번호
    MgtKey = "20220720-ASP-002"

    ' 첨부할 전자명세서 종류코드
    ' - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
    SubItemCode = 121

    ' 전자명세서 관리번호
    SubMgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.AttachStatement(CorpNum, KeyType, MgtKey, SubItemCode, SubMgtKey)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message = Presponse.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>전자명세서 첨부</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message : <%=message%> </li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>