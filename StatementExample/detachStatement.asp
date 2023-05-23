<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 전자명세서에 첨부된 다른 전자명세서를 첨부해제합니다.
    ' - https://developers.popbill.com/reference/statement/asp/api/etc#DetachStatement
    '**************************************************************

    ' 팝빌회원 사업자번호
    CorpNum = "1234567890"

    ' 첨부할 명세서 종류코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
    itemCode = 121

    ' 첨부할 문서번호
    mgtKey = "20220720-ASP-001"

    ' 첨부해제할 명세서 종류코드- 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
    subItemCode = 121

    ' 첨부해제할 명세서 문서번호
    subMgtKey = "20220720-100"

    On Error Resume Next

    Set result = m_StatementService.DetachStatement(CorpNum, itemCode, mgtKey, subItemCode, subMgtKey)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = result.code
        message = result.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>다른 전자명세서 첨부해제</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>