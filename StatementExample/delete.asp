<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 삭제 가능한 상태의 전자명세서를 삭제합니다.
    ' - 전자명세서를 삭제하면 사용된 문서번호(mgtKey)를 재사용할 수 있습니다.
    ' - 삭제 가능한 상태: "임시저장", "취소", "승인거부", "발행취소"
    ' - https://developers.popbill.com/reference/statement/asp/api/issue#Delete
    '**************************************************************

    ' 팝빌회원 사업자번호
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 명세서 종류코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
    itemCode = "121"

    ' 문서번호
    mgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set result = m_StatementService.Delete(CorpNum, itemCode, mgtKey, UserID)

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
                <legend>전자명세서 삭제</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>