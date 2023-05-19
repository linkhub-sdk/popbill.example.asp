<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' "임시저장" 상태의 전자명세서를 발행하여, "발행완료" 상태로 처리합니다.
    ' - 팝빌 사이트 [전자명세서] > [환경설정] > [전자명세서 관리] 메뉴의 발행시 자동승인 옵션 설정을 통해
    '   전자명세서를 "발행완료" 상태가 아닌 "승인대기" 상태로 발행 처리 할 수 있습니다.
    ' - 전자명세서 발행 함수 호출시 수신자에게 발행 안내 메일이 발송됩니다.
    ' - https://developers.popbill.com/reference/statement/asp/api/issue#Issue
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외 10자리
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    userID = "testkorea"

    ' 명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
    itemCode = "121"

    ' 문서번호
    mgtKey = "20220720-ASP-002"

    ' 메모
    memo = "전자명세서 발행 테스트"

    On Error Resume Next

    Set result = m_StatementService.Issue(testCorpNum, itemCode, mgtKey, memo, userID)

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
                <legend>전자명세서 발행</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>