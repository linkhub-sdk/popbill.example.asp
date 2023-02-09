<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 발신자가 발행한 전자명세서를 발행취소합니다.
    ' - "발행취소" 상태의 전자명세서를 삭제(Delete API) 함수를 이용하면, 전자명세서 관리를 위해 부여했던 문서번호를 재사용 할 수 있습니다.
    ' - https://developers.popbill.com/reference/statement/asp/api/issue#Cancel
    '**************************************************************

    ' 팝빌회원 사업자번호, "-"제외 10자리
    testCorpNum = "1234567890"
    
    ' 팝빌회원 아이디
    userID = "testkorea"

    ' 명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
    itemCode = "121"

    ' 문서관리번호
    mgtKey = "20220720-ASP-001"

    ' 메모
    memo = "전자명세서 발행취소"

    On Error Resume Next

    Set result = m_StatementService.CancelIssue(testCorpNum, itemCode, mgtKey, memo, userID)

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
                <legend>전자명세서 발행취소</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>