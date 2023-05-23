<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' "임시저장" 상태의 명세서에 1개의 파일을 첨부합니다. (최대 5개)
    ' - https://developers.popbill.com/reference/statement/asp/api/etc#AttachFile
    '**************************************************************

    ' 팝빌회원 사업자번호, "-"제외 10자리
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 명세서 종류코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
    itemCode = "121"

    ' 문서관리번호
    mgtKey = "20220720-ASP-001"

    ' 첨부 파일 경로
    filePath = "C:\popbill.example.asp\Popbill\로고.gif"

    On Error Resume Next

    Set result = m_StatementService.AttachFile(CorpNum, itemCode, mgtKey, filePath, UserID)

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
                <legend>전자명세서 파일첨부</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>
