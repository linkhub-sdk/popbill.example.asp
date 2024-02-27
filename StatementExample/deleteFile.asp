<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' "임시저장" 상태의 전자명세서에 첨부된 1개의 파일을 삭제합니다.
    ' - 파일 식별을 위해 첨부시 부여되는 'FileID'는 첨부파일 목록 확인(GetFiles API) 함수를 호출하여 확인합니다.
    ' - https://developers.popbill.com/reference/statement/asp/api/etc#DeleteFile
    '**************************************************************

    ' 팝빌회원 사업자번호, "-"제외 10자리
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
    itemCode = "121"

    ' 문서번호
    mgtKey = "20220720-ASP-001"

    ' 파일아이디, 첨부파일목록(getFiles) API의 AttachedFile값
    FileID = "2556D18D-9380-4843-B748-5B8120C31BA5.PBF"

    On Error Resume Next

    Set result = m_StatementService.DeleteFile(CorpNum, itemCode, mgtKey, FileID, UserID)

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
                <legend>첨부파일 삭제</legend>
                    <ul>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
            </fieldset>
        </div>
    </body>
</html>
