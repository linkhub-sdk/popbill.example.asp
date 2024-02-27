<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 전자명세서에 첨부된 파일의 목록을 확인합니다.
    ' - 응답항목 중 파일아이디(AttachedFile) 항목은 파일삭제(DeleteFile API) 호출시 이용할 수 있습니다.
    ' - https://developers.popbill.com/reference/statement/asp/api/etc#GetFiles
    '**************************************************************

    ' 팝빌회원 사업자번호, "-"제외 10자리
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
    itemCode = "121"

    ' 문서번호
    mgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set result = m_StatementService.GetFiles(CorpNum, itemCode, mgtKey, UserID)

    If Err.Number <> 0 Then
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
                <legend>첨부파일 목록 확인</legend>
                <ul>
                    <% If code = 0 Then
                           For i=0 To result.length-1
                    %>
                        <fieldset class="fieldset2">
                            <legend>첨부파일 [<%=i+1%>] </legend>
                            <ul>
                                <li>serialNum(첨부파일 일련번호) : <%=result.Get(i).serialNum%></li>
                                <li>attachedFile(파일아이디-첨부파일 삭제시 사용) : <%=result.Get(i).attachedFile%></li>
                                <li>displayName(첨부파일명) : <%=result.Get(i).displayName%></li>
                                <li>regDT(첨부일시) : <%=result.Get(i).regDT%></li>
                            </ul>
                        </fieldset>
                    <%
                        Next
                        Else
                    %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>
