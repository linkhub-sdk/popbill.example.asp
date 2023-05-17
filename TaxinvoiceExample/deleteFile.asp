<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' "임시저장" 상태의 세금계산서에 첨부된 1개의 파일을 삭제합니다.
    ' - 파일을 식별하는 파일아이디는 첨부파일 목록(GetFiles API) 의 응답항목 중 파일아이디(AttachedFile) 값을 통해 확인할 수 있습니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#DeleteFile
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    testUserID = "testkorea"

    ' 세금계산서 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType = "SELL"

    ' 문서번호
    MgtKey = "20220720-ASP-002"

    ' 파일아이디, 첨부파일 목록(getFiles) AttachedFile 값 참조.
    FileID = "7CB2F557-51F6-43A8-BECA-A856BDDB2CCB.PBF"

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.DeleteFile(testCorpNum, KeyType ,MgtKey, FileID, testUserID)

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
                <legend>세금계산서 첨부파일 삭제</legend>
                    <ul>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    </ul>
            </fieldset>
         </div>
    </body>
</html>