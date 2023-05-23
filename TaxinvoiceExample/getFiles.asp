<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 세금계산서에 첨부된 파일목록을 확인합니다.
    ' - 응답항목 중 파일아이디(AttachedFile) 항목은 파일삭제(DeleteFile API) 호출시 이용할 수 있습니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#GetFiles
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외 10자리
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    testUserID = "testkorea"

    ' 세금계산서 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType = "SELL"

    ' 문서번호
    MgtKey = "20220720-ASP-002"

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.GetFiles(CorpNum, KeyType ,MgtKey, testUserID)

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
                <legend>세금계산서 첨부파일 목록 확인</legend>
                    <%
                        If code = 0 Then
                            For i=0 To Presponse.length -1
                    %>
                            <fieldset class="filedset2">
                            <legend> 첨부파일 : <%=i+1%> </legend>
                                <ul>
                                    <li> serialNum(순번) : <%=Presponse.Get(i).serialNum%></li>
                                    <li> AttachedFile(파일명) : <%=Presponse.Get(i).AttachedFile%></li>
                                    <li> DisplayName(파일아이디) : <%=Presponse.Get(i).DisplayName%></li>
                                    <li> regDT(등록일시) : <%=Presponse.Get(i).regDT%></li>
                                </ul>
                            </fieldset>
                    <%
                        Next
                        Else
                    %>
                            <ul>
                                <li>Response.dcode : <%=code%> </li>
                                <li>Response.message : <%=message%> </li>
                            </ul>
                    <%
                        End If
                    %>
            </fieldset>
        </div>
    </body>
</html>
