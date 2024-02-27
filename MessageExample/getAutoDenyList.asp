<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 전용 080 번호에 등록된 수신거부 목록을 반환합니다.
    ' - https://developers.popbill.com/reference/sms/asp/api/info#GetAutoDenyList
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    On Error Resume Next

    Set Presponse = m_MessageService.GetAutoDenyList(CorpNum)

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
                <legend>080 수신거부 목록 확인</legend>
                <%
                    For i=0 To Presponse.length -1
                %>
                <fieldset class="fieldset2">
                <ul>
                    <li>number(수신거부번호) : <%=Presponse.Get(i).number%> </li>
                    <li>regDT(등록일시) : <%=Presponse.Get(i).regDT%> </li>
                </ul>
                </fieldset>
                <%
                    Next
                %>

        </div>
    </body>
</html>
