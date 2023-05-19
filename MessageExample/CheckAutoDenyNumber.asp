<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팝빌회원에 등록된 080 수신거부 번호 정보를 확인합니다.
    ' - https://developers.popbill.com/reference/sms/asp/api/info#CheckAutoDenyNumber
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set Presponse = m_MessageService.CheckAutoDenyNumber(testCorpNum, UserID)

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
                <legend>080 수신거부 정보 확인</legend>
                <fieldset class="fieldset2">
                <ul>
                    <li>smsdenyNumber(수신거부번호) : <%=Presponse.smsdenyNumber%> </li>
                    <li>regDT(등록일시) : <%=Presponse.regDT%> </li>
                </ul>
                </fieldset>
        </div>
    </body>
</html>