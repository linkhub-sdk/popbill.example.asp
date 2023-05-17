<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************'
    ' 전자명세서 관련 메일전송 항목에 대한 전송여부를 수정한다.
    ' - https://developers.popbill.com/reference/statement/asp/api/etc#UpdateEmailConfig
    '
    ' 메일전송유형
    ' SMT_ISSUE : 공급받는자에게 전자명세서가 발행 되었음을 알려주는 메일입니다.
    ' SMT_ACCEPT : 공급자에게 전자명세서가 승인 되었음을 알려주는 메일입니다.
    ' SMT_DENY : 공급자에게 전자명세서가 거부 되었음을 알려주는 메일입니다.
    ' SMT_CANCEL : 공급받는자에게 전자명세서가 취소 되었음을 알려주는 메일입니다.
    ' SMT_CANCEL_ISSUE : 공급받는자에게 전자명세서가 발행취소 되었음을 알려주는 메일입니다.
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    userID = "testkorea"

    ' 메일 전송 유형
    emailType = "SMT_ISSUE"

    ' 전송 여부 (true = 전송, false = 미전송)
    sendYN = true

    On Error Resume Next

    Set Presponse = m_StatementService.updateEmailConfig(testCorpNum, emailType, sendYN, UserID)

    If Err.Number <> 0 then
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
                <legend>알림메일 전송설정 수정</legend>
                <ul>
                    <li> Response.code : <%=code%> </li>
                    <li> Response.message : <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
