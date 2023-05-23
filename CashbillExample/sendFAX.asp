<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 현금영수증을 팩스로 전송하는 함수로, 팝빌 사이트 [문자·팩스] > [팩스] > [전송내역] 메뉴에서 전송결과를 확인 할 수 있습니다.
    ' - 팩스 전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
    ' - https://developers.popbill.com/reference/cashbill/asp/api/etc#SendFAX
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    '팝빌회원 아이디
    UserID = "testkorea"

    '문서번호
    mgtKey = "20220720-ASP-001"

    '발신번호
    sender = ""

    '수신팩스번호
    receiver = ""

    On Error Resume Next

    Set Presponse = m_CashbillService.SendFAX(CorpNum, mgtKey, Sender, Receiver, UserID)

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
                <legend>현금영수증 팩스전송 </legend>
                <ul>
                    <li>Response.code : <%=code%></li>
                    <li>Response.message : <%=message%></li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>