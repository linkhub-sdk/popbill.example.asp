<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 현금영수증과 관련된 안내 SMS(단문) 문자를 재전송하는 함수로, 팝빌 사이트 [문자·팩스] > [문자] > [전송내역] 메뉴에서 전송결과를 확인 할 수 있습니다.
    ' - 메시지는 최대 90byte까지 입력 가능하고, 초과한 내용은 자동으로 삭제되어 전송합니다. (한글 최대 45자)
    ' - 알림문자 전송시 포인트가 차감됩니다. (전송실패시 환불처리)
    ' - https://developers.popbill.com/reference/cashbill/asp/api/etc#SendSMS
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 문서번호
    mgtKey = "20220720-ASP-001"

    ' 발신번호
    sender = ""

    ' 수신번호
    receiver = ""

    ' 메시지 내용, 90byte를 초과한 내용은 길이가 조정되어 전송됩니다.
    contents = "현금영수증 알림문자 테스트입니다"

    On Error Resume Next

    Set Presponse = m_CashbillService.SendSMS(CorpNum, mgtKey, Sender, Receiver, Contents, UserID)

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
                <legend>알림문자 재전송 </legend>
                <ul>
                    <li>Response.code : <%=code%></li>
                    <li>Response.message : <%=message%></li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>