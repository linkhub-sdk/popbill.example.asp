<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 최대 90byte의 단문(SMS) 메시지 1건 전송을 팝빌에 접수하며, 모든 수신자에게 동일 내용을 전송합니다. (최대 1,000건)
    '  - 메시지 내용이 90Byte 초과시 메시지 내용은 자동으로 제거됩니다.
    '  - https://developers.popbill.com/reference/sms/asp/api/send#SendSMS
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 광고성 메시지 여부 ( true , false 중 택 1)
    ' └ true = 광고 , false = 일반
    adsYN = False

    ' 예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
    reserveDT = ""

    ' 발신번호
    senderNum = ""

    ' 메시지 내용, 90byte초과시 길이가 조정되어 전송됨
    content = "동보메시지 내용"

    ' 수신정보배열, 최대 1000건
    Set msgList = CreateObject("Scripting.Dictionary")

    For i = 0 To 99
        Set message = New Messages

        ' 수신번호
        message.receiver = ""

        ' 수신자명
        message.receivername = " 수신자이름"+CStr(i)

        msgList.Add i, message
    Next

    ' 전송요청번호
    ' 팝빌이 접수 단위를 식별할 수 있도록 파트너가 할당한 식별번호.
    ' 1~36자리로 구성. 영문, 숫자, 하이픈(-), 언더바(_)를 조합하여 팝빌 회원별로 중복되지 않도록 할당.
    RequestNum = ""

    On Error Resume Next

    ReceiptNum = m_MessageService.SendSMS(CorpNum, senderNum, content, msgList, reserveDT, adsYN, RequestNum, UserID)

    If Err.Number <> 0 then
        code = Err.Number
        message =  Err.Description
        Err.Clears
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>단문 문자메시지 동보전송 </legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>ReceiptNum(접수번호) : <%=ReceiptNum%> </li>
                    </ul>
                <%	Else  %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <%	End If	%>
            </fieldset>
        </div>
    </body>
</html>