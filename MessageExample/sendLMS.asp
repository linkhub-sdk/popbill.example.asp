<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 최대 2,000byte의 장문(LMS) 메시지 1건 전송을 팝빌에 접수합니다.
    ' - https://developers.popbill.com/reference/sms/asp/api/send#SendLMS
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 광고성 메시지 여부 ( true , false 중 택 1)
    ' └ true = 광고 , false = 일반
    adsYN = False

    ' 예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
	'reserveDT = "20220728200000"

    Set msgList = CreateObject("Scripting.Dictionary")

    ' 문자전송정보
    Set message = New Messages

    ' 발신번호
    message.sender = ""

    ' 발신자명
    message.senderName = "발신자명"

    ' 수신번호
    message.receiver = ""

    ' 수신자명
    message.receivername = " 수신자이름"

    ' 메시지내용, 2000byte초과시 길이가 조정되어 전송됨
    message.content = "발신 내용. 장문"

    ' 메시지제목
    message.subject = ""

    msgList.Add 0, message

    ' 전송요청번호
    ' 팝빌이 접수 단위를 식별할 수 있도록 파트너가 할당한 식별번호.
    ' 1~36자리로 구성. 영문, 숫자, 하이픈(-), 언더바(_)를 조합하여 팝빌 회원별로 중복되지 않도록 할당.
    RequestNum = ""

    On Error Resume Next

    ReceiptNum = m_MessageService.SendLMS(CorpNum, "", "","", msgList, reserveDT, adsYN, RequestNum, UserID)

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
                <legend>장문 문자메시지 1건 전송 </legend>
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
