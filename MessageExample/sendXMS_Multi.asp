<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 메시지 크기(90byte)에 따라 단문/장문(SMS/LMS)을 자동으로 인식하여 1건의 메시지를 전송을 팝빌에 접수하며, 수신자 별로 개별 내용을 전송합니다. (최대 1,000건)
    ' - 단문(SMS) = 90byte 이하의 메시지, 장문(LMS) = 2000byte 이하의 메시지.
    ' - https://docs.popbill.com/message/asp/api#SendXMS
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    userID = "testkorea"

    ' 광고성 메시지 여부 ( true , false 중 택 1)
    ' └ true = 광고 , false = 일반
    adsYN = False

    ' 예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
    reserveDT = ""


    ' 문자전송정보 배열, 최대 1000건
    Set msgList = CreateObject("Scripting.Dictionary")
    
    For i = 0 To 9
        Set message = New Messages

        ' 발신번호
        message.sender = ""

        ' 발신자명
        message.senderName = "발신자명"

        ' 수신번호
        message.receiver = ""

        ' 수신자명
        message.receivername = " 수신자이름"+CStr(i)

        ' 메시지내용, 90byte기준으로 단/장문 자동인식 전송
        message.content = "문자내용이 90byte 이하인경우 단문(sms)로 전송됩니다."

        msgList.Add i, message
    Next

    For i = 50 To 99
        Set message = New Messages

        ' 발신번호
        message.sender = ""

        ' 발신자명
        message.senderName = "발신자명"

        ' 수신번호
        message.receiver = ""

        ' 수신자명
        message.receivername = " 수신자이름"+CStr(i)

        ' 메시지내용, 90byte기준으로 단/장문 자동인식 전송
        message.content = "단/장문 자동인식 메시지 테스트입니다. 문자내용의 길이가 90byte 이상인경우 장문(LMS)로 전송됩니다 단/장문 자동인식 메시지 테스트입니다."

        ' 메시지제목
        message.subject = "장문 제목입니다"

        ' 파트너 지정키, 수신자 구별용 메모
        message.interOPRefKey = "20220725-00"+CStr(i)

        msgList.Add i, message
    Next

    ' 전송요청번호
    ' 팝빌이 접수 단위를 식별할 수 있도록 파트너가 할당한 식별번호.
    ' 1~36자리로 구성. 영문, 숫자, 하이픈(-), 언더바(_)를 조합하여 팝빌 회원별로 중복되지 않도록 할당.
    requestNum = ""	

    On Error Resume Next


    receiptNum = m_MessageService.SendXMS(testCorpNum, "", "", "", msgList, reserveDT, adsYN, requestNum, userID)

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
                <legend>단/장문 자동인식 문자메시지 100건 전송 </legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>ReceiptNum(접수번호) : <%=receiptNum%> </li>
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