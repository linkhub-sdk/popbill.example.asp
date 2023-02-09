<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 메시지 크기(90byte)에 따라 단문/장문(SMS/LMS)을 자동으로 인식하여 1건의 메시지를 전송을 팝빌에 접수하며, 모든 수신자에게 동일 내용을 전송합니다. (최대 1,000건)
    ' - 단문(SMS) = 90byte 이하의 메시지, 장문(LMS) = 2000byte 이하의 메시지.
    ' - https://developers.popbill.com/reference/sms/asp/api/send#SendXMS
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

    ' 발신번호
    senderNum = ""

    ' 메시지 제목
    subject = "문자전송 제목"

    ' 메시지 내용, 90byte 기준으로 단/장문 메시지가 자동인식되어 전송
    content = "문자전송 내용, 90Byte초과시 LMS(장문) 메시지로 전송됨" 


    ' 수신정보 배열, 최대 1000건	
    Set msgList = CreateObject("Scripting.Dictionary")
    
    For i =0 To 49
        Set message = New Messages

        ' 수신번호
        message.receiver = "000111222"

        ' 수신자명
        message.receivername = " 수신자이름"+CStr(i)

        msgList.Add i, message
    Next

    For i =50 To 99
        Set message = New Messages

        ' 수신번호
        message.receiver = "000111222"

        ' 수신자명
        message.receivername = " 수신자이름"+CStr(i)

        msgList.Add i, message
    Next
    
    ' 전송요청번호
    ' 팝빌이 접수 단위를 식별할 수 있도록 파트너가 할당한 식별번호.
    ' 1~36자리로 구성. 영문, 숫자, 하이픈(-), 언더바(_)를 조합하여 팝빌 회원별로 중복되지 않도록 할당.
    requestNum = ""	

    On Error Resume Next

    receiptNum = m_MessageService.SendXMS(testCorpNum, senderNum, subject,content, msgList, reserveDT, adsYN, requestNum, userID)

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
                <legend>단/장문 자동인식 문자메시지 동보전송 </legend>
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