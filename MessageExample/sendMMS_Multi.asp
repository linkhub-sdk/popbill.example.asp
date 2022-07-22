<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 최대 2,000byte의 메시지와 이미지로 구성된 포토문자(MMS) 다수건 전송을 팝빌에 접수하며, 수신자 별로 개별 내용을 전송합니다. (최대 1,000건)
    ' - 이미지 파일 포맷/규격 : 최대 300Kbyte(JPEG), 가로/세로 1,000px 이하 권장
    ' - https://docs.popbill.com/message/asp/api#SendMMS
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

    For i =0 To 99
        Set message = New Messages
        ' 발신번호
        message.sender = "07043042991"

        ' 발신자명
        message.senderName = "발신자명"

        ' 수신번호
        message.receiver = "000111222"

        ' 수신자명
        message.receivername = " 수신자이름"+CStr(i)

        ' 메시지 내용, 2000byte초과시 길이가 조정되어 전송됨.
        message.content = "MMS 메시지 내용"
        
        ' 메시지 제목
        message.subject = "MMS 메시지 제목"
    
        msgList.Add i, message
    Next
    
    ' 포토메시지 이미지파일, 300Kbyte JPEG 포맷 전송가능
    FilePaths = Array("C:\popbill.example.asp\test.jpg")

    ' 전송요청번호
    ' 팝빌이 접수 단위를 식별할 수 있도록 파트너가 할당한 식별번호.
    ' 1~36자리로 구성. 영문, 숫자, 하이픈(-), 언더바(_)를 조합하여 팝빌 회원별로 중복되지 않도록 할당.
    requestNum = ""	

    On Error Resume Next

    receiptNum = m_MessageService.SendMMS(testCorpNum, "", "", "", msgList, FilePaths, reserveDT, adsYN, requestNum, userID)

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
                <legend>MMS 문자메시지 1건 전송 </legend>
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