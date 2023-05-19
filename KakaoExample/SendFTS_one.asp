<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 텍스트로 구성된 1건의 친구톡 전송을 팝빌에 접수합니다.
    ' - 친구톡의 경우 야간 전송은 제한됩니다. (20:00 ~ 익일 08:00)
    ' - 전송실패시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/send#SendFTS
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    testUserID = "testkorea"

    ' 팝빌에 등록된 플러스친구 아이디
    plusFriendID = "@팝빌"

    ' 팝빌에 사전 등록된 발신번호
    ' altSendType = 'C' / 'A' 일 경우, 대체문자를 전송할 발신번호
    ' altSendType = '' 일 경우, null 또는 공백 처리
    ' ※ 대체문자를 전송하는 경우에는 사전에 등록된 발신번호 입력 필수
    senderNum = ""

    ' 친구톡 내용, 최대 1000자
    content = "친구톡 메시지 내용입니다"

    ' 대체문자 제목
    ' - 메시지 길이(90byte)에 따라 장문(LMS)인 경우에만 적용.
    altSubject = "대체문자 제목"

    ' 대체문자 유형(altSendType)이 "A"일 경우, 대체문자로 전송할 내용 (최대 2000byte)
    ' └ 팝빌이 메시지 길이에 따라 단문(90byte 이하) 또는 장문(90byte 초과)으로 전송처리
    altContent = "대체문자 메시지 내용"

    ' 대체문자 유형 (null , "C" , "A" 중 택 1)
    ' null = 미전송, C = 알림톡과 동일 내용 전송 , A = 대체문자 내용(altContent)에 입력한 내용 전송
    altSendType = "C"

    ' 예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
    reserveDT = ""

    ' 광고성 메시지 여부 ( true , false 중 택 1)
    ' └ true = 광고 , false = 일반
    ' - 미입력 시 기본값 false 처리
    adsYN = False

    ' 수신정보 구성
    Set receiverList = CreateObject("Scripting.Dictionary")
    Set rcvInfo = New KakaoReceiver
    ' 수신자번호
    rcvInfo.rcv = "01011222"
    ' 수신자명
    rcvInfo.rcvnm = " 수신자이름"
    receiverList.Add 0, rcvInfo

    ' 친구톡 버튼정보 구성
    Set btnList = CreateObject("Scripting.Dictionary")
    Set btnInfo = New KakaoButton
    btnInfo.n = "버튼이름"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbil.com"
    btnInfo.u2 = "http://www.llinkhub.co.kr"
    btnList.Add 0, btnInfo

    Set btnInfo = New KakaoButton
    btnInfo.n = "메시지 전달"
    btnInfo.t = "MD"
    btnList.Add 1, btnInfo

    ' 전송요청번호
    ' 팝빌이 접수 단위를 식별할 수 있도록 파트너가 할당한 식별번호.
    ' 1~36자리로 구성. 영문, 숫자, 하이픈(-), 언더바(_)를 조합하여 팝빌 회원별로 중복되지 않도록 할당.
    requestNum = ""

    On Error Resume Next

    receiptNum = m_KakaoService.SendFTS(testCorpNum, plusFriendID, senderNum, content, _
        altContent, altSendType, reserveDT, adsYN, receiverList, btnList, requestNum, testUserID, altSubject)

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
                <legend>친구톡 1건 전송</legend>
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