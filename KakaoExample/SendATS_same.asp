<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 승인된 템플릿 내용을 작성하여 다수건의 알림톡 전송을 팝빌에 접수하며, 모든 수신자에게 동일 내용을 전송합니다. (최대 1,000건)
    ' - 사전에 승인된 템플릿의 내용과 알림톡 전송내용(content)이 다를 경우 전송실패 처리됩니다.
    ' - 전송실패시 사전에 지정한 변수 'altSendType' 값으로 대체문자를 전송할 수 있고, 이 경우 문자(SMS/LMS) 요금이 과금됩니다.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/send#SendATS
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    testUserID = "testkorea"

    ' 승인된 알림톡 템플릿코드
    ' └ 알림톡 템플릿 관리 팝업 URL(GetATSTemplateMgtURL API) 함수, 알림톡 템플릿 목록 확인(ListATStemplate API) 함수를 호출하거나
    '   팝빌사이트에서 승인된 알림톡 템플릿 코드를  확인 가능.
    templateCode = "019020000163"

    ' 팝빌에 사전 등록된 발신번호
    ' altSendType = 'C' / 'A' 일 경우, 대체문자를 전송할 발신번호
    ' altSendType = '' 일 경우, null 또는 공백 처리
    ' ※ 대체문자를 전송하는 경우에는 사전에 등록된 발신번호 입력 필수
    senderNum = ""

    ' 알림톡 내용, 최대 1000자
    content = "[ 팝빌 ]" & vbCrLf
    content = content + "신청하신 #{템플릿코드}에 대한 심사가 완료되어 승인 처리되었습니다." & vbCrLf
    content = content + "해당 템플릿으로 전송 가능합니다." & vbCrLf & vbCrLf
    content = content + "문의사항 있으시면 파트너센터로 편하게 연락주시기 바랍니다. " & vbCrLf & vbCrLf
    content = content + "팝빌 파트너센터 : 1600-8536" & vbCrLf
    content = content + "support@linkhub.co.kr"

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

    Set receiverList = CreateObject("Scripting.Dictionary")

    ' 수신정보 배열, 최대 1000건
    For i =0 To 9
        Set rcvInfo = New KakaoReceiver

        ' 수신자번호
        rcvInfo.rcv = "01011222"+ CStr(i)

        ' 수신자명
        rcvInfo.rcvnm = " 수신자이름"

        ' 파트너 지정키, 수신자 구별용 메모, 미사용시 공백처리
        rcvInfo.interOPRefKey = "20220720-" +CStr(i)

        receiverList.Add i, rcvInfo
    Next



    ' 전송요청번호
    ' 팝빌이 접수 단위를 식별할 수 있도록 파트너가 할당한 식별번호.
    ' 1~36자리로 구성. 영문, 숫자, 하이픈(-), 언더바(_)를 조합하여 팝빌 회원별로 중복되지 않도록 할당.
    RequestNum = ""

    ' 알림톡 버튼정보를 템플릿 신청시 기재한 버튼정보와 동일하게 전송하는 경우 btnList를 선언만 하고 함수호출.
    Set btnList = CreateObject("Scripting.Dictionary")

    ' 알림톡 버튼 URL에 #{템플릿변수}를 기재한경우 템플릿변수 영역을 변경하여 버튼정보 구성
    'Set btnInfo = New KakaoButton
    'btnInfo.n = "템플릿 안내"
    'btnInfo.t = "WL"
    'btnInfo.u1 = "https://www.popbil.com"
    'btnInfo.u2 = "http://www.llinkhub.co.kr"
    'btnList.Add 0, btnInfo

    On Error Resume Next

    ReceiptNum = m_KakaoService.SendATS(CorpNum, templateCode, senderNum, content, altContent, altSendType, reserveDT, receiverList, RequestNum, testUserID, btnList, altSubject)

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
                <legend>알림톡 동일내용 대량전송</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>ReceiptNum(접수번호) : <%=ReceiptNum%> </li>
                    </ul>
                <% Else %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <% End If %>
            </fieldset>
        </div>
    </body>
</html>
