<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 파트너가 할당한 전송요청번호를 통해 팩스 1건을 재전송합니다.
    ' - 발신/수신 정보 미입력시 기존과 동일한 정보로 팩스가 전송되고, 접수일 기준 최대 60일이 경과되지 않는 건만 재전송이 가능합니다.
    ' - 팩스 재전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
    ' - 변환실패 사유로 전송실패한 팩스 접수건은 재전송이 불가합니다.
    ' - https://developers.popbill.com/reference/fax/asp/api/send#ResendFAXRN
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 원본 팩스 전송시 할당한 전송요청번호(RequestNum)
    orgRequestNum = "1"

    ' 발신자 번호
    sendNum = "07043042991"

    ' 발신자명
    sendName = "발신자명"

    ' 전송예약시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
    reserveDT = ""

    ' 팩스 제목
    title = "팩스 재전송"

    ' 수신정보가 기존전송정보와 동일한 경우
    ReDim receivers(-1)

    ' 수신정보가 기존전송정보 다를 경우 아래 코드 참조
    'Dim receivers(0)
    'Set receivers(0) = New FaxReceiver

    ' 수신번호
    'receivers(0).receiverNum = "07066666"

    ' 수신자명
    'receivers(0).receiverName = "수신자 명칭"


    ' 재전송 팩스의 전송요청번호
    ' 파트너가 전송 건에 대해 관리번호를 구성하여 관리하는 경우 사용.
    ' 1~36자리로 구성. 영문, 숫자, 하이픈(-), 언더바(_)를 조합하여 팝빌 회원별로 중복되지 않도록 할당.
    RequestNum = ""

    On Error Resume Next

    url = m_FaxService.ResendFAXRN(CorpNum, orgRequestNum, sendNum, senderName, receivers, reserveDT , UserID, title, RequestNum)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>팩스 재전송</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>recepitNum (접수번호) : <%=url%> </li>
                    <% Else %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>