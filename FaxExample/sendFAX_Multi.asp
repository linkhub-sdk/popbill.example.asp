<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%

    '**************************************************************
    ' 동일한 팩스파일을 다수의 수신자에게 전송하기 위해 팝빌에 접수합니다. (최대 전송파일 개수 : 20개) (최대 1,000건)
    ' - https://developers.popbill.com/reference/fax/asp/api/send#SendFAX
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 발신번호
    sendNum = ""

    ' 전송예약시간 yyyyMMddHHmmss, reserveDT값이 null 경우 즉시전송
    reserveDT = ""

    ' 수신정보 배열 최대 1000건
    Dim receivers(1)
    Set receivers(0) = New FaxReceiver
    ' 팩스 수신번호
    receivers(0).receiverNum = ""
    ' 팩스 수신자명
    receivers(0).receiverName = "수신자 명칭"
    ' 파트너 지정키, 수신자 구별용 메모
    receivers(0).interOPRefKey = "20220720-001"

    Set receivers(1) = New FaxReceiver
    ' 팩스 수신번호
    receivers(1).receiverNum = ""
    ' 팩스 수신자명
    receivers(1).receiverName = "수신자 명칭"
    ' 파트너 지정키, 수신자 구별용 메모
    receivers(1).interOPRefKey = "20220720-002"

    ' 팩스전송할 파일 (최대 20개)
    FilePaths = Array("C:\popbill.example.asp\대한민국헌법.doc","C:\popbill.example.asp\test.jpg")

    ' 광고팩스 전송여부 , true / false 중 택 1
    ' └ true = 광고 , false = 일반
    ' └ 미입력 시 기본값 false 처리
    adsYN = False

    ' 팩스제목
    title = "팩스 동보전송 제목"

    ' 전송요청번호
    ' 파트너가 전송 건에 대해 관리번호를 구성하여 관리하는 경우 사용.
    ' 1~36자리로 구성. 영문, 숫자, 하이픈(-), 언더바(_)를 조합하여 팝빌 회원별로 중복되지 않도록 할당.
    RequestNum = ""

    On Error Resume Next

    url = m_FaxService.SendFAX(CorpNum, sendNum, receivers, FilePaths, reserveDT, UserID, adsYN, title, RequestNum)

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
                <legend>팩스 전송</legend>
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
