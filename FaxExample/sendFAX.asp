<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 팩스를 전송합니다.
    ' - https://docs.popbill.com/fax/asp/api#SendFAX
    '**************************************************************

    '팝빌 회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		

    '팝빌 회원 아이디
    userID = "testkorea"			

    '발신자 번호
    sendNum = "07043042992"	

    '전송예약시간 yyyyMMddHHmmss,  공백처리시 즉시전송
    reserveDT = ""	
    
    '수신자 정보 
    Dim receivers(0)
    Set receivers(0) = New FaxReceiver

    '수신번호
    receivers(0).receiverNum = "070111222"

    '수신자명
    receivers(0).receiverName = "수신자 명칭"

    '팩스전송할 파일 (최대 20개)
    FilePaths = Array("C:\popbill.example.asp\대한민국헌법.doc","C:\popbill.example.asp\test.jpg")

    '광고팩스 전송여부
    adsYN = False

    '팩스제목
    title = "ASP  팩스 전송 테스트"

    '전송요청번호 (팝빌 회원별 비중복 번호 할당)
    '영문,숫자,'-','_' 조합, 최대 36자
    requestNum = ""		

    On Error Resume Next

    url = m_FaxService.SendFAX(testCorpNum , sendNum, receivers, FilePaths, reserveDT , userID, adsYN, title, requestNum)

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