<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 세금계산서를 팩스로 전송하는 함수로, 팝빌 사이트 [문자·팩스] > [팩스] > [전송내역] 메뉴에서 전송결과를 확인 할 수 있습니다.
    ' - 함수 호출시 포인트가 과금됩니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#SendFAX
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외 10자리
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    testUserID = "testkorea"
    
    ' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType= "SELL"

    ' 문서번호 
    MgtKey = "20220720-ASP-001"

    ' 발신번호
    Sender = ""

    ' 수신팩스번호
    Receiver = ""

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.SendFAX(testCorpNum, KeyType, MgtKey, Sender, Receiver, testUserID)
    
    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else 
        code = Presponse.code
        message = Presponse.message
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
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message : <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>