<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 현금영수증과 관련된 안내 메일을 재전송 합니다.
    ' - https://docs.popbill.com/cashbill/asp/api#SendEmail
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"	 

    ' 팝빌회원 아이디
    userID = "testkorea"		 

    ' 문서번호	
    mgtKey = "20220720-ASP-001"		 

    ' 수신 메일주소
    receiver = ""		

    On Error Resume Next
        
    Set Presponse = m_CashbillService.SendEmail(testCorpNum, mgtKey, receiver, userID)

    If Err.Number <> 0 then
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
                <legend>알림메일 재전송</legend>
                <ul>
                    <li>Response.code : <%=code%></li>
                    <li>Response.message : <%=message%></li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>