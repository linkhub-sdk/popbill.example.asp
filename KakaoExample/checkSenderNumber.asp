<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 카카오톡 발신번호 등록여부를 확인합니다.
    ' - 발신번호 상태가 '승인'인 경우에만 리턴값 'Response'의 변수 'code'가 1로 반환됩니다.
    ' - https://docs.popbill.com/kakao/asp/api#CheckSenderNumber
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"
    
    ' 팝빌회원 아이디
    userID = "testkorea"			   

    ' 확인할 발신번호
    SenderNumber = ""  

    On Error Resume Next

    Set result = m_KakaoService.CheckSenderNumber(testCorpNum, SenderNumber, userID)

    If Err.Number <> 0 then
        code = Err.Number
        message =  Err.Description
        Err.Clears
    Else
        code = result.code
        message = result.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>발신번호 등록여부 확인</legend>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>	
            </fieldset>
         </div>
    </body>
</html>