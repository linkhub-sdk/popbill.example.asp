<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 팝빌 인증서버에 등록된 인증서의 유효성을 확인합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/cert#CheckCertValidation
    '**************************************************************

    ' 팝빌회원 사업자번호
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    userID = "testkorea"
    
    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.checkCertValidation(testCorpNum, userID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message =Presponse.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>공인인증서 유효성 확인</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>