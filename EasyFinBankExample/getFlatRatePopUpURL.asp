<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 정액제 신청 팝업 URL을 반환합니다.
    ' - https://docs.popbill.com/easyfinbank/asp/api#GetFlatRatePopUpURL
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		
    
    '팝빌회원 아이디
    userID = "testkorea"					
    
    On Error Resume Next
    
        url = m_EasyFinBankService.GetFlatRatePopUpURL(testCorpNum, userID)
    
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
                <legend>정액제 신청 팝업 URL</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>URL : <%=url%> </li>
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