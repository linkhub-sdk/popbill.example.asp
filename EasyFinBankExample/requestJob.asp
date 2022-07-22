<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    '  계좌 거래내역을 확인하기 위해 팝빌에 수집요청을 합니다. (조회기간 단위 : 최대 1개월)
    ' - 조회일로부터 최대 3개월 이전 내역까지 조회할 수 있습니다.
    ' - 반환 받은 작업아이디는 함수 호출 시점부터 1시간 동안 유효합니다.
    ' - https://docs.popbill.com/easyfinbank/asp/api#RequestJob
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"	

    ' 팝빌회원 아이디
    testUserID = "testkorea"		
    
    ' 기관코드
    BankCode = ""

    ' 계좌번호
    AccountNumber = ""		

    ' 시작일자, 표시형식(yyyyMMdd)
    SDate = "20211201"			
    
    ' 종료일자, 표시형식(yyyyMMdd)
    EDate =	"20211231"				

    
    On Error Resume Next

    jobID = m_EasyFinBankService.requestJob(testCorpNum, BankCode, AccountNumber, SDate, EDate, testUserID)

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
                <legend>수집 요청</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>jobID(작업아이디) : <%=jobID%> </li>
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