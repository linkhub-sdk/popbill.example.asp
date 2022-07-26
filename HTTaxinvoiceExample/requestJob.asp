<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 홈택스에 신고된 전자세금계산서 매입/매출 내역 수집을 팝빌에 요청합니다. (조회기간 단위 : 최대 3개월)
    ' - 주기적으로 자체 DB에 세금계산서 정보를 INSERT 하는 경우, 조회할 일자 유형(DType) 값을 "S"로 하는 것을 권장합니다.
    ' - https://docs.popbill.com/httaxinvoice/asp/api#RequestJob
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"	

    ' 팝빌회원 아이디
    testUserID = "testkorea"		

    ' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType= SELL					

    ' 검색 일자유형, W-작성일자, I-발행일자, S-전송일자
    DType = "S"						

    ' 시작일자, 표시형식(yyyyMMdd)
    SDate = "20220701"			
    
    ' 종료일자, 표시형식(yyyyMMdd)
    EDate =	"20220720"				

    
    On Error Resume Next

    jobID = m_HTTaxinvoiceService.requestJob(testCorpNum, KeyType, DType, SDate, EDate, testUserID)

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