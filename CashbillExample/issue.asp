<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' [발행완료] 상태의 현금영수증을 [발행취소] 합니다.
    ' - 발행취소는 국세청 전송전에만 가능합니다.
    ' - 발행취소된 형금영수증은 국세청에 전송되지 않습니다.
    ' - 현금영수증 국세청 전송 정책 : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=asp
    ' - https://docs.popbill.com/cashbill/asp/api#CBIssue
    '**************************************************************

    '팝빌 회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		 

    '팝빌 회원 아이디
    userID = "testkorea"			 

    '문서번호
    mgtKey = "20211224-ASPR001"			 

    '메모 
    memo = "현금영수증 발행메모"	 

    On Error Resume Next

    Set Presponse = m_CashbillService.Issue(testCorpNum, mgtKey, memo, UserID)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else 
        code = Presponse.code
        message = Presponse.message
        confirmNum = Presponse.confirmNum
        tradeDate = Presponse.tradeDate
    End If

    On Error GoTo 0 

%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>현금영수증 발행</legend>
                <ul>
                    <li> Response.code : <%=code%> </li>
                    <li> Response.message : <%=message%> </li>
                    <li> Response.confirmNum : <%=confirmNum%> </li>
                    <li> Response.tradeDate : <%=tradeDate%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>