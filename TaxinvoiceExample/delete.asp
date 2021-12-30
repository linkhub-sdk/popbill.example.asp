<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 1건의 전자세금계산서를 삭제합니다.
    ' - 세금계산서를 삭제해야만 문서번호(mgtKey)를 재사용할 수 있습니다.
    ' - 삭제가능한 문서 상태 : [임시저장], [발행취소], [발행예정 취소],
    '   [발행예정 거부]
    ' - https://docs.popbill.com/taxinvoice/asp/api#Delete
    '**************************************************************

    ' 팝빌회원 사업자번호 ("-"제외)
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    testUserID = "testkorea"
    
    ' 세금계산서 발행유형, SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType= "SELL"

    ' 세금계산서 문서번호 
    MgtKey = "20211201-001"

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.Delete(testCorpNum, KeyType, MgtKey, testUserID)
    
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
                <legend>세금계산서 삭제</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>