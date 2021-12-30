<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 계좌 거래내역에 메모를 저장합니다.
    ' - https://docs.popbill.com/easyfinbank/asp/api#SaveMemo
    '**************************************************************

    ' 팝빌회원 사업자번호
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 거래내역 아이디, Search API 반환항목 중 TID
    TID = "01912181100000000120191231000001"

    ' 메모 
    Memo = "20211201-asp 테스트"

    On Error Resume Next

    Set Presponse = m_EasyFinBankService.SaveMemo(CorpNum, TID, Memo, UserID)
    
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
                <legend>거래내역 메모 저장</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>