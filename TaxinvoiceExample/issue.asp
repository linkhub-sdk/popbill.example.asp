<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' "임시저장" 또는 "(역)발행대기" 상태의 세금계산서를 발행(전자서명)하며, "발행완료" 상태로 처리합니다.
    ' - 세금계산서 국세청 전송정책 [https://developers.popbill.com/guide/taxinvoice/asp/introduction/policy-of-send-to-nts]
    ' - "발행완료" 된 전자세금계산서는 국세청 전송 이전에 발행취소(CancelIssue API) 함수로 국세청 신고 대상에서 제외할 수 있습니다.
    ' - 세금계산서 발행을 위해서 공급자의 인증서가 팝빌 인증서버에 사전등록 되어야 합니다.
    '   └ 위수탁발행의 경우, 수탁자의 인증서 등록이 필요합니다.
    ' - 세금계산서 발행 시 공급받는자에게 발행 메일이 발송됩니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/issue#Issue
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    testUserID = "testkorea"

    ' 세금계산서 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType= "SELL"

    ' 문서번호 
    MgtKey = "20220720-ASP-002"
    
    ' 메모
    Memo = "발행 메모"

    ' 발행 안내메일 제목, 미기재시 기본양식으로 전송
    EmailSubject = ""
    
    ' 지연발행 강제여부  (true / false 중 택 1)
    ' └ true = 가능 , false = 불가능
    ' - 발행마감일이 지난 세금계산서를 발행하는 경우, 가산세가 부과될 수 있습니다.
    ' - 가산세가 부과되더라도 발행을 해야하는 경우에는 forceIssue의 값을
    '   true로 선언하여 발행(Issue API)를 호출하시면 됩니다.
    ForceIssue = False

    On Error Resume Next
    
    Set Presponse = m_TaxinvoiceService.Issue(testCorpNum, KeyType ,MgtKey, Memo ,EmailSubject, ForceIssue, testUserID)
    
    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        ntsConfirmNum = ""
        Err.Clears
    Else 
        code = Presponse.code
        message = Presponse.message
        ntsConfirmNum = Presponse.ntsConfirmNum
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>세금계산서 발행</legend>
                <ul>
                    <li>응답코드 (Response.code) : <%=code%> </li>
                    <li>응답메시지 (Response.message) : <%=message%> </li>
                    <% If ntsConfirmNum <> "" Then %>
                    <li>국세청승인번호 (Response.ntsConfirmNum) : <%=ntsConfirmNum%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>