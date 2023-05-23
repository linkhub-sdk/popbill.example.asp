<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************'
    ' 전자세금계산서 관련 메일전송 항목에 대한 전송여부를 수정한다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#UpdateEmailConfig
    '
    ' 메일전송유형
    ' [정발행]
    ' TAX_ISSUE : 공급받는자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
    ' TAX_ISSUE_INVOICER : 공급자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
    ' TAX_CHECK : 공급자에게 전자세금계산서가 수신확인 되었음을 알려주는 메일입니다.
    ' TAX_CANCEL_ISSUE : 공급받는자에게 전자세금계산서가 발행취소 되었음을 알려주는 메일입니다.
    '
    ' [역발행]
    ' TAX_REQUEST : 공급자에게 세금계산서를 전자서명 하여 발행을 요청하는 메일입니다.
    ' TAX_CANCEL_REQUEST : 공급받는자에게 세금계산서가 취소 되었음을 알려주는 메일입니다.
    ' TAX_REFUSE : 공급받는자에게 세금계산서가 거부 되었음을 알려주는 메일입니다.
    '
    ' [위수탁발행]
    ' TAX_TRUST_ISSUE : 공급받는자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
    ' TAX_TRUST_ISSUE_TRUSTEE : 수탁자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
    ' TAX_TRUST_ISSUE_INVOICER : 공급자에게 전자세금계산서가 발행 되었음을 알려주는 메일입니다.
    ' TAX_TRUST_CANCEL_ISSUE : 공급받는자에게 전자세금계산서가 발행취소 되었음을 알려주는 메일입니다.
    ' TAX_TRUST_CANCEL_ISSUE_INVOICER : 공급자에게 전자세금계산서가 발행취소 되었음을 알려주는 메일입니다.
    '
    ' [처리결과]
    ' TAX_CLOSEDOWN : 거래처의 휴폐업 여부를 확인하여 안내하는 메일입니다.
    ' TAX_NTSFAIL_INVOICER : 전자세금계산서 국세청 전송실패를 안내하는 메일입니다.
    '
    ' [정기발송]
    ' ETC_CERT_EXPIRATION : 팝빌에 등록된 인증서의 만료예정을 안내하는 메일입니다.
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    '팝빌회원 아이디
    UserID = "testkorea"

    '메일 전송 유형
    emailType = "TAX_ISSUE"

    '전송 여부 (true = 전송, false = 미전송)
    sendYN = true

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.updateEmailConfig(CorpNum, emailType, sendYN, UserID)

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
                <legend>알림메일 전송설정 수정</legend>
                <ul>
                    <li> Response.code : <%=code%> </li>
                    <li> Response.message : <%=message%> </li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>