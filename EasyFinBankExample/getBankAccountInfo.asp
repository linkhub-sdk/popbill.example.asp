<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 팝빌에 등록된 은행계좌 정보를 확인한다.
    ' - https://docs.popbill.com/easyfinbank/asp/api#GetBankAccountInfo
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		
    
    '팝빌회원 아이디
    UserID = "testkorea"
    
    ' [필수] 기관코드
    ' 산업은행-0002 / 기업은행-0003 / 국민은행-0004 /수협은행-0007 / 농협은행-0011 / 우리은행-0020
    ' SC은행-0023 / 대구은행-0031 / 부산은행-0032 / 광주은행-0034 / 제주은행-0035 / 전북은행-0037
    ' 경남은행-0039 / 새마을금고-0045 / 신협은행-0048 / 우체국-0071 / KEB하나은행-0081 / 신한은행-0088 /씨티은행-0027
    BankCode = ""

    ' [필수] 계좌번호 하이픈('-') 제외
    AccountNumber = ""

    On Error Resume Next
        Set result = m_EasyFinBankService.GetBankAccountInfo(testCorpNum, BankCode, AccountNumber, UserID)
        If Err.Number <> 0 Then
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
                <legend>계좌정보 조회</legend>
                <%
                    If code = 0 Then
                %>
                        <ul>

                            <li>bankCode (기관코드) : <%=result.bankCode%></li>
                            <li>accountNumber (계좌번호) : <%=result.accountNumber%></li>
                            <li>accountName (계좌 별칭) : <%=result.accountName%></li>
                            <li>accountType (계좌 유형) : <%=result.accountType%></li>
                            <li>state (계좌 상태) : <%=result.state%></li>
                            <li>regDT (등록일시) : <%=result.regDT%></li>
                            <li>memo (메모) : <%=result.memo%></li>

                            <li>contractDT (정액제 서비스 시작일시) : <%=result.contractDT%></li>
                            <li>useEndDate (정액제 서비스 종료일) : <%=result.useEndDate%></li>
                            <li>baseDate (자동연장 결제일) : <%=result.baseDate%></li>
                            <li>contractState (정액제 서비스 상태) : <%=result.contractState%></li>
                            <li>closeRequestYN (정액제 서비스 해지신청 여부) : <%=result.closeRequestYN%></li>
                            <li>useRestrictYN (정액제 서비스 사용제한 여부) : <%=result.useRestrictYN%></li>
                            <li>closeOnExpired (정액제 서비스 만료 시 해지 여부) : <%=result.closeOnExpired%></li>
                            <li>unPaidYN (미수금 보유 여부) : <%=result.unPaidYN%></li>
                        </ul>
                <%
                    Else
                %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>	
                <%	
                    End If
                %>
            </fieldset>
         </div>
    </body>
</html>