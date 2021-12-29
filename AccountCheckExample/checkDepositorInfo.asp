<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
        <title>예금주조회 API SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"--> 
    <%
        '**************************************************************
        ' 1건의 예금주실명을 조회합니다.
        ' - https://docs.popbill.com/accountcheck/asp/api#checkDepositorInfo
        '**************************************************************
        '팝빌회원 사업자번호
        CorpNum = "1234567890"	

        '팝빌회원 아이디
        UserID = "testkorea"
        
        '기관코드
        BankCode = "0004"

        '계좌번호
        AccountNumber = "035821302044"
        
        ' 등록번호 유형 ( P / B 중 택 1 ,  P = 개인, B = 사업자)
        identityNumType = ""

        ' 등록번호
        ' - IdentityNumType 값이 "B" 인 경우 (하이픈 '-' 제외  사업자번호(10)자리 입력 )
        ' - IdentityNumType 값이 "P" 인 경우 (생년월일(6)자리 입력 (형식 : YYMMDD))
        identityNum = ""

        On Error Resume Next
            Set result = m_AccountCheckService.checkDepositorInfo(CorpNum, BankCode, AccountNumber, identityNumType, identityNum, UserID)
            
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
                <legend>계좌성명조회</legend>
            <%
                If Not IsEmpty(result) Then  
            %>
                <ul>
                    <li>bankCode (기관코드) : <%= result.bankCode%></li>	
                    <li>accountNumber (계좌번호) : <%= result.accountNumber%></li>	
                    <li>accountName (예금주 성명) : <%= result.accountName%></li>	
                    <li>checkDate (확인일시) : <%= result.checkDate%></li>	
                    <li>identityNumType (등록번호 유형) : <%= result.identityNumType%></li>	
                    <li>identityNum (등록번호) : <%= result.identityNum%></li>	
                    <li>result (응답코드) : <%= result.result%></li>	
                    <li>resultMessage (응답메시지) : <%= result.resultMessage%></li>	
                </ul>
            <%
                End If 
                If Not IsEmpty(code) then
            %>
                <ul>	
                    <li>Response.code : <%= code %> </li>
                    <li>Response.message : <%= message %></li>
                </ul>
            <%
                End If
            %>		

            </fieldset>
    </body>
</html>