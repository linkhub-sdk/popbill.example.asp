<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 팝빌 연동회원 가입을 요청합니다.
    ' - https://docs.popbill.com/statement/asp/api#JoinMember
    '**************************************************************

    ' 회원정보 객체 생성
    Set joinInfo = New JoinForm

    '링크아이디 
    joinInfo.LinkID = "TESTER"		   

    '사업자번호, "-"제외 10자리
    joinInfo.CorpNum = "1234567890"    

    '대표자성명
    joinInfo.CEOName = "대표자성명"	
    
    '상호명
    joinInfo.CorpName =  "상호"	
    
    '주소
    joinInfo.Addr =   "주소"		   

    '업태
    joinInfo.BizType =  "업태"		   

    '종목
    joinInfo.BizClass = "종목"

    '아이디 (6자 이상 20자 미만)
    joinInfo.ID =  "userid"

    '비밀번호 (8자 이상 20자 이하) 영문, 숫자 ,특수문자 조합
    joinInfo.Password =  "asdf1234!@#$"

    '담당자명
    joinInfo.ContactName = "담당자명"    

    '담당자연락처
    joinInfo.ContactTEL = "02-999-9999"   

    '담당자 휴대폰번호
    joinInfo.ContactHP = "010-1234-5678"	

    '팩스번호
    joinInfo.ContactFAX = "02-999-9999"		

    '담당자 이메일
    joinInfo.ContactEmail = "test@test.com"

    On Error Resume Next

    Set Presponse = m_StatementService.JoinMember(joinInfo)
    
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
                <legend>연동회원 가입</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>