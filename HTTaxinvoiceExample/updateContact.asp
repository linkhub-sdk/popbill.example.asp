<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/member#UpdateContact
    '**************************************************************

    ' 팝빌회원 사업자번호
    CorpNum = "1234567890"		 

    ' 팝빌회원 아이디 
    UserID = "testkorea"				 


    ' 담당자 정보 객체 생성
    Set contInfo = New ContactInfo
    
    ' 담당자 아이디 
    contInfo.id = UserID

    ' 담당자명
    contInfo.personName = "ASPTest"

    ' 담당자 연락처
    contInfo.tel = ""

    ' 담당자 이메일주소
    contInfo.email = ""

    ' 담당자 조회권한 1 - 개인권한 / 2 - 읽기권한  / 3 - 회사권한
    contInfo.searchRole = 3

    On Error Resume Next

    Set Presponse = m_HTTaxinvoiceService.UpdateContact(CorpNum, contInfo, UserID)
    
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
                <legend>담당자 정보수정</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>