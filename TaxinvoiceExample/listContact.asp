<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/member#ListContact
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_TaxinvoiceService.ListContact(testCorpNum, UserID)

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
                <legend>담당자 목록 조회</legend>
                <%
                    If code = 0 Then
                        For i=0 To result.Count-1
                %>
                            <fieldset class="fieldset2">
                                <legend> 담당자정보 [ <%=i+1%> / <%=result.Count%> ] </legend>
                                    <ul>
                                        <li> id(아이디) : <%=result.Item(i).id%></li>
                                        <li> personName(담당자 성명) : <%=result.Item(i).personName%></li>
                                        <li> email(담당자 이메일) : <%=result.Item(i).email%></li>
                                        <li> tel(담당자 연락처) : <%=result.Item(i).tel%></li>
                                        <li> regDT(등록일시) : <%=result.Item(i).regDT%></li>
                                        <li> searchRole(담당자 조회권한) : <%=result.Item(i).searchRole%></li>
                                        <li> mgrYN(관리자 여부) : <%=result.Item(i).mgrYN%></li>
                                        <li> state(상태) : <%=result.Item(i).state%></li>
                                    </ul>
                                </fieldset>
                <%
                        Next
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