<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팝빌 인증서버에 등록된 공동인증서의 정보를 확인합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/cert#GetTaxCertInfo
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외 10자리
    testCorpNum = "1234567890"

    On Error Resume Next

    Set resultObj = m_TaxinvoiceService.GetTaxCertInfo(testCorpNum)

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
                <legend>인증서 정보 확인</legend>
                <%

                    If code = 0 Then
                %>
                <ul>
                    <li>regDT (등록일시) : <%=resultObj.regDT %></li>
                    <li>expireDT (만료일시) : <%=resultObj.expireDT %></li>
                    <li>issuerDN (인증서 발급자 DN) : <%=resultObj.issuerDN %></li>
                    <li>subjectDN (등록된 인증서 DN) : <%=resultObj.subjectDN %></li>
                    <li>issuerName (인증서 종류) : <%=resultObj.issuerName %></li>
                    <li>oid (OID) : <%=resultObj.oid %></li>
                    <li>regContactName (등록 담당자 성명) : <%=resultObj.regContactName %></li>
                    <li>regContactID (등록 담당자 아이디) : <%=resultObj.regContactID %></li>
                </ul>

                <%
                    Else
                %>
                    <ul>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    </ul>
                <%
                    End If
                %>
            </fieldset>
        </div>
    </body>
</html>