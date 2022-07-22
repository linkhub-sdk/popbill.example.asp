<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"--> 
    <%
        '**************************************************************
        ' ���ε� �˸��� ���ø� ������ Ȯ���մϴ�.
        ' - https://docs.popbill.com/kakao/asp/api#GetATSTemplate
        '**************************************************************

        ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
        testCorpNum = "1234567890"		

        ' ���ø� �ڵ�
        templateCode = "021120000347"

        ' �˺�ȸ�� ���̵�
        UserID = "testkorea"

        On Error Resume Next

        Set resultObj = m_KakaoService.GetATSTemplate(testCorpNum, templateCode, UserID)

        If Err.Number <> 0 then
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
                <legend>�˸��� ���ø� ���� Ȯ�� </legend>
                    <% 
                        If code = 0 Then
                    %>
                        <fieldset class="fieldset2">
                            <legend>  ���ø� ���� </legend>
                            <ul>
                                <li> templateCode : <%=resultObj.templateCode%></li>
                                <li> templateName : <%=resultObj.templateName%></li>
                                <li> template : <%=resultObj.template%></li>
                                <li> plusFriendID : <%=resultObj.plusFriendID%></li>
                                <li> ads : <%=resultObj.ads%></li>
                                <li> appendix : <%=resultObj.appendix%></li>
                            </ul>
                        <%
                            For i=0 To UBound(resultObj.btns) -1
                        %>
                                <fieldset class="fieldset3">
                                    <legend> ��ư���� [ <%=i+1%> / <%= UBound(resultObj.btns)%> ] </legend>
                                    <ul>
                                        <li>n : <%=resultObj.btns(i).n%> </li>
                                        <li>t : <%=resultObj.btns(i).t%> </li>
                                        <li>u1 : <%=resultObj.btns(i).u1%> </li>
                                        <li>u2 : <%=resultObj.btns(i).u2%> </li>
                                    </ul>
                            </fieldset>
                        <% 
                                Next
                        %>
                        </fieldset>
                        <%
                        Else
                    %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
            </fieldset>
         </div>
    </body>
</html>