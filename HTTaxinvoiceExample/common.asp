<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTTaxinvoiceService.asp"-->
<%
	'��ũ���̵� 
	LinkID = "TESTER"

	'���Ű
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_HTTaxinvoiceService = new HTTaxinvoiceService
	m_HTTaxinvoiceService.Initialize LinkID, SecretKey

	'����ȯ�� ������, ���߿�(True), �����(False)
	m_HTTaxinvoiceService.IsTest = True
%>