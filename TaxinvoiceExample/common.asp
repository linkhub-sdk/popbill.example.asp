<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
	LinkID = "TESTER"
	SecretKey =  "t4B19Ph5K2aIh9oNd91Q99Vwe9jST2/2IJbWjxhCgsA="
	set m_TaxinvoiceService = new TaxinvoiceService
	m_TaxinvoiceService.Initialize LinkID, SecretKey
	m_TaxinvoiceService.IsTest = True
%>