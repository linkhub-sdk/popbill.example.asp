<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTTaxinvoiceService.asp"-->
<%
	'링크아이디 
	LinkID = "TESTER"

	'비밀키
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_HTTaxinvoiceService = new HTTaxinvoiceService
	m_HTTaxinvoiceService.Initialize LinkID, SecretKey

	'연동환경 설정값, 개발용(True), 상업용(False)
	m_HTTaxinvoiceService.IsTest = True
%>