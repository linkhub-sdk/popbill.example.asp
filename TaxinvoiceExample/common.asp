<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
	'연동상담시 발급받은 연동아이디 
	LinkID = "TESTER"
	'연동상담시 발급받은 비밀키, 유출에 주의
	SecretKey ="ut8QMlWBgUYLCgvHqit0rmPdyBPNeWUziQLT0osDvXQ="

	set m_TaxinvoiceService = new TaxinvoiceService
	m_TaxinvoiceService.Initialize LinkID, SecretKey

	'연동환경설정값, 테스트완료후 상업용 전환시 False로 값을 수정하거나 주석처리.
	m_TaxinvoiceService.IsTest = True
%>