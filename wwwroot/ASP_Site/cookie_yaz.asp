<%
response.Cookies("Ziyaretçi")("Adi")="Bedri"
response.Cookies("Ziyaretçi")("Soyadi")="Akay"
response.Cookies("Ziyaretçi")("Email")="Bedri@yasalegitim.com"
response.Cookies("Ziyaretçi")("yasi")=25
response.Cookies("Font")="Arial"
response.Cookies("Þablon")="Þablon 1"
response.Cookies("Adý")="Bedrettin" 


response.Cookies("Ziyaretçi").Expires = "30/01/2006"
response.Cookies("Font").Expires = "30/01/2006 10:47:00"
response.Cookies("Þablon").expires="30/01/2006"
response.Cookies("Adý").Expires = "30/01/2006"
response.Cookies("Adý").domain = "yasalegitim.com"
response.Cookies("Adý").path = "/data/"
response.Cookies("Adý").secure = true
%>