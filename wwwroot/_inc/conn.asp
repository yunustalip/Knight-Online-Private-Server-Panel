<!--#include file="DbSettings.Asp"-->
<!--#include file="Connect.Asp"-->
<%

Dim Bag,Conne
Set Bag = New Baglanti
Set Conne = Bag.Connect(Sunucu,VeriTabani,Kullanici,Sifre)

%>