<% =Now & "<br>"%> 

<%="YYYY=" & DatePart("YYYY", Date)  & "<br>"%> 
<%="Q="    & DatePart("Q", Date)     & "<br>"%> 
<%="M="    & DatePart("M", Date)     & "<br>"%> 
<%="Y="    & DatePart("Y", Date)     & "<br>"%> 
<%="D="    & DatePart("D", Date)     & "<br>"%> 
<%="W="    & DatePart("W", Date)     & "<br>"%> 
<%="WW="   & DatePart("WW", now)    & "<br>"%> 
<%="H="    & DatePart("H", now)     & "<br>"%> 
<%="N="    & DatePart("N", now)     & "<br>"%> 
<%="S="    & DatePart("S", now)     & "<br>"%> 

<%="�stenen Tarih=" & DateAdd("ww", 4, now)     & "<br>"%> 
<%dogumtarihi="01/04/1985"%>
<%="G�n=" & DateDiff("YYYY", date, dogumtarihi)     & "<br>"%>