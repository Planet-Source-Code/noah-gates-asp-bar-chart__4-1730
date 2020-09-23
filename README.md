<div align="center">

## ASP bar chart


</div>

### Description

This code is an example how to use a random color generator and a table to create a Bar Chart.
 
### More Info
 
Numeric values to graph

The code is an example only. The user will need to adjust it to display the data which they wish to graph.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Noah Gates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/noah-gates.md)
**Level**          |Unknown
**User Rating**    |3.0 (15 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Graphics/ Sound](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics-sound__4-15.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/noah-gates-asp-bar-chart__4-1730/archive/master.zip)





### Source Code

```
'-------------------------------------------------------------------------------
' This is a sample of how to create your own bar chart using ASP
'-------------------------------------------------------------------------------
<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Document Title</TITLE>
</HEAD>
<basefont FACE="Arial, Helvetica, sans-serif" size=2>
<link REL="STYLESHEET" HREF="../../Stylesheets/MENUPAGES.css">
<body BACKGROUND="../../Images/Redside/Background/Back2.jpg" BGCOLOR="White">
<STYLE MEDIA="print">  /* turn off navigation and ad bar */
  #header, .menu, .header,   #adbar, DIV#jumpbar, #navbar {display: none}
  /* remove position of contents */  #why {position: static; width: auto}
  /* cleanup all the colors */  #why H3, PRE, FONT {color: black}
  /* cleanup the copyright */
  DIV#why P.copyright {color: black; font: 10pt verdana}
  /* display the print header */  #printMessage {display: block} </STYLE>
<%
Dim totals(30)
Dim Names(30)
StatIndex = -1
%>
  <FORM ACTION="Inboundstats.asp" METHOD="POST" NAME="Inboundstats">
<%
On Error Resume Next
Randomize
'-------------------------------------------------------------------------------
' Purpose: To generate random colors
'-------------------------
Function rndcolor()
DIM LightDarkRange
LightDarkRange = 150
forered = Int(Rnd * LightDarkRange)
foregreen = Int(Rnd * LightDarkRange)
foreblue = Int(Rnd * LightDarkRange)
color = "#" & cstr(Hex(RGB(forered, foregreen, foreblue)))
rndcolor=color
End Function
%>
<center>
<table border="0">
<td align="left"><img src="../../images/ac204.gif"></center></td>
<td>  </td>
<td align=center>  </td>
<td align="center"><h1><font color=blue>Inbound Calls </font><h1></td>
</table>
</center>
<center>
<table border="1">
<tr VALIGN=BOTTOM>
<%
for I = 0 to 20
	vol = int(chartinfo(I))
	label = int(chartinfo(I))
<td align="center">
<table>
<td height="<% response.write vol %>" bgcolor="<% response.write rndcolor() %>" align="bottom">
<B><font color="#FFFFFF"><% response.write label %></B>
</font></td></table></td>
next
%>
</tr>
<tr>
<%
for i = 0 to StatIndex
%>
<%if totals(i) > 0 then%>
<td>
<table>
<td><FONT face="Small Fonts" size=1 color="blue"><strong> <%response.write names(i)%> <strong><font></td>
</table>
<%end if%>
<%next%>
</td>
</tr>
</table>
</center>
</BODY>
</HTML>
ngates@partnershealth.com
```

