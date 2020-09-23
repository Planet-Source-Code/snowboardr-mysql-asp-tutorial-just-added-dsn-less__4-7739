<div align="center">

## MySQL & ASP tutorial    \[Just added: dsn\-less\]


</div>

### Description

Learn how to Install MySQL, import an Access database into MySQL database, and display data on a ASP page. Step by step tutorial with screenshots! Please leave questions or comments.. Updates: I have included a dsn-less connection!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[snowboardr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/snowboardr.md)
**Level**          |Intermediate
**User Rating**    |5.0 (144 globes from 29 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/snowboardr-mysql-asp-tutorial-just-added-dsn-less__4-7739/archive/master.zip)





### Source Code

```
<br>
<br>
<div align="center"><a href="http://www.vzio.com/?ref=psc"><img src="http://www.vzio.com/images/logos/vzio_fl2.gif" border="0"></a>
</div>
<br>
<br>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24" align="center">
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FF0000">UPDATE</font></b>
  - DSN-LESS connection <a href="#dsnless">click here</a>.</font></td>
 </tr>
 <tr>
 <td height="18">
  <div align="center"></div>
 </td>
 </tr>
 <tr>
 <td height="18"><font size="3" face="Verdana, Arial, Helvetica, sans-serif">MySQL
  & ASP</font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
  MySql Database</font></b></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Using
  a MySql database in ASP is pretty simple, but it's hard to come accross
  a straight forward tutorial on ASP & MySQL. So here it is, a mysql tutorial
  for you.. Enjoy!</font></td>
 </tr>
 <tr>
 <td height="12"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>1.</b>
  Start by downloading the latest version of MySQL at <a href="http://www.mysql.com" target="_blank">http://www.mysql.com</a>.</font></td>
 </tr>
 <tr>
 <td height="18">
  <ul>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">The latest
   versions are listed to the right on the front page.</font></li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">On the
   next page you will find different OS's, download the windows version.</font></li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Then find
   the mirror closes to you and start the download!</font></li>
  </ul>
 </td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>2.</b>
  Extract and run the setup.exe file</font></td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
  <ul>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Install
   to c:\mysql</font></li>
  </ul>
  </font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>3.</b>
  Open Notpad and enter the following information:</font></td>
 </tr>
 <tr>
 <td height="2">
  <ul>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Save it
   to Windows Root Directory (C:/WINDOWS)</font></li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Save it
   as <b>my.ini</b></font></li>
  </ul>
 </td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#FFFFFF">
  <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
  [mysqld]<br>
  <br>
  basedir=c:/mysql<br>
  datadir=c:/mysql/data</font></div>
 </td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>4.</b>
  Open Windows command prompt.</font></td>
 </tr>
 <tr>
 <td height="4"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18">
  <ol>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Start
   > Programs > Accessories > Comand Prompt</font></li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Type the
   following:</font></li>
  </ol>
 </td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18">
  <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>C:\></b>cd
  C:/mysql/bin</font></div>
 </td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
  <b>C:\mysql/bin></b>mysqld-nt --install</font></td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>5.</b>
  The service is now installed. It can be started and stopped with the Windows
  Service manager, or the NET START/STOP commands.</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="18"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>6.</b>
  Now download MyODBC which is the driver to connect to mysql from within
  ASP.</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="59">
  <ol>
  <li>
   <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="http://www.mysql.com" target="_blank">http://www.mysql.com</a></font></div>
  </li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">On the
   right side click MyODBC<b><i>_VERSION#</i></b></font></li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Download
   the exe and install</font></li>
  </ol>
 </td>
 </tr>
 <tr>
 <td height="9"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>7.</b>
  A nice program to have when using MySQL is MySQL-Front which is a pretty
  easy to understand program that access your mysql database(s).</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="16">
  <ul>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="http://www.anse.de/mysqlfront/" target="_blank">http://www.anse.de/mysqlfront/</a></font></li>
  </ul>
 </td>
 </tr>
 <tr>
 <td height="16"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="16" bgcolor="#EEEEEE"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>8.</b>
  Start up MySQL front.</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="16">
  <ol>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Click
   <i>New</i></font></li>
  <li><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Name the
   connect whatever you like..</font></li>
  <li><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Host /
   IP: <b>Localhost</b> or <b>Your IP</b></font></li>
  <li><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Password
   can be left as root for now..</font></li>
  <li><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Click
   connect.</font></li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Right
   click root@localhost and click create new database.</font></li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Name the
   new database tutorialdb</font></li>
  </ol>
  <p align="center"><img src="http://www.vzio.com/tutorials/mysql/createnew.gif" width="276" height="127"></p>
 </td>
 </tr>
 <tr>
 <td height="16">
  <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></div>
 </td>
 </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24" align="center">
 <tr>
 <td height="18"> </td>
 </tr>
 <tr>
 <td height="18" bgcolor="#CCCCCC"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#003366">Page
  2 - MySQL Database</font></b></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="../images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>9.</b>
  Entering data into the database can be done several ways, manualy in MySQL-Front
  or by Importing from another database. Here we are going to import this
  database: <a class="lesson" href="http://www.vzio.com/tutorials/database.mdb" target="_blank">Download
  database</a></font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="12"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>10.</b>
  Im-/Export > ODBC Import</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18">
  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/tutorials/mysql/importmdb.gif" width="176" height="213"></font></div>
 </td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>11.</b>
  Check the MS Access File radio button then tables will load in the table
  list box below, select the tables you want to important and then click import.</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18">
  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/tutorials/mysql/mdb.gif" width="494" height="467"></font></div>
 </td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>12.</b>
  Now you can view the table and its data.</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18">
  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/tutorials/mysql/front_b.gif" width="526" height="269"></font></div>
 </td>
 </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24" align="center">
 <tr>
 <td height="18"> </td>
 </tr>
 <tr>
 <td height="18"> </td>
 </tr>
 <tr>
 <td height="18"> </td>
 </tr>
 <tr>
 <td height="18" bgcolor="#CCCCCC"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#003366">Page
  3 - MySQL Database</font></b></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>13.</b>
  Now we need to create a dsn connection this really easy and takes about
  3 minutes at most.</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="12"> </td>
 </tr>
 <tr>
 <td height="12"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#003366"><a name="dsnless"></a></font><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif">NEW!
  > DSN-LESS connnection:</font></b></td>
 </tr>
 <tr>
 <td height="12"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Ok
  here it is, people have been asking for how to connect to MySQL without
  a dsn.</font></td>
 </tr>
 <tr>
 <td height="17"> </td>
 </tr>
 <tr>
 <td height="12"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%<br>
  set conn = createobject("ADODB.Connection")<br>
  conn.open = "DRIVER={MySQL ODBC 3.51 Driver};"_<br>
  & "SERVER=localhost;"_ <br>
  & "DATABASE=useremails;"_<br>
  & "UID=root;PWD=; OPTION=35;" </font>
  <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">set rs =
  conn.Execute("SELECT * FROM UserEmails")<br>
  %></font>
 </td>
 </tr>
 <tr>
 <td height="12"> </td>
 </tr>
 <tr>
 <td height="12"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>14.</b>
  Start > Administrator Tools > Data Sources (ODBC)</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18">
  <ul>
  <li>
   <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Click
   system dns tab</font></div>
  </li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Click
   the "Add..." button.</font></li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Scroll
   to the bottom and select MySQL x.xx Driver</font></li>
  </ul>
 </td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>15.</b>
  Fill in the data like it is shown below in the screenshot.</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18">
  <ul>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">When finished
   filling it in, click "Test Data Source" to see if it works..</font></li>
  <li><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Click
   OK</font></li>
  </ul>
 </td>
 </tr>
 <tr>
 <td height="18">
  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/tutorials/mysql/dns_setup.gif" width="509" height="366"></font></div>
 </td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18" bgcolor="#EEEEEE"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>16.</b>
  Now for the ASP... (<b><i>If the ASP doesn't show up, please visit page
  3 below...</i></b>)</font></td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="http://www.vzio.com/images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="18">
  <div align="left">
  <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b> </b></font></p>
  <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b> <%
   </b></font></p>
  <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
   <font color="#0000FF">Set</font> my_conn = Server.CreateObject("ADODB.Connection")
   <br>
   <font color="#0000FF">Set</font> rs = Server.CreateObject("ADODB.Recordset")
   </font></p>
  </div>
  <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">my_conn.Open
  "DSN=<b>mysql_dsn</b>" <font color="#009900">' Data source name</font><br>
  </font></p>
  <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">strSQL
  = "SELECT * FROM UserEmails" <br>
  <br>
  <font color="#009900">' Execute SQL statement </font><br>
  Set rs = my_conn.Execute(strSQL) <br>
  <br>
  Do while not rs.eof ' Do until it reaches end of db <br>
  <br>
  <br>
  Response.Write "User ID# " & rs("userID") & "<br>
  " <br>
  Response.Write "Name: " & rs("name") & "<br>
  " <br>
  Response.Write "Email: " & rs("email") & "<br>
  " <br>
  </font></p>
  <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
  rs.MoveNext<font color="#009900"> ' Next record </font><br>
  loop <br>
  <br>
  my_conn.close <font color="#009900">' Close database connection </font><br>
  <font color="#0000FF">set</font> my_conn = nothing <font color="#009900">'obj
  variable released </font><br>
  <br>
  <b>%></b> </font></p>
  <p align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><i><br>
  </i> </b> </font></p>
 </td>
 </tr>
 <tr>
 <td height="18">
  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a class="tutorial" href="http://www.vzio.com/tutorials/mysql_database.asp?page=1" target="_blank">Page
  1</a> - <a href="http://www.vzio.com/tutorials/mysql_database.asp?page=2" target="_blank">Page
  2</a> - <a href="http://www.vzio.com/tutorials/mysql_database.asp?page=3" target="_blank">Page
  3</a></font></div>
 </td>
 </tr>
 <tr>
 <td height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
 </tr>
 <tr>
 <td height="18">
  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
  <%
 <font color="#990000">Response.End</font> %>
  </font></div>
 </td>
 </tr>
 <tr>
 <td height="18">
  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">(The
  End) </font></div>
 </td>
 </tr>
 <tr>
 <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="../images/small/spacer.gif" width="10" height="5"></font></td>
 </tr>
 <tr>
 <td height="2">
  <div align="center"><a href="http://www.vzio.com/" target="_blank">View
  other tutorials by me at my ASP web development Web site.</a></div>
 </td>
 </tr>
 <tr>
 <td height="2"> </td>
 </tr>
 <tr>
 <td height="2">
  <div align="center"><b>Leave a comment! And vote if you liked my tutorial.</b></div>
 </td>
 </tr>
</table>
<br>
<br>
<div align="center"><a href="http://www.vzio.com/?ref=psc"><img src="http://www.vzio.com/images/logos/vzio_fl2.gif" border="0"></a>
</div>
<br>
<br>
```

