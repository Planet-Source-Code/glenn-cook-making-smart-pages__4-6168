<div align="center">

## Making Smart Pages


</div>

### Description

Sending and retrieving data between web pages is the key to developing functional web-applications. This tutorial will show you how to utilize the session in your "weblications" (that's my new favorite geek-slang) using a mix of some basic ASP and the HTTP protocol. I will show you how to get info from a user in "page1.asp" and extract it again later in "page5.asp." I will address some of the techniques you see used at popular websites like Amazon.com, many popular search engines, and just about any site that is interactive. I am assuming that you have some experience with HTML.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Glenn Cook](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/glenn-cook.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__4-7.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/glenn-cook-making-smart-pages__4-6168/archive/master.zip)





### Source Code

<p><font face="Verdana" size="2"><strong>Sending Data</strong></font></p>
<p><font face="Verdana" size="2">First, let's try to understand how we send the
data from page to page. The data is sent using two possible methods with HTTP
(Hypertext Transfer Protocol). You can pass data with HTML tags using the FORM
&quot;POST or GET&quot; methods, or with hyperlinks.</font></p>
<blockquote>
 <ol>
  <li><font face="Verdana" size="2">If you use the FORM method you basically
   tell the browser to wrap up a bunch of input box information within the
   form tags on your web page and post everything to <a href="http://www.aspalliance.com/glenncook/cookiecode.asp"><font color="mediumblue" face>a
   script</font></a> that can accept and translate the information sent.</font>
  <li><font face="Verdana" size="2">If you are sending data using a hyperlink,
   you have to customize the hyperlink to declare and define a string
   variable. It's easy! All you do is point your hyperlink to the URL, you
   add a question mark, and you add your variable. Test this example: <a href="http://www.aspalliance.com/glenncook/vbsample.asp?HomerVariable=HelloMyNameIsPuka"><font color="mediumblue" face>http://www.aspalliance.com/glenncook/vbsample.asp?HomerVariable=HelloMyNameIsPuka</font></a></font>
  <li><font face="Verdana" size="2">For another example go to <a href="http://www.nasdaq.com"><font color="mediumblue" face>NASDAQ.com
   </font></a>and look at the HTML source for that page and find the form
   post code. See how it names the input fields? Now, go back to the actual
   page and submit a ticker symbol like AMZN (Amazon.com- my pick of the
   month) and click &quot;Get Flash Quotes.&quot; If you look at the Url in
   your browser, you'll notice a bunch of variables that you sent to the
   page. If you only put in one ticker symbol, that will be the only one you
   see listed- the rest of the variables will be empty.<br>
   </font></li>
 </ol>
</blockquote>
<p><font face="Verdana" size="2"><strong>Requesting the Data<br>
<br>
</strong>With ASP, all you have to do to get your data is to &quot;Request&quot;
it. When you set up the asp script page to receive the data, you usually want to
use a &quot;If Then&quot; or a &quot;Select Case&quot; statement at the top of
your page. But before I get into the semantics of these statements let's just
get the data!</font></p>
<p><font face="Verdana" size="2">If a form is sending the data you use this asp
code to grab each input field's data:</font>
<ul>
 <li><font color="#008000" face="Courier New" size="2">&lt;% Request.Form(&quot;NameOfField1&quot;)%&gt;</font></li>
</ul>
<p><font face="Verdana" size="2">Or you can print the field's data to a page
using this code anywhere in your asp script page:</font>
<ul>
 <li><font face="Courier New"><font size="2">&lt;</font><font color="#008000" size="2">%=Request.Form(&quot;NameOfField1&quot;)%&gt;</font></font><font face="Verdana" size="2">
  (Notice that little &quot;=&quot; sign? Yep, that's all you need to know to
  print asp data into your HTML code.)</font></li>
</ul>
<p><font size="2" face="Verdana">If the data is being sent by a hyperlink use
this code to grab the variable:</font>
<ul>
 <li><font color="#008000" face="Courier New" size="2">&lt;%
  Request.Querystring(&quot;NameOfVariable&quot;) %&gt;</font></li>
</ul>
<p><font size="2" face="Verdana">Likewise, you can print the data anywhere in
the page using:</font>
<ul>
 <li><font color="#008000" face="Courier New" size="2">&lt;%=Request.QueryString(&quot;NameOfVariable&quot;)%&gt;</font></li>
</ul>
<p><font face="Verdana" size="2"><strong>Ok, so you know how to get the data,
now what do you do with it?</strong></font></p>
<p><font face="Verdana" size="2">Well, the sky is the limit! Manipulating this
data is how you make smart webpages. You can use the data to write cookies,
connect to databases and extract certain recordsets, or you can control the HTML
code that prints to the user's screen, etc. In this tutorial though I will focus
on extracting the user's information early in the visit to the site, and how you
can extract that data anytime during the session-not unlike a shopping cart
tracks the contents of your order.</font></p>
<p><font face="Verdana" size="2">ASP lets you create some global variables at
the beginning of the user session in the Global.asa file. When you hit an ASP
site the first thing ASP does is look to the Global.asa file to get any
important information it needs to know for that user session. A session begins
when a user first visits the site after the OnStart event fires in the
Global.asa file. The Global.asa file is like the config.sys file in DOS when you
first turn on your computer. It configures the user session, and lets you free
up memory space for any variables you might want. Each user session lasts as
long as the person is visiting the site. The session ends when the user closes
the browser, after an amount of time set by the system administrator, or by code
telling ASP to end the session. What you want to do with the Global.asa is free
up some memory space by declaring a few (or many) variables. As long as the
session exists, the memory space for these variables will also exist. Logically,
anytime you want to know what information is contained within these session
variables you can extract it. Just as easily you can write information to these
variables anytime. (See where I'm going with this?)</font></p>
<p><font face="Verdana" size="2">In your Global.asa file you&amp;rsquo;ll find
the Session_OnStart event and this is where you want to create your variables.</font></p>
<div align="center">
 <center>
 <table border="1" width="100%">
  <tbody>
   <tr>
    <td vAlign="top" width="50%"><font color="#008000" size="1" face="Courier New">&lt;SCRIPT
     LANGUAGE=VBScript RUNAT=Server&gt;<br>
     Sub Application_OnStart<br>
     End Sub<br>
     Sub Application_OnEnd<br>
     End Sub</font></td>
    <td vAlign="top" width="50%"><font color="#ff0000" face="Arial" size="2">Simplified:
     This code executes the very first time a user hits the weblication and
     ends some time after the last person leaves.</font></td>
   </tr>
   <tr>
    <td vAlign="top"><font color="#008000" size="1" face="Courier New">Sub
     Session_OnStart<br>
     Session(&quot;Variable1&quot;) = 0<br>
     Session(&quot;Variable2&quot;) = 0<br>
     Session(&quot;Variable3&quot;) = 0<br>
     End Sub</font></td>
    <td vAlign="top"><font color="#ff0000" face="Arial" size="2">This code
     is what you're concerned with. When a user hits your site, ASP goes
     right to the Global.asa and fires this event to start a session. We've
     also declared a few variables that we will use while the user is at
     the site. These variables end when a user ends the session.</font></td>
   </tr>
   <tr>
    <td vAlign="top"><font color="#008000" face="Courier New" size="1">Sub
     Session_OnEnd<br>
     End Sub<br>
     &lt;/SCRIPT&gt;</font></td>
    <td vAlign="top"><font color="#ff0000" face="Arial" size="2">This code
     fires at the end of a user session.</font></td>
   </tr>
  </tbody>
 </table>
 </center>
</div>
<p><font face="Verdana" size="2">This is really all there is to it. Now you can
use the information you extract from the user and write it to one of these
variables. I would use this code at the beginning of the asp script that
receives the information sent by a form. All I&amp;rsquo;m doing is writing the
submitted data to the session variables, which I can extract anytime during the
session.</font></p>
<p><font color="#008000" face="Courier New" size="2">&lt;%<br>
dim variable1<br>
dim variable2<br>
dim variable3</font></p>
<p><font color="#008000" face="Courier New" size="2">variable1 = Request.Form(&quot;Field1&quot;)<br>
variable2 = Request.Form(&quot;Field2&quot;)<br>
variable3 = Request.Form(&quot;Field3&quot;)</font></p>
<p><font color="#008000" face="Courier New" size="2">Session(&quot;Variable1&quot;)
= variable1<br>
Session(&quot;Variable2&quot;) = variable2<br>
Session(&quot;Variable3&quot;) = variable3<br>
%&gt;</font></p>
<p><font face="Verdana"><font size="2">Now you can print that variable to a page
anytime by using the code, </font><font color="#008000" size="2">&lt;%=Session(&quot;Variable1&quot;)%&gt;</font><font size="2">.
You can also extract that information to use in an SQL statement:</font></font><font face="Arial" size="2"><br>
</font><font face="Courier New"><font color="#008000" size="2">&lt;%<br>
dim MartinPrince<br>
MartinPrince = Session(&quot;Variable1&quot;)<img align="right" src="http://www.aspalliance.com/glenncook/images/bumhacker.jpg" width="114" height="128"><br>
%&gt;<br>
</font><font color="#007f00" size="2">&lt;% &quot;SELECT * FROM Employees Where
FirstName = '&quot; &amp; MartinPrince &amp; &quot;'&quot; %&gt;</font></font></p>

