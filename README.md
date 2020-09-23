<div align="center">

## Tutorial For Winsock Control\(Updated Now also Includes FAQ\)


</div>

### Description

Updated Version Of Winsock Tutorial. For beginners who want to learn about Winsock Control and Networking. A must read for someone who want to Implement a Client-Server Interface.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-08-23 19:00:00
**By**             |[Keral\.C\.Patel\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/keral-c-patel.md)
**Level**          |Intermediate
**User Rating**    |4.3 (527 globes from 123 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Tutorial\_F1646789182003\.zip](https://github.com/Planet-Source-Code/keral-c-patel-tutorial-for-winsock-control-updated-now-also-includes-faq__1-48418/archive/master.zip)





### Source Code

<html>
<head>
<title></title>
<bgsound src="FLOURISH.mid" loop="-1">
</head>
<table border="1" width="100%" bgcolor="#66CCFF">
<tr>
<td></td>
</tr>
</table>
<p align="center"><i>Hello Everybody, This Winsock Tutorial is
for anyone who has not heard of winsock or have never programmed
with winsock control. First of all I would like to tell you that
there are two type of protocols in winsock control through which
we can have a successful connection. They are TCP and UDP But
here we will only discuss TCP. UDP is also Great But generally
TCP Protocol is Used. Now Lets Start....</i></p>
<p align="center"><i><font color="#FF0000">Designing
Part:-</font></i></p>
<p align="center"><i>First of all add winscok control to a
Standard exe project named 'Client'. Now Place that Winsock
Control on the form. It is invisible at runtime so its location
is not important. Place Two Text-Boxes named txtIP and txtSend
also place Command Buttons named cmdConnect and cmdSend on this
Form and in Last Place a List-Box control names 'lstMessages' on
the Form. Set Text-Boxes' Text property to "" and cmdConnect and
cmdSend's Caption Property to "Connect" and "OK" respectively.
Rename our Form to 'frmClient'. Set cmdSend's Default Property to
True. We will let the Default name for the Winsock Control as
this is the Winsock Tutorial.</i></p>
<p align="center"><i>Open another Standard exe project in another
window. All the Controls would be same as Client Project except
txtIP and cmdConnect they both are not needed here. Name this
Project as 'Server' and its Form as 'frmServer'.</i></p>
<p align="center"><i><font color="#FF0000">Now the Coding Part
for the Client Project. Write the Following Code into Code
Window:-</font></i></p>
<p align="left"><font color="#0080C0">Private Declare
Function</font> SendMessage <font color="#0080C0">Lib</font>
"user32" Alias "SendMessageA" (<font color="#0080C0">ByVal</font>
hwnd <font color="#0080C0">As Long</font>, <font color=
"#0080C0">ByVal</font> wMsg <font color="#0080C0">As Long</font>,
<font color="#0080C0">ByVal</font> wParam <font color=
"#0080C0">As Long</font>, lParam <font color="#0080C0">As</font>
<font color="#0080C0">Any</font>) <font color="#0080C0">As
Long</font></p>
<p align="left"><font color="#0080C0">Private Declare
Function</font> ReleaseCapture <font color="#0080C0">Lib</font>
"user32" () <font color="#0080C0">As Long</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
cmdConnect_Click()</p>
<p align="left"><font color="#0080C0">On Error Resume
Next</font></p>
<p align="left">Winsock1.Connect txtIP.Text, "1412" <font color=
"#00B900">'Just remember this Port Number Should be Same on which
our Server is Listening</font></p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
cmdSend_Click()</p>
<p align="left"><font color="#0080C0">On Error Resume
Next</font></p>
<p align="left">Winsock1.SendData "Client:- " & txtSend.Text</p>
<p align="left">lstMessages.AddItem "Client:- " &
txtSend.Text</p>
<p align="left">txtSend.Text = ""</p>
<p align="left">txtSend.SetFocus</p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
Form_MouseDown(Button <font color="#0080C0">As Integer</font>,
Shift <font color="#0080C0">As Integer</font>, X <font color=
"#0080C0">As Single</font>, Y <font color="#0080C0">As
Single</font>)</p>
<p align="left"><font color="#00B900">'For making the Form
Movable</font></p>
<p align="left">ReleaseCapture</p>
<p align="left">SendMessage Me.hwnd, &HA1, 2, 0&</p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
Label1_Click()</p>
<p align="left"><font color="#0080C0">On Error Resume
Next</font></p>
<p align="left"><font color="#00B900">'Letting server know that
client has Disconnected.</font></p>
<p align="left">Winsock1.SendData "Client is Disconnected!"</p>
<p align="left">DoEvents</p>
<p align="left">Unload Me</p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
Winsock1_DataArrival(<font color="#0080C0">ByVal</font>
bytesTotal <font color="#0080C0">As Long</font>)</p>
<p align="left"><font color="#0080C0">On Error Resume
Next</font></p>
<p align="left"><font color="#0080C0">Dim</font> str <font color=
"#0080C0">As String</font></p>
<p align="left">Winsock1.GetData str</p>
<p align="left">lstMessages.AddItem str</p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="center"><i><font color="#FF0000">And The Following Code
into The Server project. It is Much Same as The Client Part
Except that we have to Set Winsock Control to listen on specific
Port on the Form's Load Event.</font></i></p>
<p align="left"><font color="#0080C0">Private Declare
Function</font> SendMessage <font color="#0080C0">Lib</font>
"user32" Alias "SendMessageA" (<font color="#0080C0">ByVal</font>
hwnd <font color="#0080C0">As Long</font>, <font color=
"#0080C0">ByVal</font> wMsg <font color="#0080C0">As Long</font>,
<font color="#0080C0">ByVal</font> wParam <font color=
"#0080C0">As Long</font>, lParam <font color="#0080C0">As</font>
<font color="#0080C0">Any</font>) <font color="#0080C0">As
Long</font></p>
<p align="left"><font color="#0080C0">Private Declare
Function</font> ReleaseCapture <font color="#0080C0">Lib</font>
"user32" () <font color="#0080C0">As Long</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
cmdSend_Click()</p>
<p align="left"><font color="#0080C0">On Error Resume
Next</font></p>
<p align="left"><font color="#00B900">'This data will be sent to
the Client</font></p>
<p align="left">Winsock1.SendData "Server:- " & txtSend.Text</p>
<p align="left">lstMessages.AddItem "Server:- " &
txtSend.Text</p>
<p align="left">txtSend.Text = ""</p>
<p align="left">txtSend.SetFocus</p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
Form_Load()</p>
<p align="left"><font color="#0080C0">On Error Resume
Next</font></p>
<p align="left"><font color="#00B900">'If one Copy of Our
Application is already running then don't load a new
one</font></p>
<p align="left"><font color="#0080C0">If Not</font>
App.PrevInstance = <font color="#0080C0">True Then</font></p>
<p align="left">Winsock1.LocalPort = 1412 'This can be any Valid
Port Number</p>
<p align="left"><font color="#00B900">'Wait for Clients to
Connect with Your Server.</font></p>
<p align="left">Winsock1.Listen</p>
<p align="left"><font color="#0080C0">End If</font></p>
<p align="left">End Sub</p>
<font color="#0080C0">Private Sub</font> Form_MouseDown(Button
<font color="#0080C0">As Integer</font>, Shift <font color=
"#0080C0">As Integer</font>, X <font color="#0080C0">As
Single</font>, Y <font color="#0080C0">As Single</font>)
<p align="left"><font color="#00B900">'for making a form
Movable</font></p>
<p align="left">ReleaseCapture</p>
<p align="left">SendMessage Me.hwnd, &HA1, 2, 0&</p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
Label1_Click()</p>
<p align="left"><font color="#0080C0">On Error Resume
Next</font></p>
<p align="left"><font color="#00B900">'So that it will not raise
an error after sending the data to the server which is already
disconnected</font></p>
<p align="left">Winsock1.SendData "Server is Disconnected!"</p>
<p align="left"><font color="#00B900">'Here DoEvents gives time
to perform the winsock operation before unloading it from
memory</font></p>
<p align="left">DoEvents</p>
<p align="left"><font color="#00B900">'Now Unload it</font></p>
<p align="left">Unload Me</p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
Winsock1_ConnectionRequest(<font color="#0080C0">ByVal</font>
requestID <font color="#0080C0">As Long</font>)</p>
<p align="left"><font color="#0080C0">On Error Resume
Next</font></p>
<p align="left"><font color="#00B900">'First Check if the Winsock
Control is Connected or not If connected then Close it</font></p>
<p align="left"><font color="#0080C0">If</font> Winsock1.State
&lt;&gt; sckClosed <font color="#0080C0">Then</font>
Winsock1.Close</p>
<p align="left"><font color="#00B900">'Now accept the
Request</font></p>
<p align="left">Winsock1.Accept requestID</p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="left"><font color="#0080C0">Private Sub</font>
Winsock1_DataArrival(<font color="#0080C0">ByVal</font>
bytesTotal <font color="#0080C0">As Long</font>)</p>
<p align="left"><font color="#0080C0">On Error Resume
Next</font></p>
<p align="left"><font color="#0080C0">Dim</font> str <font color=
"#0080C0">As String</font></p>
<p align="left"><font color="#00B900">'Now we will store data
that has came into this string</font></p>
<p align="left">Winsock1.GetData str</p>
<p align="left"><font color="#00B900">'And Display that data in
the listbox</font></p>
<p align="left">lstMessages.AddItem str</p>
<p align="left"><font color="#0080C0">End Sub</font></p>
<p align="center"><font color="#FF0000">That's It Bye Until Next
tutorial In which we will see about the ByteArrays() and UDP
Protocol. You can Download the Demo for Both of these Project to
Study it and Please Note that if You are testing it on a
Stand-alone Computer then Let the IP Address Be "127.0.0.1".
Yeah, You can change the Port Number but you will have to change
it in Both the Projects. They Both have to be Same for Winsock to
Communicate.</font> <font color="#FF0000">This whole tutorial and
FAQ is also included in the zipfile. The samples included have
some extra code added to it. I will keep updating the FAQ's for
you people. If you have learned Something from this and want to
thank-me then</font></p>
<p align="center"><font color="#FF0000"><b>Please scroll down a
little and Vote for me.</b></font></p>
<p align="center">Written By:- <font color=
"#FF0080"><u>Keral.C.Patel.</u></font></p>
<p align="center">Email:- keral82@keral.com</p>
<table border="1" width="100%" bgcolor="#66CCFF">
<tr>
<td></td>
</tr>
</table>
<div align="center">
<p><font size="6"><i><font color=
"#FF8080">FAQ</font></i></font></p>
<p align="left">Q. What is this TCP/IP I have heard a lot about
it?---<font color="#008000">(By Abhishek.Net)</font></p>
<p align="left">A. TCP/IP refers to two network protocols (or
methods of data transport) used on the Internet. They are
Transmission Control Protocol and Internet Protocol,
respectively. These network protocols belong to a larger
collection of protocols, or a protocol suite. These are
collectively referred to as the TCP/IP suite. Protocols within
the TCP/IP suite work together to provide data transport on the
Internet. In other words, these protocols provide nearly all
services available to today's Net surfer. Some of those services
include Transmission of electronic mail, File transfers, Usenet
news delivery and Access to the World Wide Web. I think that most
platforms supports TCP/IP. Some of them are DOS, UNIX, Windows,
Macintosh and OS2.</p>
<p align="left">Q. Why should I specify "127.0.0.1" as my IP for
testing this code on my PC?---<font color="#008000">(By
Vrutant7287)</font></p>
<p align="left">A. This is also a detailed subject that why
should we specify "127.0.0.1" as our IP when testing something
locally. You can specify different IP and connect to that PC if
you have proper settings. E.g.:- You have a networked environment
and say there are three PC's, PC1, PC2 and PC3. You are on PC1
and you want to get connected with PC2 or PC3 then you can
specify the IP of PC2 or PC3 you will have a successful
connection only if there is another part of you application
running over there and You have set it up to listen for
connections on specific ports on that PC. For testing or running
the application locally (On standalone PC) you have to specify
"127.0.0.1" as IP. One More trick You can even specify the name
of your computer as IP. It will work.</p>
<p align="left">Q. Why Specific Port and Please tell me more
about Ports.---<font color="#008000">(By SuperCoder77)</font></p>
<p align="left">A. Here we will discuss this point with an
example. I think it will make it easier for everybody to
understand. Say For example on our server side there is an
application with a Winsock Control. In the Form Load or any
similar event we are initializing our server-side winsock control
to Listen on specific port by its Listen Method. If we don't
specify Port number then our application will get confused and it
will get data which is not meant for it. It can cause many
errors. That's why use specific port for data transactions. Ports
are the virtual gateways for communication with other objects. I
cannot cover all the things about ports over here It is out of
the scope of this tutorial.</p>
<p align="left">Q. What is sckClosed?---<font color="#008000">(By
Jack)</font></p>
<p align="left">A. It is a predefined Constant for the state of
the winsock control. If sckClosed is True then our Winsock Socket
is closed. And I would also like to explain about requestID. The
line after checking the state of our Winsock Control. In this
line of code Whenever a Client tries to connect with the Server
on the Port on which Server is listening then Server-side
Winsock's Connection Request event fires. Here we check about the
State of our control and fix it if necessary. Then we accept the
request from the client and thus a connection is established
between the Client and Server through which data can be
transferred.</p>
<p align="left">Q. I wanted to know that will GetData Method get
whole string into the variable that has been passed to it as an
argument in the parameter?---<font color="#008000">(By Emily
Gratell)</font></p>
<p align="left">A. Yeah. When Ever Winsock Control Gets any data
its Data Arrival event will fire. This is where we put our Code.
First we declare a variable and when we pass that variable in the
parameter of the GetData method of our Winsock Control it will
get all the data that was sent from the Other-side on that
specific Port.</p>
<p align="left">Q. What are the uses of Winsock Control and If I
learn this will it benefit me?---<font color="#008000">(By Ronny
Ronson)</font></p>
<p align="left">A. It is used in Client-Server environments. It
is used in the utilities for Banks and Hospitals and bigger
Corporations where there is a centralized server and all the
other Workstations are connected to it. Now It depends on you
that what benefit it will do. If you are thinking about making
Softwares for firms and banks and places where Client-Server
Interface is needed then you will surely benefit from this. This
Tutorial doesn't explains everything in detail but then also it
will get you started. I had read somewhere that whatever happens
to the Software market a programmer who knows how to implement
Client-Server Interface will never suffer.</p>
<p align="left">Q. Can I make a torjan from this? Will it execute
whatever command I send to it?---<font color="#008000">(By
Arpan.Mehta)</font></p>
<p align="left">A. I was not going to post this online but I am
getting many emails for this. Networking is a very powerful
technology and if its knowledge goes into wrong hands then, he or
she can create a havoc by using it for illegal purposes. I
personally don't recommend it. I don't believe in destruction I
believe in creation. My advice is to be creative. Now the answer
to this question is that you can surely make a trojan from it.
But be sure that where ever your trojan goes it will need VB
runtime Files if you make it in VB. This is just one idea, you
will get many bigger ideas as you go further in this subject of
TCP and networking and unleash its power.</p>
<p align="left"><b><font color="#FF0000">Note from the
Author:-</font></b> I am very pleased that people have came out
with questions. I am getting more and more questions everyday so
I thought that It would be better if I would provide a small FAQ
on this. If your question is not listed over here and you have
something different then please Email me at
<u>keral82@keral.com</u> I will try my best to answer your
questions. Regards. <u><b><font face="Trebuchet MS" color=
"#0080C0">Keral.</font></b></u></p>
</div>
<table border="1" width="100%" bgcolor="#66CCFF">
<tr>
<td></td>
</tr>
</table>
</html>

