<div align="center">

## ASP Lesson \#2 \- Get Started with Personal Webserver


</div>

### Description

In this Lesson I will teach you how to get a virtual Webserver for your ASP files if you don't own a real webserver. If you don't have a webserver installed, you can't view your ASP files.<p>This is no good to read for you who allready owns a webserver, or have installed PWS allready.<p>If you want to know how to install PWS proberly and how you can configure it to maximum performance, this might be something to read.<p>Please rate it to whatever you think it should have.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Larsson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-larsson.md)
**Level**          |Beginner
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Documents/ Frames](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/documents-frames__4-27.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-larsson-asp-lesson-2-get-started-with-personal-webserver__4-6415/archive/master.zip)





### Source Code

<b>ASP Lesson #2 - Get Started with Personal Webserver</b><p>
Now you who have read my lesson #1 will wonder why my lesson #2 will be about the software Microsoft Personal Webserver
and not about ASP? It depends on that most people dont have access to a working Webserver. Now I will explain how to setup
PWS and make the common settings.<p><p>
Because that ASP can't be run local in a simple webbrowser, the ASP files must be put within a Web Server....and I suppose
most of you reading this don't want to stay online doing the examples and other things. I'm thinking about you who got a
modem connection. I will now show you how you can run ASP-Pages local on your own HD.<br><hr><br><b>What is PWS?</b><p>
PWS stands for Personal Webserver and it's a light edition of Microsoft's standard software for common webservers(Microsoft Internet Information Server).
PWS is allthough a software that is free and that was made for one main reason. That common people can simultate a part
of their HD to be a webserver.<p>
Perhaps you're asking yourself why this is needed? Since ASP is pages that must be generated/executed by a webserver before
they can be read by a webbrowser. To do this, we use PWS to simulate the webserver and generate the HTML code out of the
ASP files.<br><hr><br><b>Where can I get PWS?</b><p>
PWS is freeware. There is mainly three ways to obtain PWS:<p><table border="0">
<tr><td>*</td><td><b>Windows 98-CDROM</b>. If you use Windows 98, the software PWS is on the Windows 98 CD-Rom.</td>
<tr><td>*</td><td><b>Internet</b>. If you use Windows 95 you have to download PWS from <a href="http://www.microsoft.com/">http://www.microsoft.com</a>.</td>
<tr><td>*</td><td><b>Other</b>. Since PWS is freeware and is a quite common and popular software it usually comes with
many computer magazines that you can buy in any store.</td></table>
<br><hr><br><b>Install PWS</b><p>When you have PWS installation files all you have to do is to install the software.
The hard part of this is to configure the software correct and to use PWS.<p>
After you have run Setup.exe make sure that you choose the Custom Installation so that you may choose install path and
other options. You'll only need a few of them. Make sure that these components are beeing installed before you continue:<p><table border="0">
<tr><td>*</td><td>Common Programfiles</td><tr><td>*</td><td>Microsoft Data Access Components 1.5</td><tr><td>*</td>
<td>Personal Webserver(PWS)</td><tr><td>*</td><td>Transaction Server</td></table><p>These for dots are the only ones you
need to use ASP. The other settings is up to you to choose if you need. When this is ready, just install the software.<br><hr><br><b>Cofigure PWS</b><p>
When the installation is finished there will be a folder in the root of c:\ which is called Inetpub ( if you didn't changed this in the installation ).
In this folder there will be a lot of other folders. The main folder we are going to use is the one called wwwroot.<p>c:\Inetpuv\wwwroot\<p>
In this folder you can create all the folders that will contain all your projects. It's beacause in wwwroot the ASP files
will be virtually simulated on your virtually webserver. Theres is allthough a little problem. You have to star the webserver
before all this is going to work for you. I will teach you that now.<p>
Lets start PWS. You should have a shortcut to it in your start menu. It can also be in the lower right corner in Windows.
You should see a software window split into two frames. Take a little look in the right one, there should be a quite big
button which is labeled start. Click it!<p>
After a few seconds the button will change it self into a stopp button, this means that your webserver now have been started.
This is a good sign, if this happens, your server should be working.<p>
Let us now configure your read and write properties for your server. In the left frame there should be a couple of icons.
Lower left you will see an icon named Advanced. Click it! You will now see a list with information of different folders. Here is where you choose which folders
have write persmission or not. You can only change properties for folders under the folder wwwroot.<p>
Start, without exiting PWS, the Explorer in Windows and open the folder wwwroot. Now create a folder named ASP ( or whatever you like to ). Now return to PWS.<p>
Click the button ADD. Choose browse and choose the folder you just created. Tap the folder you created and click OK.
In the next field it says: New Virtual Directory, change it to ASP ( or whatever your new directory was named ). In this field you will find three options that you can choose from.<p>It is, Read, Execute and Scripts. Make sure all theese options are marked and then click the OK button. You have now
made the folder: c:\inetpub\wwwroot\asp\ is a virtual folder that have Read, Execute and Scripts rights. If you don't do
this you won't be able to run any ASP code from this directory.<p>
If you would like to add more folders under the wwwroot directory you simply just create them in the Explorer, and then
give them Read, Execute and Scripts rights in PWS.<br><hr><p><b>Use PWS</b><p>
Ok, now create a simple file in you ASP directory named: Default.asp. Open it in notepad and add this code: <p>
<html><p>This file is executed on my personal webserver.<p></html><p>
Save the file and then start your webbrowser. To open the file you can't just type the url ( c:\inetpub\wwwroot\asp\default.asp ).
If you do so the file will be executed, but all the asp code will not be treated. To access this and all other files under
the wwwroot directory do like this:<p>
http://localhost/<p>And then the name on the folder you want to open. Or the file you wich to view, in this case ASP.<p>
http://localhost/asp<p>You <b>don't</b> have to write default.asp. Why not? Since this is a filename that the webserver
will be searching for, it's called the index file. The index file for HTML is Index.htm/l.<p>Now I have teached you how
to setup a webserver for you asp files. In the next lesson I will teach you what this guide actually is about: ASP!<p><br><b>End of ASP Lesson #2</b><p><hr>

