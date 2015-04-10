Attribute VB_Name = "MReadMe"
'Module MReadMe
'
'Readme.txt
'©2015 Kevin Lincecum AKA FrodoBaggins   email: baggins DOT frodo AT_SYMBOL gmail DOT com
'License: Free usage as long as you send me an email and mention me somewhere in your readme, about, etc
'
'
'
'Hi , I 'm Kevin Lincecum, aka frodobaggins. A long time ago I developed a car pc application for
'playing media and other things. "FrodoPlayer" if you want to search for it. Like a lot of beginning
'programmers, I chose VB to do my programming in, mainly because of the rapid application development
'aspect of it. I never really encountered any difficulties with the project, except when it came to
'doing advanced stuff with windows media player. I used windows media player as the "engine" of my
'media playing project, and for the most part, it was great. However, my users and I eventually wanted
'to change the visualizations, and use the equalizer. Well Microsoft says you can't from VB..or
'really from C++ either.
'
'But it turns out you can, IF you use the remoting services to host the control in local mode, then
'you can skin the player, and control the objects through the skin. BUT, Microsoft says you can't remote
'the control in VB, only in C++. Well I didn't accept this then, or now. I searched off and on for a long
'time trying to figure it out, even after I let that part of my life go, that application long behind me.
'It still bothered me even years later, and after some more searching, it seems that no one else figured
'it out, or did, but didn't share it ! Grrrr!
'
'Recently, I was looking over some things where I had been playing with doing this in .NET. I had it working
'pretty good there, and it got me interested in doing it in VB6. It turns out it's not that difficult to do
'and most of the information on how to do it was in the documentation the whole time, I just wasn't looking
'well enough.
'
'Anyway , here 's how it was done.
'
'When I was looking before, an aquaintence I knew from the MP3Car forums, Chuck Holbrook, aka godofcpu, posted
'in a microsoft mailing list some hints on how to control the visualization from C++ once the player was remoted.
'It didn 't seem to difficult to implement if I could get the player remoted, but getting the player remoted was
'the real problem.
'
'Microsoft says you can't remote the player in anything but C++. We all know that's bull, but figuring it out is
'a bear. (Even though the information was actually in the docs![not for vb]) Well screwing around a few years ago
'I wanted to do it in vb.net, so I began the search anew. I ran across Eric Gunnersons page which led me to a post
'by Jonathan Dibble on how to remote the player in C#. It was pretty trivial to convert this code to VB.Net, and
'soon I had a remoted player.
'
'A short time after, I had complete control of the visualizations and EQ (thanks to the hints before from godofcpu)
'in VB.NET. I was overjoyed, and used it a bit in some personal projects. I wondered then if I could back port it
'to VB6, but never got very far because life got in the way. It happens!
'
'Fast forward to a few days ago, and I decided, better late than never. I looked at the code again, and the docs
'again. The lights went off in my head. The docs and code samples from .NET said I needed to implement
'IWMPRemoteMediaServices, and IServiceProvider. To use IServiceProvider, I also needed to implement IOleClientSite.
'
'So I first made a new type library with for VB with the IWMPRemoteMediaServices interface, then made a class
'from the interface.
'
'I started to make another TypLib for IOleClientSite, and IServiceProvider, but then remembered Eduardo A. Morcillo
'aka Edanmo, had done some excellent work in the OLE area. Browsed google to find his website "Namespace Edanmo,"
'and sure enough, he had two excellent ole type libraries with the definitions already there!
'
'I implemented all the interfaces, tied it all together with ductape, spit, and bubblegum, and called SetClientSite on
'the WMP Control... And BAM, I got a call to my IServiceProvider interface. I wired that up to my IWMPRemoteMediaServices
'interface and that worked too. (Several crashes later).
'
'Now I made a simple skin from my old VB.Net code I knew worked, and tried wmp.uimode = "custom"..
'It didn 't work.
'For a long time.
'And longer.
'Then I realized my skin was FUBAR.
'So I fixed it, and HOLY MF CRAP IT WORKED. I EVEN PASSED IT A SCRIPTABLE OBJECT.
'GOT INFO BACK FROM IT! WOOT!!
'
'I celebrated.
'
'Then I wrote the skin up properly to pass the visualization objects and eq back to my test code, and a few lines later,
'that worked too!
'
'I celebrated some more. I realize this is an OLD issue, but I was still excited.
'
'Now a little while after that, I realized how stressed I had been back then that no one seemed to want to help with this
'issue, and decided that it was time to show the world, just in case it was still usefull.
'
'
'So, I have coded up a nice little test harness with I hope all the pieces to the puzzle for you to peruse and
'use to your hearts content. All I ask for in return is if you actually use any of this, or find this helpful,
'that you mention me somewhere in your about box, readme, etc. You probably want to mention some of the others too,
'depending on what you do with it.
'
'
'Better late than never,
'Kevin Lincecum
'AKA frodobaginns
'baggins DOT frodo AT_SYMBOL gmail DOT com
'
'
'
'
'
'P.S. This project is not meant to be a documentation of using wmp in a custom program, there are plenty of
'examples on how best to do that.
'
'Also, read the comments. It's real easy if you are not careful to cause an improper teardown (aka crash) of objects with this code.
'I may in the future, or you may (I suggest) to wrap this up in a custom control or something to remove these obstacles from
'your main app. Get it right, then just use it!
'
'
'
