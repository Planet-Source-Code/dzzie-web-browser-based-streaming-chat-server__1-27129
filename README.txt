I know this looks like alot of files...it isnt so bad
the only ones you have to worry about are:

HTTP.bas, Form1.frm, & Globals.bas

the rest are just my library files of support functions
they are like glue kinda..just macros for dumbstuff to try
to keep the rest of it clean.

anyway..


This program is a fully functional skeleton of a web based chatserver.


What does web based mean exactly? 

well not just internet based, but it is WEB BROWSER based.
Everyone can just connect to your IP in their browser 
and this program will function as a web server and serve up
streaming chat to an unlimited number of people 
(this has been almost exclusivly tested with IE)



So what does skeleton mean?

well this represents the most basic framework for this concept
to function. Users can login and exchange chat...private posts
are not yet implemented, neither is the user config or user/
admin options.

I havent been able to find any open source programs like this
that showed the concept on how to stream real time data into 
the browser like the web based chatrooms do so i had to screw
around reading RFC's and sniffing packets from an actual one.

Anyway the reason this is skeleton is because i know no one
would care how i did it if i gave them a complete thing...and really
for my needs...i dont need much more than this...

this way you can see how it is done and learn...this framework
should be robust and clean enough that you can build in whatever
options you want without any problems.

if you want to understand how to do this for yourself then you are
set because you dont have to wade hundreds of lines of code that only
implement stupid functions or fancy user interface features.

actually the UI of the exe sucks...it isnt the point so would only
add extra crap code to the mix.

also rember there are alot of security issues to deal with if you
were to make one of these for wide scale use..browsers are easily
exploited and can potentionally comprimise the whole computer...
so make your filters wisely and make sure they cant be fooled and 
that all data is checked..also think about users ghosting others etc.




So whats the flow of events?

Its a web server that can serve up 4 different pages.

login.html- they ughh login here :) data posted to frames.html

frames.html- acts like cgi script,it validates the user login
		and then parses the frames.html page to suit.
		if they are using a name in use login fails and
		they are redirected to sorry.html

banner.html - server parsed form..adds in current chatters name
		and list of other chatters on the server..no
		attempt is made to stop ghosting etc..skeleton rember

body.html  - this is the streaming chat frame where the chat will
		pop up as people post it. This is done through the
		use of a special HTTP header declaring it as a 
		chunked data transfer...i think (but dont know)
		this use of it isnt quite how it was envisioned 
		being used and a bit of a jerry rig :) but it is how
		the big online chat sights do it so 

login.html is the default document...any other page requested
will trigger the oops sorry cant find page message...there is no
chance of file disclosure vunurabilities with this app...it simply
dosent even try to find other files asked for and has all the file
locations hard coded in to be in the app.path directory.

these web pages are also bare bones...they are meant to be clear
not to boast ever feature and all fancied out.



Final words...


I wrote all of this from scratch...had no help anywhere even the rfc's suck :(
the closest i came to finding documentation on this was in the packet captures
i got from logging into the online chatrooms :-\

if anyone has stumbled across documentation geared torwards programmers on this
technique for streaming in the chat please mail me links :)

i know the libraries are kinda bulky and there are alot...but it is to early
to pick and choose now.

if you like the libraries feel free to use them in your personal projects
..please read the comment headers for software license and please leave the
headers intact..

rember no one has to share any code...but where would we be with out it?

seeing others mistakes...and others intutions helps us all out immesurably..
I know i learned more from others code at PSC than from entire books...so please
just respect others time and current coding level regardless of where you are
(or think you are)

this form of web chat can also be accomplished using php or asp actually
it probably would have been easier to not have to create my own web server
for the dumb thing..but i wanted it to be stand alone so *shrugs*

have fun

	-dzzie