First you must register
winSubHook.tlb (Found under directory Dependencies)
Note that you only need this for compiling, doesn´t need to be distributed
with the application.

Before you can use the components and the demo you need to compile
TCP-Server.dll
TCP-Client.dll

After that set a reference in each of the demoproject to the appropiate .dll
and compile.

The TCP Components are very influenced by some authors and due credit and
a big thanks for teaching me about Socket programming goes to following authors.
Coding Genius
Edwin Vermeer
Trevor Herselman
Emiliano Scavuzzo for influencing me to use Paul catons WinsubHook
Other authors on www.planet-Source-code.com/vb
www.allapi.net

And a big thanks to Paul Caton for providing me with the winsubhook that
are giving me so much better response than the common subclassing techniques.
Only one problem, I cant explain why my TCP-framework respons so much better
with his solution than the common solution. If anyone could try to explain this
to me I would be grateful :-)

Note that the TCP-Components are not needed to be compiled
You could integrate them directly into your project.
For an example look at 
http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=52538&lngWId=1

The demo projects are not to be rated, they are a mess and not much energy is used
on developing them. They could be buggy. They are just included to give you an idea
on how to use the TCP-components.