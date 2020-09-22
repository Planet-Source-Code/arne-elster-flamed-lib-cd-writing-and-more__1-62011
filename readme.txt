
Flamed v4 - [rm]


contents:

1. Preparing the demos
2. How Flamed works
3. F.A.Q.
4. Credits



>>> 1. Preparing the demos


1) Open FlamedLib.vbp and compile it.

2)    For Windows 98/Me: Make sure you have installed an ASPI Layer.
                         I heard a good one should be "Adaptec ASPI v4.70".
   For Windows NT/2k/XP: If you have administrator rights, just run the demos.
                         If not, Make sure you have installed an ASPI Layer.


3) Run the demos.




>>> 2. How Flamed works

Low Level CD/DVD communication is about sending and recieving packets of data.
You simply send a packet like:

Inquiry packet (6 bytes):

Byte | Bit 7 | Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 | hex
 0      0        0      0       1       0       0        1      0      12h   Operation Code
 1      0        0      0       0       0       0        0      0       0h
 2      0        0      0       0       0       0        0      0       0h
 3      0        0      0       0       0       0        0      0       0h
 4      0        0      0       0       0       0        0      0       0h
 5      0        0      0       0       0       0        0      0       0h

The inquiry OpCode will cause the drive to return some information about itself.
e.g. vendor, product id, revision and other stuff.

The BIG question: How to know which packets exist and what's their parameters?

The T10 commitee designs drafts for SCSI units:
http://t10.org/

The only drafts you need for CD-ROM/R/RW are MMC-3 and SPC-2,
for DVD MMC-4 and MMC-5.

So, I told you about SCSI, but what's about IDE/ATAPI?
IDE is very similar to SCSI, because it was designed to be very compatible.
The most significant difference is that every ATAPI command packet
has 12 bytes, whereas SCSI packets have 6, 10, 12 and sometimes even 16 bytes.

BUT: We don't have to argue with that!
SCSI Layers will do the job for us. Well, most of them.

The 2 most popular ones are the ASPI Layer developed by Adaptec and
the SPTI, a DeviceIoControl function included in the newer versions
of windows (NT/2K/XP).

Both of them allow communicating with CD/DVD-ROM drives like they all had
a SCSI interface.
The SPTI should have a wider support for physical interfaces like USB and firewire.
I don't know about the interfaces support of the Adaptec ASPI.

Anyways, after sending a packet you only have to wait for the drive to respond,
and read the data.

Pretty simple, eh?




>>> 3. Frequently asked quenstions

1.  What does ASPI and SPTI stand for?
2.  Nice and good, but VB Accelerator has some true low level CD burning stuff!
3.  It tells me: "Write Parameters Mode page couldn't be sent!"
4.  If I select my drive, it just gives a bunch of errors!
5.  DOS can't read the burned disks!
6.  Will you add DVD support?
7.  Sometimes my app suddenly crashes at design time?!
8.  I use analog CDDA playing and do not hear anything.
9.  What about adding progress to CD-RW blanking?
10. It says "No interfaces found". What does that mean?
11. What about drive compatibility?
12. Windows says the disk is empty after I've written to it.
13. "Could not initialize encoder" ?


1. What does ASPI and SPTI stand for?

ASPI - Advanced SCSI Programming Interface
SPTI - SCSI Pass-Through Interface


2. Nice and good, but VB Accelerator has some true low level CD burning stuff!

Argh, no!
Just because it looks complicated and there's a big fat "VB Accelerator"
above it, it doesn't mean it's perfect!
(You may guess how may variations I heard of that sentence...)
Let me explain:
Windows XP brought a new feature - The built in CD writing service,
called IMAPI (Image Mastering API).
As it is a service, you can program it.
For this Microsoft added an easy to use (:P) interface for intergrating
that service into your apps (Yes, even more dependencies on XP!!).
Means, you can NOT use it with other Windows versions.
And it is not low level, it just uses the IMAPI service, which can also be deactivated...
NOTHING which runs in Ring 3 is low level! Not even Flamed v4! :)


3. It tells me: "Write Parameters Mode page couldn't be sent!"

Your drive isn't MMC compilant and so not compatible to the Flamed project.


4. If I select my drive, it just gives a bunch of errors!

1) You have a USB/Firewire drive and use the ASPI.
2) A nasty bug in clsCDROM.
3) Your drive is too old and does not support the commands.


5. DOS can't read the burned disks!

FL_ISO9660Writer is far away from being somewhat strict.


6. Will you add DVD support?

No. I don't even have a DVD reader.
But you can.


7. Sometimes my app suddenly crashes at design time?!

You probably hit the stop button to often and used a
subclassed class of Flamed Lib.

Subclassed classes (windows created in classes):
FL_DoorMonitor
FL_FreeDB (Winsock)
FL_CDPlayer


8. I use analog CDDA playing and do not hear anything.

Either your drive isn't connected to the sound card,
or it does not support analog playback (<= must be pretty old).
You probably already checked the volume ;)


9. What about adding progress to CD-RW blanking?

To be honest, I don't see any way of doing this atm.
Will work on it...


10. It says "No interfaces found". What does that mean?

Windows 9x/Me: You need an ASPI driver.
Windows NT/2K/XP: You need administrator priviledges or
                  an ASPI driver.


11. What about drive compatibility?

Difficult topic.
Some drives have buggy firmwares which return invalid data,
drives have different physical interfaces wich have a
different behaviour, ...
I don't want to argue with that,
I just keep it as generic as I can.
Have a look at the CDRecord source code if you want
to add support for more drives ;)


12. Windows says the disk is empty after I've written to it.

This is a bug (?) which already many people have reported.
I can't reproduce this, it works for me.
Sometimes it helps to quit VB/your App and reload the disk
or simply reboot.


13. "Could not initialize encoder" ?

Flamed uses the ACM to encode/decode MP3.
Therefore you need an ACM MP3 Codec (Lame is a good one).
If the error still occurs try different bitrates.





>>> 4. Credits

http://hochfeiler.it/alvise/ - introduction to ASPI in VB
                             - simple CD-R writing with ASM and ASPI

http://www.codeproject.com/  - UDF CD writing with ASPI

http://t10.org/              - SCSI drafts

http://www.activevb.de/      - fantastic community

http://www.vbarchiv.de/      - nice tipps'n'tricks

http://www.vb-fun.de/        -        "

http://vbforums.com/         - good community

http://club.cdfreaks.com/    - interesting homemade tools

http://www.cdtool.pwp.blueyonder.co.uk/ - SPTI sources

http://pscode.com/           - huge collection of projects

http://www.moxcdburn.net/    - open source CD writing library

http://www.ecma-international.org/publications/standards/Ecma-119.htm - ISO9660 specification

http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=59434&lngWId=1 - pretty safe sub classing

http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=55304&lngWId=1 - Mixer control class

http://foren.activevb.de/cgi-bin/foren/view.pl?forum=13&msg=834&root=834&page=1 - using the ACM from VB


Special thanks:

To DaGreeza for testing the lib




[rm]

.
