# Gravience Outline

### The true point of Gravience is to show Visual Basics true powers and denote them as stupid

VBS was used in the *I love you* virus in the early 2000s; it was a terrible virus that used VBS' stupid escalation. 

Gravience outlines this stupidity; VBS is still an extrodinarily powerful language that can run system code from the very visual basics.

The outline:

1. *payload.vbs* - This is a mutable file but it essentially loads Gravience core systems. This is a decoy; ideally, files would be downloaded from the website
and this would be packaged in a more secretive way, but I don't intend to make this virus dangerous; that would be bad and illegal (even though it borders on illegality)

2. *taskii.vbs* is ran by *payload.vbs*; this is an immutable file that allows all of Gravience to function. It's commented in detail and its main method explains most of it.

The most important part about Taskii that you need to know about is that, with very little warning, it adds 
taskii as a system startup application; this is extremely dangerous code that, if ran as 
admin, would add this file as a system registry key. 

Taskii also runs other tasks; it is a modular program. This way, if the HSKEY registry is reset, taskii always checks if it's on the registry
before performing its modular tasks. 

3. Taskii runs rick.vbs which opens a YouTube video, plays the rickroll for 10 seconds, then shuts the browser. 

This isn't a fully working and bug free feature as it can close a browser tab you don't want to. 