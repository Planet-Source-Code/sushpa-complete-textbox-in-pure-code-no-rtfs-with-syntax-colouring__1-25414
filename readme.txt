TextBox Demo - Java & C/C++ Code Editor
Originally by Rang3r, modified by Sushant Pandurangi
More code and wonders at sushantshome.tripod.com

This custom TextBox is written in pure code. It looks like a standard textbox, but it has loads and loads of features
such as comment, string and keyword colouring. Additional graphic effects can be given by you using the Draw() 
feature (which you will see later.) This demo application shows you how to use the OCX, in the form of an editor for 
Java source code files. It allows you to open, edit and manipulate java files. Since having had a look at this code, at 
first sight, it seemed pretty difficult to sort of make some additions to the OCX. It was a brilliant attempt by Rang3r 
to write this (and even more brilliant to let me continue it!); however I feel some of my additions might make you a 
bit more interested in this. Rang3r wrote it in VB, but hopes to convert it to C++ sometime.

Tips and tricks:

* Set a picture for the 'BBuffer' PictureBox in the OCX. This will serve as a background picture for the TextBox!
* While using the DRAW event, you can draw lines under specific words, draw icons, etc. with loads of functions:
	a. Line
	b. PaintPicture
	c. Circle
	d. PSet, Point
In addition, you can also specify the default PictureBox properties for the Canvas object that is passed in the
Draw event.

Some modifications by me:

* Comment colouring
* Save files feature
* Number colouring
* The 'Text' property
* Evenly spaced lines

1. Comment colouring
This was originally written to parse and highlight VB keywords and stuff. However, since there was no provision for 
comment colouring using the VB ' symbol, I decided to put in this feature using Java & C/C++ (and also other) languages' 
system of using // and /* */ comments. This TextBox now interprets both types of comments.

2. Saving & the 'Text' prop
The original box lacked one more important feature - a Text property, as in other TextBoxes. When you typed or loaded 
a file, the text was drawn to the screen, and not, well, typed, so it was quite important to have a 'Text' property. The 
one I implemented might have faster alternatives, if you see, but it's good enough to work for the time being. It also forms 
the base of the new 'Save' function. (doh, it saves files to disk - read below as well)

3. Evenly spaced lines
The previous version used 16 pixels as the granted text height (I changed this to BBuffer.TextHeight("X") and everything seems fine now).

Syntax highlight demo:

//This is a one line comment. This is a one line comment. Java is nice. Java is nice. Rang3r is a genius. So am I.
The next line is in the default colour.

/*This is a multiline comment and it spawns lines and lines of text and contains info.
 */

public class Readme extends Rang3r implements Sushant
{
private int nothing;

Readme() {}

public static void main() {}

private int getData() {
do {(new Readme()).setVisible( true )} while 1;
}

}

Sushant Pandurangi, 6 Jul 2001. [sushant@phreaker.net]
Below you find the original readme.txt written by Rang3r.

2000-09-21

this is a demo of my textbox control...
its entierly coded in vb but it is supposed to be converted to c++ some day. it supports syntax highlighting (wothout flickering
as rtf boxes do) and allso a "DRAW" event-(if you mark a word as keyword=true) that word will raise the draw event 
every time it is drawn to screen and when that is done you can draw extra graphics on the textbox (ie lines under 
"end sub" etc or icons or whatever)

features:

* Syntax highlighting
* Draw extra graphics event
* Strings ( it colors everything inside a string in a separate color )
* open method (doh , it loads a text into the box)

features that is supposed to come:

* selections (stream and box)
* different fontstyles (so you can draw some words bold and some italic etc)
* cut-copy-paste

you can contact me at:
* ROGER.JOHANSSON@SNABBMAT.SE
* fabbor_roger@hotmail.com

id really like if you guys out there try to improve it, and if you do... please send me a copy of the new version

things to improve:

* faster rendering speed when rendering whole screen... (scrolling with the scrollbars)
* horizontal scrolling when writing outside of the screen







