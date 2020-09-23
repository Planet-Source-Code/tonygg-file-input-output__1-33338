<div align="center">

## File Input Output


</div>

### Description

File Input Output Append
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[TonyGG](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tonygg.md)
**Level**          |Intermediate
**User Rating**    |4.3 (34 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tonygg-file-input-output__1-33338/archive/master.zip)





### Source Code

```
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML>
<HEAD>
<TITLE>File and Directory - Articles - Handling Files in Visual Basic - In, Out, Shake It All About</TITLE>
</HEAD>
<BODY LeftMargin=0 TopMargin=0 Marginheight="0" Marginwidth="0" bgcolor="#FFFFFF">
<TABLE width=100% align=left bgcolor="#FFFFFF">
 <TBODY>
 <TR>
 <TD>
  <FONT face="Arial, Helvetica" color=#003399 size=+1><B>Handling Files in Visual Basic</B></FONT></P>
   <P><I>By Alex Allan</I></P>
   <P><FONT face="Arial, Helvetica" color=#003399 size=+1><B>In, Out,
   Shake It All About</B></FONT></P>
   <P>First off, let's take a look at the file modes input, output and
   append. You use these three types to read or write plain text - such
   as that found in .txt, .bat and .ini files.</P>
   <P>But when should you use each mode? Well, it depends on what you
   want to do. Use the following list to help you decide...</P>
   <UL>
    <LI>The <B>Output</B> mode creates a blank file and allows you to
    write information into it.
    <LI>The <B>Append</B> mode is similar to the Output mode but
    appends (adds to) an existing file.
    <LI>The <B>Input</B> mode opens a file for reading. </LI></UL>
   <P><B><I>Top Tip:</B> You may hear these file modes being referred
   to as 'sequential files'. That's 'cause once you have read or
   written to a line, you can't go back to it unless you close and
   re-open. In other words, the modes are one way – sequential.</I></P>
   <P>So, for example, to open a file for <B>output</B>, you'd use the
   Open statement like this:</P></FONT><PRE><FONT color=#00007f>Open</FONT> "c:\windows\faq.txt" <FONT color=#00007f>For</FONT> Output <FONT color=#00007f>As</FONT> #1</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>That's all fine and dandy, but how do you <B>use</B> each mode
   after you’ve <B>open</B>ed the file?</P>
   <P>To <B>write</B> to a file we use (Output and Append
   only):</P></FONT><PRE><FONT color=#00007f>Print</FONT> #filenumber, expression</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>To <B>read</B> from a file we use (Input only):</P></FONT><PRE>Input #filenumber, variablelist</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>This probably looks completely confusing at the moment, so let's
   figure out what it all means.</P>
   <P>Let's imagine you've opened a file for output, like
   this...</P></FONT><PRE><FONT color=#00007f>Open</FONT> "c:\groovy\myfile.txt" <FONT color=#00007f>For</FONT> Output <FONT color=#00007f>As</FONT> #1</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>...we now want to put information into this file using the Print
   statement, like this:</P></FONT><PRE><FONT color=#00007f>Print</FONT> #1, "Hello World!"</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>This inserts the information you pass it direct to the file in
   #1.</P>
   <P>If you'd opened a file for Input, like this...</P></FONT><PRE><FONT color=#00007f>Open</FONT> "c:\groovy\myotherfile.txt" <FONT color=#00007f>For</FONT> Input <FONT color=#00007f>As</FONT> #1</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>...you can read information from the file, like
this...</P></FONT><PRE>Input #1, MyVariableName</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>This reads information from the file in #1 and puts it into your
   variable. </P>
   <P>Let's use an example - it's easier to explain that way!</P>
   <P><FONT face="Arial, Helvetica" color=#003399 size=+1><B>Building a
   Sample</B></FONT></P>
   <P>Let's build a sample to demonstrate accessing files:</P>
   <P>Open Visual Basic and double-click on "Standard EXE"</P>
   <P>You should be left with a blank form.</P>
   <UL>
    <LI>Throw a simple Text Box onto the form
    <LI>Go to the Properties window and change MultiLine to True
    <LI>Create two Command buttons, setting the caption of the first
    to Read and the second to Write. </LI></UL>
   <P>So far, it should look something like this:</P>
   <P align=center><IMG height=217
   src="http://home.iprimus.com.au/ganino/image2.gif"
   width=228 border=1></P>
   <P>Now, double-click on Read and insert the following
   code:</P></FONT><PRE><FONT color=#00007f>Private</FONT> <FONT color=#00007f>Sub</FONT> Command1_Click()
 'Outline:  - Asks the user for a file.
 <FONT color=#007f00>'    - Reads all the data into Text1.</FONT>
 <FONT color=#00007f>Dim</FONT> FilePath <FONT color=#00007f>As</FONT> <FONT color=#00007f>String</FONT>
 <FONT color=#00007f>Dim</FONT> Data <FONT color=#00007f>As</FONT> <FONT color=#00007f>String</FONT>
 FilePath = InputBox("Enter the path for a text file", "The Two R's", _
 "C:\WINDOWS\WINNEWS.TXT")
 'Asks the user for some input via an input box.
 <FONT color=#00007f>Open</FONT> FilePath <FONT color=#00007f>For</FONT> Input <FONT color=#00007f>As</FONT> #1
 'Opens the file given by the user.
 <FONT color=#00007f>Do</FONT> <FONT color=#00007f>Until</FONT> <FONT color=#00007f>EOF</FONT>(1)
 'Does this loop until <FONT color=#00007f>End</FONT> Of File(<FONT color=#00007f>EOF</FONT>) for file number 1.
  Line Input #1, Data
  'Read one line and puts it into the varible Data.
  Text1.Text = Text1.Text & vbCrLf & Data
  'Adds the read line into Text1.
  MsgBox <FONT color=#00007f>EOF</FONT>(1)
 <FONT color=#00007f>Loop</FONT>
 <FONT color=#00007f>Close</FONT> #1
 'Closes this file.
<FONT color=#00007f>End</FONT> Sub</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>Read the comments (anything prefixed by an apostrophe).</P>
   <P>Next, double-click on Write and insert the following
   code:</P></FONT><PRE><FONT color=#00007f>Private</FONT> <FONT color=#00007f>Sub</FONT> Command2_Click()
 'Outline:  - Asks the user for a file.
 <FONT color=#007f00>'    - Writes it all to a the file.</FONT>
 <FONT color=#00007f>Dim</FONT> FilePath <FONT color=#00007f>As</FONT> <FONT color=#00007f>String</FONT>
 FilePath = InputBox("Enter the path for a text file", _
 "The Two R's")
 'Asks the user for some input via an input box.
 <FONT color=#00007f>Open</FONT> FilePath <FONT color=#00007f>For</FONT> Output <FONT color=#00007f>As</FONT> #1
 'Opens the file given by the user (for Output).
 <FONT color=#00007f>Print</FONT> #1, Text1.Text
 'Writes the data into the file number #1
 <FONT color=#00007f>Close</FONT> #1
<FONT color=#00007f>End</FONT> Sub</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>Once again, read all the comments. Do you understand what's
   happening?</P>
   <P>Hit F5 to run your program!</P>
   <P>Congratulations! You've just created a simple text editor. The
   Write button performs a simple 'Save As' whilst the Read button is
   the equivalent of 'Open'.</P>
   <P><FONT face="Arial, Helvetica" color=#003399 size=+1><B>Didn't I
   Mention Those?</B></FONT></P>
   <P>You may have noticed a few things in my code that I haven't told
   you about. Let's take a peek at a few geeky code words...</P></FONT><PRE>Line Input #filenumber, MyVariableName</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>- This is the same as Input but it reads the whole line instead
   of stopping at a comma (which can be useful - sometimes)</P></FONT><PRE><FONT color=#00007f>EOF</FONT>(#filenumber)</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>- This outputs a true or false value, depending on whether it has
   hit the 'end' of the file. In my code, I used it in the Do…Loop for
   Read. When EOF=True, the loop ends.</P></FONT><PRE><FONT color=#00007f>Close</FONT> #filenumber</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>- This closes the file. It allows other files to open this
   file.</P>
   <P>And if you'll be getting real friendly with files in Visual
   Basic, here are a few other file writing functions you may be
   interested in:</P></FONT><PRE>Write #filenumber, outputlist</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>This is similar to Print but it writes "screen formatted data"
   instead of raw data to a file. This means that values (numbers) have
   hashes ("#") put around them and strings have quote marks put around
   them. This makes the text less human readable but easier to read for
   your programs.</P></FONT><PRE>LOC(#filenumber)</PRE><FONT face="verdana, arial, helvetica"
   size=-1>
   <P>The LOC function returns the current read/write position within
   an open file.</P></FONT><PRE><FONT color=#00007f>LOF</FONT>(#filenumber)</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>The LOF function gives you the length of the file
open.</P></FONT><PRE>FreeFile</PRE><FONT face="verdana, arial, helvetica" size=-1>
   <P>This will give you next available file number.</P>
   <P>Example:</P></FONT><PRE>MyFileNumber = FreeFile
<FONT color=#00007f>Open</FONT> "c:\windows\faq.txt" <FONT color=#00007f>For</FONT> Input as #MyFileNumber</PRE><FONT
   face="verdana, arial, helvetica" size=-1></FONT><PRE>Input$(number_of_chars_to_return, #filenumber)</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>This is the same as Input and Line Input except it has no limits
   (except the ones you set in number_of_chars_to_return). By using LOF
   - this can be used to get the whole file. We could replace all the
   stuff in the Do…Loop in the Read button, with:</P></FONT><PRE>Text1.Text = Input$(<FONT color=#00007f>LOF</FONT>(1),#1)</PRE><FONT
   face="verdana, arial, helvetica" size=-1>
   <P>Why not have a play around and see what these do? Go on, have a
   go!</P>
   <P>That's about it for the Input, Output and Append modes.
</TD></TR></TBODY></TABLE>
</BODY></HTML>
```

