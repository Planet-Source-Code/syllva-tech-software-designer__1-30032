	-Dedicated List Load/Save Fileformat-
		by Samuel Truscott

1.0: Introduction
2.0: Example of Fileformat
3.0: Usage in Visual Basic
4.0: Contact Info

1.0: Introduction

I made this because every load/save list module/code i try just fails.
The purpose of this code is that it's in a special format which means
you can use it for just your application.

This could be used for example, in a media player as a playlist code.

I've even put in errorhandlers in case the files don't exist, in which
case an item will be added to the list saying "Error: No data in list!".

2.0: Example of Fileformat

Here's a code example:

[List Header]
Total=3

[List]
1=Item 1
2=Item 2
3=Item 3

This is a lot more stable than the common load/save commands because it
uses single string code, so each string is saved one-by-one and then a
total is set to say which of the lines to add into the List.

3.0: Usage in Visual Basic

To load a list:
Code: LoadList List, filename
E.g.: LoadList List1, "C:\mylist.txt"

To save a list:
Code: SaveList List, filename
E.g.: SaveList List1, "C:\mylist.txt"

4.0: Contact Info

Email: samst@btinternet.com
Website: http://www.btinternet.com/~samst/