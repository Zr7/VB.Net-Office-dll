VB.Net-Office-dll
=================

A dll library to create and make the use of excel in vb.net more easy 

The primairy goal was to create a library, in order to create projects in vb.net that
can interact with excel files.
The ease in this project is to avoid to repeat certain commands over and over again,
and instead use a single dll to handle those actions.
The other aspect in mind, was to reduce program file size. There is no need for
big hughe files, if they can be reduced to a minimal.

Why not a simple module ? 
it could well fit in a module, but that would increment your base file code, 
and make it harder to distribute over a mobile connection.
You also have to recompile the same code, while a single dll is allready build and ready
to use.
