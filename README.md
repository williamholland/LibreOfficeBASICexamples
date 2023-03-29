# LibreOffice BASIC - Examples

This is where I keep my notes and examples of LibreOffice BASIC macros

## MACRO Organisation

"
In LibreOffice, macros are grouped in modules, modules are grouped in libraries, and libraries are grouped in library containers. A library is usually used as a major grouping for either an entire category of macros, or for an entire application. Modules usually split functionality, such as user interaction and calculations. Individual macros are subroutines and functions. Figure 11 shows an example of the hierarchical structure of macro libraries in LibreOffice

Unless your macros are applicable to a single document, and only to a single document, your macros will probably be stored in the My Macros container. The My Macros container is stored in your user area or home directory.

If a macro is contained in a document, then a recorded macro will attempt to work on that document, because it primarily uses ThisComponent for its actions.
"

Once you're in the BASIC IDE, goto `Tools > Macros > Organize Macros > Basic` to create new libraries and modules.

## Debugging

I've found the best way to debug is to use breakpoints then type the name of the marker you're looking for in 'Watch:'

You can also make a simple warning dialog like this

```basic
Doc = ThisComponent
Print Doc.URL ' shows the local path to this file
```

## Links

* [Language Reference](https://wiki.documentfoundation.org/Documentation/BASIC_Guide#The_Language_of_LibreOffice_BASIC)
* [API Reference](https://api.libreoffice.org/docs/idl/ref/annotated.html)

## Examples TODO

### Writer

* open a file, write hello world, save file
* write in 2 different fonts and styles
* insert image
* loop through paragraphs and add numbers
* make a UI dialog
* log to terminal???
