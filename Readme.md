# XLS 2 TeX #

A simple and incomplete tool to convert XLS formatted spreadsheets into nice looking TeX tables. It is very much work in progress and comes with absolutely no guarantees whatsoever.

Run it either with `racket xls2tex.rkt` or compile using `raco exe xls2tex.rkt`.

## Why Would Anyone Need This? ##

I work on the experimental [Funcalc](http://www.itu.dk/people/sestoft/funcalc/) spreadsheet engine as part of my PhD project and typing spreadsheets in Latex by hand is boring.

## Left To Do ##

 - [x] Fix string quoting in TeX output.
 - [ ] Add margin coloring for interpreted and compiled sheets.
 - [ ] Add input, output and intermediate cell coloring.
 - [ ] Allow for specifying a cell range to display.
