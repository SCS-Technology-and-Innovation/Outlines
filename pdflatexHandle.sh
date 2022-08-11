#!/bin/bash
pdflatex --output-directory=tmp/ $1
pdflatex --output-directory=tmp/ $1
pdflatex --output-directory=tmp/ $1
echo $1
FILENAME=`echo "$1" | cut -d'/' -f2`
NAME=`echo "$FILENAME" | cut -d'.' -f1`
EXTENSION=`echo "$FILENAME" | cut -d'.' -f2`
cp tmp/$NAME.pdf pdfs/$NAME.pdf

rm tmp/$NAME.pdf
rm tmp/$NAME.toc
rm tmp/$NAME.aux
rm tmp/$NAME.log
rm tmp/$NAME.out
rm tmp/$NAME.tex
echo "$NAME done"

