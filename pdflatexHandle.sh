#!/bin/bash
pdflatex $1
pdflatex $1
pdflatex $1
NAME=`echo "$1" | cut -d'.' -f1`
EXTENSION=`echo "$1" | cut -d'.' -f2`
cp $NAME.pdf pdfs/$NAME.pdf

rm $NAME.pdf
rm $NAME.toc
rm $NAME.aux
rm $NAME.log
rm $NAME.out
rm $NAME.tex
echo "$NAME done"

