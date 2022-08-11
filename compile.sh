#!/bin/bash
python3 outline.py
for file in `ls -1 *.tex`;
do
    
    if [ $file != 'outline.tex' ]
    then
        ./pdflatexHandle.sh $file > /dev/null 2>&1 & 
        # ./pdflatexHandle.sh $file > /dev/null 2>&1  
        # ./pdflatexHandle.sh $file &
        # ./pdflatexHandle.sh $file
        echo $file
    fi
   
done
	    
