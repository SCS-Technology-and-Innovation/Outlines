python3 outline.py
for file in `ls -1 *.tex`;
do
    pdflatex $file; pdflatex $file
done
	    
