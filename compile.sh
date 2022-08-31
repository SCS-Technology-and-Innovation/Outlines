rm *-*.tex
rm *-*.pdf
python3 outline.py
for file in `ls -1 *.tex`;
do

    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
done
rm *-*.toc
rm *-*.out
rm *-*.log
	    
