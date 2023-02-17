term=W23
if grep -q ListSchema sharepoint.csv; then
    sed '1d' sharepoint.csv > tmp
    mv tmp sharepoint.csv
fi
rm *-*.tex
rm *-*.pdf

mkdir -p forms
python3 outline.py
for file in `ls -1 ${term}*-*.tex`;
do
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
done
mv ${term}*pdf forms

mkdir -p sharepoint
python3 outline.py sharepoint
for file in `ls -1 ${term}*-*.tex`;
do
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
done
mv ${term}*pdf  sharepoint
rm *-*.toc
rm *-*.tex
rm *-*.out
rm *-*.log
rm *-*.aux
