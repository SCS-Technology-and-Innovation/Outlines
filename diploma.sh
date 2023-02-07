rm -f *-*.tex
rm -f *-*.pdf
mkdir -p studyplan
rm -f studyplan/*.pdf
python3 diploma.py
for file in `ls -1 *-*.tex`;
do
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
    pdflatex --interaction=batchmode $file 2>&1 > /dev/null;
done
mv studyplan*pdf studyplan
rm *-*.out
rm *-*.log
rm *-*.aux
