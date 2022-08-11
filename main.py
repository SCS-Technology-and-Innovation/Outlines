from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from functions.outlineHandler import outlineHandler
import os
import asyncio

app = FastAPI()



@app.post("/uploadOutlineExcel/")
async def create_file(outlineFile: UploadFile):
    file_location = f"tmp/{outlineFile.filename}"
    if outlineFile.filename != 'outline.xlsx':
        return {"message": 'file name must be outline.xlsx'}
    with open(file_location, "wb+") as file_object:
        file_object.write(outlineFile.file.read())
    outlineHandler(file_location)
    os.system('rm -f  pdfs/*.pdf')
    os.system('rm -f  pdfs/*.zip')
    for root, dirs, files in os.walk('tmp/'):
        for file in files:
            if file.endswith('.tex'):
                os.system(
                    f"./pdflatexHandle.sh tmp/{file}  > /dev/null 2>&1 & ")
    flag = 0
    while flag == 0:

        isFileExist = False
        await asyncio.sleep(1)
        print('waiting ...')
        for root, dirs, files in os.walk('tmp/'):
            for file in files:
                if file.endswith('.tex'):
                    isFileExist = True

        if isFileExist == False:
            flag = 1
    os.system(f"cd pdfs && zip outline.zip *.pdf && cd ..")
    os.system('rm -f  pdfs/*.pdf')

    file_path = "pdfs/outline.zip"
    return FileResponse(path=file_path, filename='outline.zip')
