FROM python:3.9

WORKDIR /dep

RUN apt update
RUN apt install -y texlive-latex-base texlive-fonts-recommended texlive-fonts-extra texlive-latex-extra


COPY requirements.txt /dep


RUN pip install --no-cache-dir --upgrade -r /dep/requirements.txt


WORKDIR /app

CMD ["uvicorn", "main:app", "--host", "0.0.0.0"]