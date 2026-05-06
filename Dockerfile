FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    texlive-xetex \
    texlive-lang-cyrillic \
    texlive-fonts-recommended \
    texlive-fonts-extra \
    texlive-latex-extra \
    texlive-science \
    fonts-liberation \
    fonts-dejavu \
    fonts-dejavu-extra \
    fontconfig \
    lmodern \
    && wget https://github.com/jgm/pandoc/releases/download/3.6.4/pandoc-3.6.4-1-amd64.deb \
    && dpkg -i pandoc-3.6.4-1-amd64.deb \
    && rm pandoc-3.6.4-1-amd64.deb \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN python3 create_reference.py
EXPOSE 5000
CMD ["python3", "app.py"]