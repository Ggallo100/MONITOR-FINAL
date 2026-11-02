# Versión web multiusuario del Sistema POA - DPDI
FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 POA_DATA_DIR=/datos POA_PORT=8080
WORKDIR /app
COPY requirements.txt requirements-web.txt ./
RUN pip install --no-cache-dir -r requirements-web.txt
COPY . .
RUN useradd --create-home poa && mkdir -p /datos && chown poa:poa /datos
USER poa
VOLUME ["/datos"]
EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8080/salud')"
CMD ["python", "servidor_web.py"]
