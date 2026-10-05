FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY bot.py rep_to_txt.py ./
COPY md2gost/ md2gost/
RUN useradd --create-home --uid 1000 app && chown -R app:app /app
USER app

CMD ["python", "bot.py"]
