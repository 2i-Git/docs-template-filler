# Docs Filler web app — small, single-process container for Azure Container Apps.
FROM python:3.12-slim

# Don't buffer stdout/stderr (so logs appear immediately) and don't write .pyc.
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=8000

WORKDIR /app

# Install dependencies first so this layer is cached across code changes.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Application code.
COPY docs_filler/ ./docs_filler/
COPY webapp/ ./webapp/

# Run as a non-root user.
RUN useradd --create-home appuser
USER appuser

EXPOSE 8000

# Azure Container Apps sets $PORT; default to 8000 for local runs.
CMD ["sh", "-c", "uvicorn webapp.main:app --host 0.0.0.0 --port ${PORT}"]
