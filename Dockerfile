FROM python:3.11-slim

# Install Java (required for tabula-py)
RUN apt-get update && apt-get install -y --no-install-recommends \
    default-jre-headless \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install dependencies
COPY ConverterApp/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY ConverterApp/ .

# Create required directories
RUN mkdir -p uploads logs

# Expose port
EXPOSE 5000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:5000/health || exit 1

# Run the application
CMD ["python", "pdftoexcel.py"]
