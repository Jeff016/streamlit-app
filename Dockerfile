# Use official Python image
FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Copy requirements and app files
COPY requirements.txt ./
COPY app.py ./

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Optional: Add health check
HEALTHCHECK CMD curl --fail http://localhost:8501 || exit 1

# Expose Streamlit port
EXPOSE 8501

# Run Streamlit app
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
