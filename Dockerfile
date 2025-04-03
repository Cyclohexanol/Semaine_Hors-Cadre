# Dockerfile
FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Copy requirements file
COPY webapp/requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY webapp/ .

# Create necessary directories
RUN mkdir -p uploads results

# Expose port
EXPOSE 5000

# Environment variables
ENV FLASK_APP=app.py
ENV PYTHONUNBUFFERED=1
ENV FLASK_SECRET_KEY=change_this_to_a_secure_value_in_production

# Run with gunicorn for production
RUN pip install gunicorn

# Command to run the application
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app"]