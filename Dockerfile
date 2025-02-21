# Use an official Python image as the base
FROM python:3.9-slim  

# Set the working directory inside the container
WORKDIR /app  

# Copy all files from your project directory to the container
COPY . .  

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt  

# Expose the port Flask runs on
EXPOSE 5000  

# Command to run the application
CMD ["python", "app.py"]
