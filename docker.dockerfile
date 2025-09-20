# Use an official Python runtime as a parent image
FROM python:3.9-slim
# Set the working directory in the container
WORKDIR /main
# Copy the current directory contents into the container at /app
COPY . /main
# Install any needed packages specified in requirements.txt
# (Assuming your code has dependencies and you've created a requirements.txt file)
RUN pip install --no-cache-dir -r requirements.txt
# Make port 8000 available to the world outside this container
EXPOSE 8000
# Run your script when the container launches
CMD ["python", "generateODS.py"]
CMD ["python", "Create_Config.py"]
CMD ["python", "main3.py"]