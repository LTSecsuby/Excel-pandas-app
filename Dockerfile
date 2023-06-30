# Start with an official Node.js runtime as a parent image
FROM node:18
FROM python:3.8


COPY . /app
# Set the working directory to /app
WORKDIR /app

# Install Python and other dependencies
RUN apt-get update -y
RUN apt-get install -y python-pip python-dev build-essential

# Copy the package.json and package-lock.json files to the working directory
COPY package*.json ./

# Install the Node.js dependencies
RUN npm install

# Copy the rest of the application code to the working directory
COPY . .

# Install Python dependencies
COPY requirements.txt .
RUN pip install -r requirements.txt

# Start the application
CMD [ "npm", "start" ]