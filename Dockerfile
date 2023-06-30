# Start with an official Node.js runtime as a parent image
FROM node:18
FROM python:3.8
FROM ubuntu:18.04

# Install Python and other dependencies
RUN apt-get update -y
RUN apt-get install -y python3
RUN apt-get install -y python3-pip

COPY . /app
WORKDIR /app

# Install Node npm
RUN apt-get update && apt-get upgrade -y && \
    apt-get install -y nodejs \
    npm

RUN echo "Node: " && node -v
RUN echo "NPM: " && npm -v

# Copy the package.json and package-lock.json files to the working directory
COPY package*.json ./

# Install the Node.js dependencies
RUN npm install

# Install Python dependencies
COPY requirements.txt .
RUN pip3 install -r requirements.txt

# Start the application
CMD [ "npm", "start" ]