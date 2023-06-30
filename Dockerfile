# Start with an official Node.js runtime as a parent image
FROM node:18

# Set the working directory to /app
WORKDIR /app

# Install Python and other dependencies
RUN apt-get update && \
    apt-get install -y python3 python3-pip

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