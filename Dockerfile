# Start with an official Node.js runtime as a parent image
FROM node:18

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
RUN echo "python3: " && python3 -v
RUN echo "PIP: " && pip3 -v

# Copy the package.json and package-lock.json files to the working directory
COPY package*.json ./app

# Install the Node.js dependencies
RUN npm install

# Install Python dependencies
COPY requirements.txt ./app
RUN pip3 install -r requirements.txt

# Start the application
CMD [ "npm", "start" ]