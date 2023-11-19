# Use an official Node.js runtime as a base image
FROM node:14

# Set the working directory within the container
WORKDIR /app

# Copy package.json and package-lock.json to the working directory
COPY package*.json ./

# Install application dependencies
RUN npm install

# Copy the rest of the application code to the working directory
COPY . .

# Expose the port your app will run on
EXPOSE 8080

# Define the command to start your Node.js application
CMD ["node", "app.js"]
