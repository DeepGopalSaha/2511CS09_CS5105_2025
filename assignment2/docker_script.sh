#!/bin/bash

# Stop any running container with the same name
docker rm -f assignment2-app-container 2>/dev/null

# Build the Docker image
docker build -t assignment2-app ./assignment2

# Run the container
docker run -d \
    --name assignment2-app-container \
    -p 8501:8501 \
    -v "$(pwd)/assignment2/Output:/app/Output" \
    assignment2-app

