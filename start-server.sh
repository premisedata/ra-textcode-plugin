#!/bin/bash

# Quick script to start a local server for testing
# Usage: ./start-server.sh

echo "Starting local server for CodeText Word Add-in..."
echo ""
echo "Note: Word Add-ins require HTTPS or localhost HTTP"
echo "This script starts an HTTP server - if it doesn't work,"
echo "you may need to use office-addin-dev-certs for HTTPS"
echo ""

# Check if Python is available
if command -v python3 &> /dev/null; then
    echo "Starting Python HTTP server on port 3000..."
    echo "Access at: http://localhost:3000/taskpane.html"
    echo ""
    echo "Press Ctrl+C to stop the server"
    echo ""
    python3 -m http.server 3000
elif command -v python &> /dev/null; then
    echo "Starting Python HTTP server on port 3000..."
    echo "Access at: http://localhost:3000/taskpane.html"
    echo ""
    echo "Press Ctrl+C to stop the server"
    echo ""
    python -m http.server 3000
else
    echo "Python not found. Trying Node.js http-server..."
    
    if command -v npx &> /dev/null; then
        echo "Starting http-server on port 3000..."
        echo "Access at: http://localhost:3000/taskpane.html"
        echo ""
        echo "Press Ctrl+C to stop the server"
        echo ""
        npx http-server -p 3000
    else
        echo "Error: Neither Python nor Node.js found."
        echo "Please install Python or Node.js to run a local server."
        exit 1
    fi
fi

