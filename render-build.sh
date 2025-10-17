#!/usr/bin/env bash
set -o errexit
echo "Starting build process..."
pip install --upgrade pip
pip install -r requirements.txt
echo "Build completed successfully!"
