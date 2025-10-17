#!/usr/bin/env bash
set -o errexit

echo "Starting build process..."

apt-get update
apt-get install -y poppler-utils ghostscript tesseract-ocr

pip install --upgrade pip
pip install -r requirements.txt

echo "Build completed successfully!"
