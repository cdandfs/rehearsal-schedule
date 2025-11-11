#!/bin/bash

# Set input and output directories
INPUT_DIR="student-schedules"
OUTPUT_DIR="student-schedules"

# Create output directory if it doesn't exist
mkdir -p "$OUTPUT_DIR"

# Loop through Markdown files in input directory
for file in "$INPUT_DIR"/*.md; do
  # Get file name without extension
  filename=$(basename "$file" .md)
  quarto render "$file"

done