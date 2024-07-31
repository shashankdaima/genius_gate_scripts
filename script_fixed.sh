#!/bin/bash

for file in ./documents/*.docx; do
    filename=$(basename "$file" .docx)
    echo "Processing $filename"
    mkdir "output/$filename"
    mammoth "$file" --output-format=markdown --output-dir="output/$filename" 
    mv "output/$filename/$filename.html" "output/$filename/$filename.md"
done