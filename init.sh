#!/bin/bash
# title: "init"
# author: Hsieh-Ting Lin
# date: "2025-05-19"
# version: 1.0.0
# description:
# --END-- #
set -ue
set -o pipefail
trap "echo 'END'" EXIT

# Check if clasp is installed
if ! command -v clasp &>/dev/null; then
  echo "Error: 'clasp' is not installed or not in your PATH."
  echo "You can install it with: npm install -g @google/clasp"
  exit 1
fi

# Check if project is already initialized
if [ -f ".clasp.json" ]; then
  echo "An existing clasp project was detected."

  read -p "Do you want to start over and create a new Slides project? (y/n): " USER_CHOICE
  if [[ "$USER_CHOICE" =~ ^[Yy]$ ]]; then
    echo "Starting over..."
    rm -f .clasp.json
    rm -f appsscript.json
  else
    echo "Opening existing Slides project..."
    clasp open-container
    exit 0
  fi
fi

# Prompt user for the title of the Google Slides presentation
read -p "Enter the title for your new Google Slides presentation: " SLIDES_TITLE

# Check if the input is empty
if [ -z "$SLIDES_TITLE" ]; then
  echo "Error: Title cannot be empty."
  exit 1
fi

# Create a new Google Slides project
clasp create --type slides --title "$SLIDES_TITLE"
if [ $? -ne 0 ]; then
  echo "Failed to create the Slides project."
  exit 1
fi

# Push the code to the new project
clasp push
if [ $? -ne 0 ]; then
  echo "Failed to push the code."
  exit 1
fi

# Open the Google Slides presentation
clasp open-container
