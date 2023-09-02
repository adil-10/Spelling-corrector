# Spelling-corrector

This script is designed to iterate over all Word documents in a folder and correct any spelling mistakes it finds using the most suggested correction.

## Usage

1. Place your Word documents (.docx files) in the "Test" folder.
2. Run the script using Python.
3. The script will correct the spelling in each document and save the modified version.

## Requirements

- Python 3.x
- The `docx` library for working with Word documents. You can install it using `pip install python-docx`.
- The `autocorrect` library for spell correction. You can install it using `pip install autocorrect`.

## How it works

- The script iterates through each Word document in the "Test" folder.
- It extracts the text from each paragraph.
- For each paragraph, it corrects the spelling using the `autocorrect` library.
- It replaces the original paragraph text with the corrected text.
- The modified document is then saved in place.

## Note

Make sure to back up your documents before running the script, as it will modify the files.

Enjoy spell-correcting your Word documents!
