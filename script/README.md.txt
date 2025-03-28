
Python script that will take a list (of undefined length) of words separated by new lines (return, enter), and generate a 15x15-letter word search puzzle for
 every 10 words listed. The word searches are spat out in a .docx file in the same directory. One page is the puzzle, the following page is the solution,
 with words (letter cells) highlighted. The words used in the puzzle are displayed under the letter grid.
----------------------------------------------------------------

Install the python-docx library: pip install python-docx

Update the SCRIPT_DIR variable with the correct path to your script directory: SCRIPT_DIR = r"C:\Users\<username>\Desktop\script"

Adjust the FONT_SIZE variable to your desired font size: FONT_SIZE = Pt(14)  # Change this value

Save the script as word_search_generator.py in the specified directory.

Place your words.txt file in the same directory.

Run the script: python word_search_generator.py