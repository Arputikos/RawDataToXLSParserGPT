# RawDataToXLSParserGPT - PL version
Parsing raw text into excel sheet - extracting specific information and features from the text to the table.

# How to use
Create an xlsx file that will have some parameters/features in the first column. (First entry should be "Features").
Then save it and close.
Open input.py, insert your OpenAI API key and text to process (like raw text from a website).
Edit prompts as needed.
The GPT4 should return json in "row_name" : "found_parameter_info" format and automatically add new column to the xls file with all those features found. It will add "???" if not found.
The example may be tweaked to process all kinds of informations from the text. It's basically an AI GPT based text parser.
