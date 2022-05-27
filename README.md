Script that takes in an excel spreadsheet with the name,email and birthday (YYYY-MM-DD format) and converts it to CSV.
Results are compared to current day, and if the days match it's assumed that it's the person's birthday.

A random birthday image is then selected from the existing list, and attached to an email body before it is sent.

Works with Gmail so long as security settings are adjusted for 3rd party applications.

Only requirement to running the script is to install Numpy which can be done with <pip install -r .\requirements.txt>

Run script with <python3 main.py>