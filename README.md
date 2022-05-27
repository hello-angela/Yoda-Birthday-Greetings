# Yoda Birthday Emails

A script that takes in an excel spreadsheet with the name,email and birthday (YYYY-MM-DD format) and converts it to CSV.
Results are compared to current datetime, and if the days match it's assumed that it's the person's birthday.

A random birthday image is then selected from the existing list, and attached to an email body before it is sent.

Works with Gmail so long as security settings are adjusted for 3rd party applications.


# How to run it locally

Clone the project

```bash
  git clone https://github.com/hello-angela/Yoda-Birthday-Greetings.git
```

Go to the project directory

```bash
  cd my-project
```

Install dependencies

```bash
  pip install -r .\requirements.txt
```

Run the script

```bash
  python3 main.py
 ```
