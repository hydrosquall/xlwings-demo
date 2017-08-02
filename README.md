# Xlwings-Demo

Experiments in integrating Microsoft Excel with Python

Roadmap

1. Hello world
2. Use Pandas to pipe a dataframe into Excel
3. Use Pandas to pipe an Enigma table into Excel
4. Extend pandas-datareader to support the new Enigma API

## Setup

Once you are inside a virtual environment, do the following:

    make install

Then, add the following to a `.env` file in your project root to do the Enigma API demo.

    ENIGMA_API_KEY=REDACTED

### On Mac

Manually install the xlwings add-in, and restart Excel

Run `xlwings quickstart`

After opening the generated workbook, add the path to your virtualenv to Interpreter field
in the Excel xlwings options.


    Example:
    Interpreter = /Users/cameron/Environments/xlwings/bin/python


### Limitations

- Only on Windows can you define callable functions. In mac, you can only
click a button to run functions, and you're not able to call any params

In my opinion, it looks like you will have an easier time updating Excel
spreadsheet from python, than you have trying to call Python from excel