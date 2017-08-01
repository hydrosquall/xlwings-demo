# Xlwings-Demo

Experiments in integrating Microsoft Excel with Python

Roadmap

1. Hello world
2. Use Pandas Datareader to load iris dataset
3. Use Pandas Datareader to load Enigma dataset

## Setup

Once you are inside a virtual environment, do the following:

    make install

### On Mac

Manually install the xlwings add-in, and restart the notebook.

Try out `xlwings quickstart`

With the generated notebook, put in the path to a virtualenv.

### Limitations

- Only on Windows can you define callable functions. In mac, you can only
click a button to run functions, and you're not able to call any params

In my opinion, it looks like you will have an easier time updating Excel
spreadsheet from python, than you have  trying to call Python from excel