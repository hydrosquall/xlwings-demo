# xlwings-demo

Excel Spreadsheets are user-familiar interfaces for working with many sorts of data. However,
writing macros/plugins in VBA involves stress-inducing and arcane syntax.

Python is an intuitive and powerful language for all sorts of programming tasks.

This repository collects basic examples of how these two powerful tools can be made to work together,
using the [`xlwings`](https://github.com/ZoomerAnalytics/xlwings) library as a bridge.


## Setup

Once you are inside a virtual environment, install all dependencies with the following:

    make install

### On Mac

Manually install the `xlwings` add-in, and restart Excel. You can get the path to the addin file by running `xlwings addin install`. 

Run `xlwings quickstart <name>`, replacing `<name>` with your project name. No special characters!

After opening the generated workbook, add the path to your virtualenv to Interpreter field
in the Excel xlwings Addins ribbon under "interpreter"

    # Example:
    /Users/cameron/Environments/xlwings/bin/python

## Demos

### helloworld

This is the basic quickstart demo that `xlwings` provides by default. 
Invoking the macro writes a basic string to the first worksheet.

### hellopandas

This demo lets users download data from the world's largest public data repository, [Enigma Public](https://public.enigma.com) directly into Excel. If you want to transform the data before it moves between the API and the sheet (i.e. perform groupby or filtering aggregations), you should edit the base `hellopandas.py` script.

Add the following to an `.env` file in the repository folder.

    ENIGMA_API_KEY=REDACTED


## Roadmap

- [x] Write hello-world example
- [x] Use Pandas to pipe a dataframe into Excel
- [x]  Use Pandas to pipe an Enigma table into Excel
- [x] Create examples of doing rollup aggregation in Python / link to pandas demos which can pipe the results into Excel readily.
    - Add example SQL queries, possibly create a blog post, using Enigma data
- [ ] Extend pandas-datareader to support the new Enigma API


## Limitations

- Only on Windows can you define callable functions (UDF). 
- On Mac, invoke functions by pressing `F5` in the VBA Editor, or by adding a button to your ribbon.
- It is easier to update Excel from Python than it is to invoke Python from Excel.
