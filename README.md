# excelSciNot

A lightweight GUI for converting cells into a readable scientific notaion with the ability to search a highlighted cell on Wolfram Alpha.

I found this useful in my excel using adventures as I do not think the scientific notation that is given by excel is that legible.

Note: Decimal Precision can be easily changed in the main.py file line 13.

## Usage

##### PLEASE NOTE THIS SMALL WIDGET ONLY WORKS FOR WINDOWS

1. Installation
   1. Lastest Windows executable can be found on Github Actions - [HERE](https://github.com/uuid1/excelSciNot/actions/runs/8309722154/artifacts/1332131231 "windows executable")
   2. Compile from source
      1. Clone the Repository:
         1. `git clone https://github.com/uuid1/excelSciNot.git cd excelSciNot`
      2. Install Dependencies
         1. `pip install -r requirements.txt`
      3. Run the Application
         1. `python main.py` (note main.py in src directory)
2. Functionality
   * Please make sure to have a excel window open before launching this app
   * Select cells in the Excel file window
   * Click the "Convert to Scientific Notation" button to convert selected cells to scientific notation
   * Click the "Search on Wolfram Alpha" button to perform a Wolfram Alpha search on the selected cell
   * View all selected cells in the list

## Contributions

Contributions are welcome! If you would like to add new features, fix bugs, or improve this thing in any way feel free to fork the repository
