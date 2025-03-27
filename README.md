# JSON_Generator
This is a JSON generator programmed in Python. It automatically generates a Python file based on the information provided in the 'ProcessViewMessages.xlsm' Excel file.
---  
It's much better to run this program again on the Python engine; anyway:
If .exe is needed:
  To convert the .py file into a .exe file, run the following command in your Windows terminal:
  
  JSON Generator> pyinstaller --onefile --windowed jsonGenerator.py
  
  The pyinstaller library is required. If you don’t have it installed, you can install it using the following command:
  
  pip install pyinstaller
  
  BE CAREFUL: Due to the fact that the Python files work with pandas and json dependencies, this process can take a lot of time (from 5 to 20 minutes)

Graphical illustration of how the tree should look like: (It's not a problem if more files or folders are present; these are the minimum requirements).

PROCCESVIEW MESSAGES
|   ProcessViewMessages.xlsm
|
\---JSON Generator
    |   JsonGenerator.py

After running the program, the tree should look like this:

PROCCESVIEW MESSAGES
|   ProcessViewMessages.xlsm
|
+---JSON FILES
|       AVA.json
|       UFA.json
|       VILOFOSS.json
|
\---JSON Generator
        JsonGenerator.py
