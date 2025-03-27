# JSON_Generator
This is a JSON generator programmed in Python. It automatically generates a Python file based on the information provided in the 'ProcessViewMessages.xlsm' Excel file.
---  
To convert the .py file into a .exe file, run the following command in your Windows terminal:

JSON Generator> pyinstaller --onefile --windowed jsonGenerator.py

The pyinstaller library is required. If you don’t have it installed, you can install it using the following command:

pip install pyinstaller

BE CAREFUL: Due to the fact that the Python files work with pandas and json dependencies, this process can take a lot of time (from 5 to 20 minutes)
