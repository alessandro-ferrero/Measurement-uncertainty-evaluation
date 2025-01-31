"""
This script requires the following libraries to be installed:

- numpy
- matplotlib
- scipy
- PySimpleGUI

You can install these libraries using pip:

pip install numpy matplotlib scipy PySimpleGUI
"""

import numpy as np
import matplotlib.pyplot as plt
import matplotlib 
from scipy.stats import norm
import PySimpleGUI as sg

def read_data_from_file(file_path):
    matrix = []
    with open(file_path, 'r') as file:
        for line in file:
            # Remove exceeding blanks and newlines, then splits the line by ';' separator
            dati = line.strip().split(';')
            # Converts the data to float
            dati = [float(dato) for dato in dati]
            # Appends the data to the matrix
            matrix.append(dati)
    return matrix

def open_file():
    """Opens a dialog wondow to select a text file."""
    file_path = sg.popup_get_file(
        "Select a text file",
        file_types=(
            ("Text files", "*.txt"),
            ("All files", "*.*")
        )
    )
    return file_path

def get_numeric_input(prompt):
    """Requests a numerical value to the user."""
    while True:
        try:
            return float(input(prompt))
        except ValueError:
            print("Please, enter a valid numerical value.")

"""
A dialog window will open to select a text file containing the samples of the cumulative
probability function of the values that can be reasonably attributed to the measurnad. 
The text file must contain a matrix of values separated by semicolons. The first column
must contain the x values and the second column must contain the y values.
"""

file_path=open_file()

# Reads the data from the file
cpf=read_data_from_file(file_path)

# Extracts the X and Y values from the matrix
X = [row[0] for row in cpf]
Y = [row[1] for row in cpf]
Xmin=min(X)
Xmax=max(X)
Ymin=0
Ymax=1

# Requests the limit value
limit_value = get_numeric_input('Enter the limit value: ')

# Computes the probability that the measured value is below the limit value
j=0
while X[j]<=limit_value:
    j+=1
    if j>=len(X):
        print('The limit value is greater than the maximum value in the cumulative probability function.')
        exit()
Prob_below_limit=Y[j]
Prob_below_limit_formatted = f'{(Prob_below_limit*100):.1f}'
print(f'The probability that the measured value is below {limit_value} is: {Prob_below_limit_formatted}%')

# Generates the plot
plt.plot(X, Y, color='blue', linewidth=2)
plt.plot([limit_value, limit_value], [Ymin, Prob_below_limit], color='red', linestyle='dashed', linewidth=1)
plt.plot([Xmin, limit_value], [Prob_below_limit, Prob_below_limit], color='red', linestyle='dashed', linewidth=1)
plt.axis([Xmin, Xmax, Ymin, Ymax])
plt.xlabel('Values that can be attributed to the measurand')
plt.ylabel('Probability')
plt.title('Cumulative probability function')
plt.grid(True)

# Shows the plot
plt.show()