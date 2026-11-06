import os, sys, random, string, subprocess # Standard Python Libraries
from docxtpl import DocxTemplate, InlineImage  # pip install docxtpl
from docx.shared import Cm, Inches, Mm, Emu  # pip install python-docx
import pandas as pd

def generate_activation_code(length):
    letters_and_digits = string.ascii_letters + string.digits
    return ''.join(random.choice(letters_and_digits) for i in range(length))

activation_code = generate_activation_code(32) 
print(activation_code)

random_number = random.randint(1, 100)
print(random_number)

# Read an Excel file
df = pd.read_excel('SARPAI.xlsx')
df = pd.read_excel('0001-1QQ4.xlsb')
df = pd.read_excel('0010-jcyQ.xlsb')
df = pd.read_excel('0011-6myO.xlsm')
# Display the contents of the Excel file
print(df) 

# the path to your executable
subprocess.call(['SARPAI.exe'])

class AI:
    def __init__(self, name):
        self.name = name

    def greet(self):
        return f"Hello, I am {self.name}. How can I assist you today?"

    def process_input(self, user_input):
        # Add your AI logic here to process user input
        return "Processing input: " + user_input

# Creating an instance of the AI class
my_ai = AI("SARPAI")

# Example usage
print(my_ai.greet())
print(my_ai.process_input("How does AI work?"))




