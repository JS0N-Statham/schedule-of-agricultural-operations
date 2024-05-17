import pandas as pd
import plotly.express as px
from flask import Flask, render_template
import os

# Read data from Excel
file_path = 'Операции ПМТП.xlsx'
sheet_data = pd.read_excel(file_path, sheet_name='Лист1')

# Process data
sheet_data.columns = ['Операции: ', 'Даты: ']

# Split Dates column into Start Date and Конец:
sheet_data[['Начало:', 'Конец:']] = sheet_data['Даты: '].str.split('-', expand=True)

# Clean data: Remove rows with invalid dates
sheet_data = sheet_data[sheet_data['Начало:'].notna() & sheet_data['Конец:'].notna()]

# Convert Start Date and Конец: to datetime format
sheet_data['Начало:'] = pd.to_datetime(sheet_data['Начало:'], format='%d.%m', errors='coerce')
sheet_data['Конец:'] = pd.to_datetime(sheet_data['Конец:'], format='%d.%m', errors='coerce')

# Drop rows where conversion to datetime failed
sheet_data = sheet_data.dropna(subset=['Начало:', 'Конец:'])

# Initialize Flask app
app = Flask(__name__)

@app.route('/')
def index():
    # Create Gantt chart
    fig = px.timeline(sheet_data, x_start='Начало:', x_end='Конец:', y='Операции: ', title='Агротехнические операции за год')
    fig.update_yaxes(categoryorder='total ascending')

    # Save the figure as HTML
    fig.write_html("templates/gantt_chart.html")

    return render_template('gantt_chart.html')

if __name__ == '__main__':
    app.run(debug=True)
