from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import subprocess
import plotly.express as px

app = Flask(__name__)

def open_excel_file():
    try:
        subprocess.Popen(['start', 'datas.xlsx'], shell=True)
    except Exception as e:
        print(f"Error opening Excel file: {e}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/add', methods=['POST'])
def add_data():
    name = request.form.get('name')
    usn = request.form.get('usn')
    branch = request.form.get('branch')
    sec = request.form.get('sec')
    hobbies = request.form.get('hobbies')
    marks = request.form.get('marks')  # New line for marks

    new_data = pd.DataFrame({'Name': [name], 'USN': [usn], 'Branch': [branch], 'Sec': [sec], 'Hobbies': [hobbies], 'Marks': [marks]})

    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
        updated_data = pd.concat([existing_data, new_data], ignore_index=True)
    except FileNotFoundError:
        updated_data = new_data

    updated_data.to_excel('datas.xlsx', index=False, engine='openpyxl')
    open_excel_file()
    return render_template('index.html', table=updated_data.to_html(classes='table table-bordered table-striped', index=False))

@app.route('/delete', methods=['POST'])
def delete_data():
    name = request.form.get('delete_name')

    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
        updated_data = existing_data[existing_data['Name'] != name]
    except FileNotFoundError:
        updated_data = pd.DataFrame()

    updated_data.to_excel('datas.xlsx', index=False, engine='openpyxl')
    open_excel_file()
    return render_template('index.html', table=updated_data.to_html(classes='table table-bordered table-striped', index=False))
# ...

@app.route('/update', methods=['POST'])
def update_data():
    old_name = request.form.get('old_name')
    new_name = request.form.get('new_name')
    new_usn = request.form.get('new_usn')
    new_branch = request.form.get('new_branch')
    new_sec = request.form.get('new_sec')
    new_hobbies = request.form.get('new_hobbies')
    new_marks = request.form.get('new_marks')  # New line for marks

    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
        existing_data.loc[existing_data['Name'] == old_name, 'Name'] = new_name
        existing_data.loc[existing_data['Name'] == old_name, 'USN'] = new_usn
        existing_data.loc[existing_data['Name'] == old_name, 'Branch'] = new_branch
        existing_data.loc[existing_data['Name'] == old_name, 'Sec'] = new_sec
        existing_data.loc[existing_data['Name'] == old_name, 'Hobbies'] = new_hobbies
        existing_data.loc[existing_data['Name'] == old_name, 'Marks'] = new_marks  # Update the Marks column
    except FileNotFoundError:
        existing_data = pd.DataFrame()

    existing_data.to_excel('datas.xlsx', index=False, engine='openpyxl')
    open_excel_file()
    return render_template('index.html', table=existing_data.to_html(classes='table table-bordered table-striped', index=False))

# ...


@app.route('/search', methods=['POST'])
def search_data():
    search_name = request.form.get('search_name')

    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
        search_result = existing_data[existing_data['Name'].str.contains(search_name, case=False, na=False)]
    except FileNotFoundError:
        search_result = pd.DataFrame()

    return render_template('index.html', table=search_result.to_html(classes='table table-bordered table-striped', index=False))
@app.route('/view-all')
def view_all_data():
    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
    except FileNotFoundError:
        existing_data = pd.DataFrame()

    return render_template('index.html', table=existing_data.to_html(classes='table table-bordered table-striped', index=False))
@app.route('/clear-all')
def clear_all_data():
    try:
        existing_data = pd.DataFrame()
        existing_data.to_excel('datas.xlsx', index=False, engine='openpyxl')
        open_excel_file()
    except Exception as e:
        print(f"Error clearing all data: {e}")

    return render_template('index.html', table=existing_data.to_html(classes='table table-bordered table-striped', index=False))
@app.route('/sort', methods=['POST'])
def sort_data():
    sort_column = request.form.get('sort_column')
    sort_order = request.form.get('sort_order')

    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
        existing_data = existing_data.sort_values(by=sort_column, ascending=(sort_order == 'asc'))
    except FileNotFoundError:
        existing_data = pd.DataFrame()

    existing_data.to_excel('datas.xlsx', index=False, engine='openpyxl')
    open_excel_file()
    return render_template('index.html', table=existing_data.to_html(classes='table table-bordered table-striped', index=False))
@app.route('/export', methods=['POST'])
def export_data():
    export_format = request.form.get('export_format')

    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')

        if export_format == 'csv':
            existing_data.to_csv('datas_exported.csv', index=False)
        elif export_format == 'xlsx':
            existing_data.to_excel('datas_exported.xlsx', index=False, engine='openpyxl')
    except FileNotFoundError:
        print("No data to export.")

    return render_template('index.html', message=f'Data exported to {export_format.upper()} file.')

import plotly.express as px
# Change the endpoint for statistics route
@app.route('/show-statistics')
def show_statistics():
    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
        statistics = existing_data.describe()
    except FileNotFoundError:
        statistics = pd.DataFrame()

    return render_template('statistics.html', table=statistics.to_html(classes='table table-bordered table-striped', index=True))

# Route for handling visualization
# Route for handling visualization
@app.route('/visualize', methods=['POST'])
def visualize_data():
    chart_column = request.form.get('chart_column')
    chart_type = request.form.get('chart_type')

    try:
        existing_data = pd.read_excel('datas.xlsx', engine='openpyxl')
        fig = px.bar(existing_data, x='Name', y='Marks', title=f'{chart_type} Chart for {chart_column} and Marks')
    except FileNotFoundError:
        fig = None

    return render_template('visualization.html', plot=fig.to_html(full_html=False))

if __name__ == '__main__':
    app.run(debug=True)

