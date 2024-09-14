from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

# Path to the Excel file
file_path = './excel.xlsx'

# Load all sheet names
xls = pd.ExcelFile(file_path)


@app.route('/', methods=['GET', 'POST'])
def index():
    selected_sheet = request.form.get('sheet') or xls.sheet_names[0]  # Default to first sheet
    search_query = request.form.get('search', '').lower()  # Search query (if provided)

    try:
        # Load the selected sheet
        df = pd.read_excel(file_path, sheet_name=selected_sheet, engine='openpyxl')

        # If a search query is provided, filter the data
        if search_query:
            df = df[df.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)]

        # Convert DataFrame to HTML
        table_html = df.to_html(index=False)

        return render_template('index.html', table=table_html, sheets=xls.sheet_names, selected_sheet=selected_sheet,
                               search_query=search_query)

    except Exception as e:
        return f"An error occurred: {e}"


if __name__ == '__main__':
    app.run(debug=True)
