from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import openpyxl as op


app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        title = request.form.get('title')
        return redirect(url_for('start', title=title))
    return render_template('index.html')


@app.route('/start')
def start():
    title = request.args.get('title', None)
    filename = 'IIT_1-kurs_22_23_vesna_TANDEM_29.03.2023.xlsx'

    wb = op.load_workbook(filename, data_only=True)

    sheet = wb.active
    max_column = sheet.max_column
    arr = []
    group = title
    for i in range(2, max_column+1):
        search_group = sheet.cell(row=2, column=i).value
        numberOfColumn = i
        if not search_group:
            continue

        if (search_group == group):
            for j in range(4, 88):
                arr.append(sheet.cell(row=j, column=i).value)


    return render_template('start.html', title=title, arr=arr)


if __name__ == '__main__':
    app.run(debug=True)












