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

    if title[0] == 'И':
        if title[-1] == '2':
            filename = 'IIT_1-kurs_22_23_vesna_TANDEM_29.03.2023.xlsx'
        if title[-1] == '1':
            filename = 'IIT_2-kurs_22_23_vesna_27.04.2023.xlsx'
        if title[-1] == '0':
            filename = 'IIT_3-kurs_22_23_vesna_02.05.2023.xlsx'
    elif title[0] == 'К':
        if title[-1] == '2':
            filename = 'III_1-kurs_22_23_vesna_10.05.2023.xlsx'
        if title[-1] == '1':
            filename = 'III_2-kurs_22_23_vesna_10.05.2023.xlsx'
        if title[-1] == '0':
            filename = 'III_3-kurs_22_23_vesna_28.04.2023.xlsx'
        if title[-1] == '9':
            filename = 'III_4-kurs_22_23_vesna_27.04.2023.xlsx'
        if title[-1] == '8':
            filename = 'III_5-kurs_22_23_vesna_03.05.2023.xlsx'
    elif title[0] == 'Б':
        if title[-1] == '2':
            filename = 'IKTST_1_k_vesna_22_23.xlsx'
        if title[-1] == '1':
            filename = 'IKTST_2_k_vesna_22_23.xlsx'
        if title[-1] == '0':
            filename = 'IKTST_3_k_vesna_22_23.xlsx'
        if title[-1] == '9':
            filename = 'IKTST_4_k_vesna_22_23.xlsx'
        if title[-1] == '8':
            filename = 'IKTST_5_k_vesna_22_23.xlsx'
    elif title[0] == 'Т':
        if title[-1] == '2':
            filename = 'IPTIP_1-kurs_22_23_vesna_10.04.2023.xlsx'
        if title[-1] == '1':
            filename = 'IPTIP_2-kurs_22_23_vesna_10.04.2023.xlsx'
        if title[-1] == '0':
            filename = 'IPTIP_3-kurs_22_23_vesna_19.04.2023.xlsx'
        if title[-1] == '9':
            filename = 'IPTIP_4-kurs_22_23_vesna_10.04.2023.xlsx'
    elif title[0] == 'Р':
        if title[-1] == '2':
            filename = 'IRI_1-kurs_22_23_vesna_TANDEM_23.03.2023.xlsx'
        if title[-1] == '1':
            filename = 'IRI_2-kurs_22_23_vesna_TANDEM_22.03.2023.xlsx'
        if title[-1] == '0':
            filename = 'IRI_3-kurs_22_23_vesna_10.04.2023.xlsx'
        if title[-1] == '9':
            filename = 'IRI_4-kurs_22_23_vesna_TANDEM_22.03.2023.xlsx'
        if title[-1] == '8':
            filename = 'IRI_5-kurs_22_23_vesna_TANDEM_22.03.2023.xlsx'
    elif title[0] == 'У':
        if title[-1] == '2':
            filename = 'ITU_1-kurs_22_23_vesna_03.05.2023.xlsx'
        if title[-1] == '1':
            filename = 'ITU_2-kurs_22_23_vesna_05.05.2023.xlsx'
        if title[-1] == '0':
            filename = 'ITU_3-kurs_22_23_vesna_02.05.2023.xlsx'

    wb = op.load_workbook(filename, data_only=True)

    sheet = wb.active
    max_column = sheet.max_column
    arr = []
    arr_prep = []
    arr_surname = []
    group = title

    for i in range(2, max_column+1):
        search_group = sheet.cell(row=2, column=i).value
        numberOfColumn = i
        if not search_group:
            continue

        if (search_group == group):
            for j in range(4, 88):
                prep = sheet.cell(row=j, column=i+1).value
                arr.append(sheet.cell(row=j, column=i).value)
                surname = sheet.cell(row=j, column=i+2).value

                if prep != "":
                    arr_prep.append(" || " + " (" + (prep) + ") " + " || ")
                else:
                    arr_prep.append("")

                if surname != "":
                    arr_surname.append(" (" + (surname) + ") " + " || ")
                else:
                    arr_surname.append("")

    return render_template('start.html', title=title, arr=arr, arr_prep=arr_prep, arr_surname=arr_surname)


if __name__ == '__main__':
    app.run(debug=True)












