from flask import Blueprint, send_file

download_controller = Blueprint('download_controller', __name__)

@download_controller.route('/download')
def download_excel():
    excel_file_path = 'Stages_DataSet.xlsx'  # Provide the correct path to your Excel file
    return send_file(excel_file_path, as_attachment=True)
