#########IMPORTS from flask:
#   render_template:    get the templates
#   request:            gives access to incoming HTTP data (fields and files)
#   send_file:          return file as a response (XSLX here)
#   abort:              return HTTP error codes
#   secure_filename:    sanitizes and cleans filenames before use
from flask import Flask, render_template, request, send_file, abort
from werkzeug.utils import secure_filename
import tempfile, os, io

# import converter function
from converter import converter


#create Flask app
app = Flask(__name__) 


#########safety checks for file upload: type and length is checked
ALLOWED_EXTENSIONS = {"json"}
app.config['MAX_CONTENT_LENGTH'] = 1 * 1024 * 1024  # 16 MB limit (adjust as needed)


#########HELPER FUNCTION TO CHECK IF FILE ALIGNS WITH RESTRICTIONS
def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

#########ROUTE: / (home)
#shows the upload form with index.html
#only get is allowed here
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

#########ROUTE: /convert
#data form content gets sent here through action="/convert" in HTML
#only post is allowed here
@app.route('/convert', methods=['POST'])
def convert_route():
    # basic validations, all return a code 400 if not met
    if 'file' not in request.files: #ensure file is there
        return "No file part in request", 400
    file = request.files['file'] 
    if file.filename == '': #ensure filename exists
        return "No selected file", 400
    if not allowed_file(file.filename): #check if filetype is allowed through helper function
        return "Only .json files are allowed", 400

    # get metadata from form fields
    #.get(key, default) safely fetches values; returns '' if missing.
    # create a dict to pass to converter
    batchno = request.form.get('name', '')
    brewdate = request.form.get('date', '')  

    # Save uploaded JSON to a temporary file (converter expects filename)
    with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as tmp_in:
        file.save(tmp_in)             
        in_path = tmp_in.name

    # create a temporary output file path (we'll read bytes then delete files)
    out_fd, out_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(out_fd)  #
    try:
        # CALL YOUR converter here - adapt to your converter's API
        # Example: convert(input_path, output_path, metadata_dict)
        converter(batchno, brewdate,in_path, out_path)


        # read the generated xlsx into memory (so we can remove temp files right away)
        with open(out_path, 'rb') as f:
            xlsx_data = f.read()

    except Exception as e:
        # keep it simple: remove temp files and return error
        try:
            os.remove(in_path)
        except OSError:
            pass
        try:
            os.remove(out_path)
        except OSError:
            pass
        return f"Conversion failed: {e}", 500

    # cleanup temp files (we already read bytes)
    try:
        os.remove(in_path)
        os.remove(out_path)
    except OSError:
        pass

    # prepare download filename based on original name
    original_base = secure_filename(file.filename.rsplit('.', 1)[0]) #make filename safe
    download_name = f"{original_base}.xlsx"

    #return xlsx as download file
    return send_file(
        io.BytesIO(xlsx_data), #wrap into bytes
        as_attachment=True, #file is force-downloaded (not shown in browser)
        download_name=download_name, #MIME type for xlsx
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


#run the app with python app.py
if __name__ == '__main__':
    app.run(debug=True)