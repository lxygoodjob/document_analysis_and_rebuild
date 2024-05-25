import os
from flask import Flask, request, make_response, send_from_directory  
import json
import requests
import uuid
from utils.docx_file_process import docx_unzip_and_parser, get_docx_text_per_para, parser_docx_trans_result
from utils.pptx_file_process import pptx_unzip_and_parser, get_pptx_text_per_para, parser_pptx_trans_result
from utils.fileconvert import convert_doc_to_docx, convert_xls_to_xlsx, convert_ppt_to_pptx
from utils.xlsx_file_process import xlsx_parser, parser_xlsx_trans_result
from utils.text_file_process import split_func

import base64


sharepath = '/opt/xxx/DT/TempShareFolder'


########## flask func ##########
app = Flask(__name__)

class JsonEncoder(json.JSONEncoder):
    """Convert numpy classes to JSON serializable objects."""

    def default(self, obj):
        if isinstance(obj, (np.integer, np.floating, np.bool_)):
            return obj.item()
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        else:
            return super(JsonEncoder, self).default(obj)

def json_dumps(data):
    return json.dumps(data, ensure_ascii=False, cls=JsonEncoder)

def wrap_resp(res, status=400, headers={'Content-Type': 'application/json'}):
    resp = make_response(json_dumps(res), status)
    resp.headers = headers
    return resp

@app.route('/doc_parser', methods = ['POST'])
def doc_parser():
    inf = request.files.get('input_file')
    uid = str(uuid.uuid4())
    save_folder = os.path.join(sharepath, uid)
    out_folder = os.path.join(save_folder, 'output')
    os.makedirs(save_folder, exist_ok=True)
    input_file = os.path.join(save_folder, inf.filename)
    inf.save(input_file)
    try:
        if inf.filename.lower().endswith(('.doc', '.docx')):
            file_type = 'docx'
            if inf.filename.lower().endswith('.doc'):
                input_file = convert_doc_to_docx(input_file, save_folder)
            unzip_file_dict, unzip_path = docx_unzip_and_parser(input_file)
            fileuuid, result = get_docx_text_per_para(unzip_file_dict, unzip_path, os.path.join(save_folder, uid +'.json'))
        elif inf.filename.lower().endswith(('.xls', '.xlsx')):
            file_type = 'xlsx'
            if inf.filename.lower().endswith('.xls'):
                input_file = convert_xls_to_xlsx(input_file, save_folder)
            fileuuid, result = xlsx_parser(input_file, os.path.join(save_folder, uid +'.json'))
        elif inf.filename.lower().endswith(('.txt')):
            file_type = 'txt'
            result = split_func(input_file)
        elif inf.filename.lower().endswith(('.ppt', '.pptx')):
            file_type = 'pptx'
            if inf.filename.lower().endswith('.ppt'):
                input_file = convert_ppt_to_pptx(input_file, save_folder)
            unzip_file_dict, unzip_path = pptx_unzip_and_parser(input_file)
            fileuuid, result = get_pptx_text_per_para(unzip_file_dict, unzip_path, os.path.join(save_folder, uid +'.json'))
        else:
            result = None
    except:
        result = None

    if result is not None:
        res = {
            'msg': 'success',
            'uid': uid,
            'file_type': file_type,
            'data': result,
            'code': 0,
            }
        return wrap_resp(res, 200)
    else:
        res = {
            'msg': 'fail',
            'code': -1,
            }
        return wrap_resp(res)


@app.route('/rebuild_doc', methods = ['POST'])
def rebuild_doc():
    uid = request.get_json()['uid']
    file_type = request.get_json()['file_type']
    trans_result = request.get_json()['trans_result']
    save_folder = os.path.join(sharepath, uid)
    out_folder = os.path.join(save_folder, 'output')
    fileuuid = os.path.join(save_folder, uid +'.json')
    save_name = str(uuid.uuid4())
    try:
        if file_type == 'docx':
            docxname = os.path.join(save_folder, save_name + '.docx')
            docxname = parser_docx_trans_result(trans_result, fileuuid, docxname)
            output_file = convert_doc_to_docx(docxname, out_folder)
            os.remove(str(docxname))
        elif file_type == 'xlsx':
            xlsxname = os.path.join(save_folder, save_name + '.xlsx')
            output_file = parser_xlsx_trans_result(trans_result, fileuuid, xlsxname)
        elif file_type == 'pptx':
            pptxname = os.path.join(save_folder, save_name + '.pptx')
            pptxname = parser_pptx_trans_result(trans_result, fileuuid, pptxname)
            output_file = convert_ppt_to_pptx(pptxname, out_folder) 
            os.remove(str(pptxname))
        else:
            output_file = None
    except:
            output_file = None
    
    if output_file is not None:
        try:
            with open(output_file, 'rb') as f:
                base64_str = base64.b64encode(f.read()).decode(encoding='utf-8')
            res = {
                'msg': 'success',
                'data': base64_str,
                'code': 0,
                }
            return wrap_resp(res, 200)
        except:
            res = {
                'msg': 'fail',
                'code': -1,
                }
            return wrap_resp(res)
    else:
        res = {
            'msg': 'fail',
            'code': -1,
            }
        return wrap_resp(res)


if __name__ == "__main__":
    app.run('0.0.0.0', 9901, debug=False, threaded=True)




