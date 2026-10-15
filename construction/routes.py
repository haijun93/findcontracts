from flask import Blueprint, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge
from .forms import FORM_LABELS
from .core import DocumentError, analyze, extract_document, generate_package

blueprint = Blueprint('construction', __name__, template_folder='templates', static_folder='static', static_url_path='/construction/static')


@blueprint.get('/construction')
def workspace():
    return render_template('construction.html', form_labels=FORM_LABELS)


@blueprint.post('/api/construction/analyze')
def analyze_files():
    request.max_content_length = 42 * 1024 * 1024
    try:
        documents = []
        for key, role in [('quote', '견적서'), ('request', '의뢰서')]:
            files = request.files.getlist(key)
            if len(files) != 1 or not files[0].filename:
                raise DocumentError('견적서와 의뢰서를 각각 하나씩 선택하세요.')
            document = files[0]
            data = document.stream.read(20 * 1024 * 1024 + 1)
            documents.append(extract_document(data, document.filename, role))
        result = analyze(*documents)
        return jsonify(result)
    except RequestEntityTooLarge:
        return jsonify(error='파일 두 개의 합계는 40MB 이하로 입력하세요.'), 413
    except DocumentError as error:
        return jsonify(error=str(error)), 422


@blueprint.post('/api/construction/generate')
def generate_documents():
    request.max_content_length = 2 * 1024 * 1024
    try:
        payload = request.get_json(silent=True)
        if not isinstance(payload, dict) or payload.get('reviewed') is not True:
            raise DocumentError('추출 내용과 원본을 검토한 뒤 검토 확인을 선택하세요.')
        provenance = payload.get('provenance', {})
        if not isinstance(provenance, dict):
            raise DocumentError('추출 근거의 형식을 확인하세요.')
        output = generate_package(payload.get('facts'), provenance)
        response = send_file(output, mimetype='application/zip', as_attachment=True,
                             download_name='공사문서_초안.zip')
        response.headers['Cache-Control'] = 'no-store'
        return response
    except RequestEntityTooLarge:
        return jsonify(error='검토 내용이 너무 큽니다.'), 413
    except DocumentError as error:
        return jsonify(error=str(error)), 422
