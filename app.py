from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename
import os
import logging
import uuid
import json
from create_bible_format import ChangeBibleFormat
from other_church import ChangeBibleFormat2
from create_song_form import ChangeLyrics

# Flask 애플리케이션 초기화
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', '3dce87f15d8d20bd3e3b80c6fe349bb997497a8899c4cd780076a7be653739d7')

# 폴더 설정
UPLOAD_FOLDER = 'uploads'
BIBLE_PPT_FOLDER = 'bible_ppts'
PPT_TEMPLATE_FOLDER = 'ppt_templates'
OUTPUT_FOLDER = 'output'
ALLOWED_PPT_EXTENSIONS = {'pptx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['BIBLE_PPT_FOLDER'] = BIBLE_PPT_FOLDER
app.config['PPT_TEMPLATE_FOLDER'] = PPT_TEMPLATE_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 최대 업로드 파일 크기 설정 (예: 100MB)

# 로그 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 업로드 가능한 파일인지 확인하는 함수
def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

# 번역본 목록 가져오기 함수
def get_bible_version_sets(folder_path):
    version_sets = []
    if os.path.isdir(folder_path):
        for folder in os.listdir(folder_path):
            full_path = os.path.join(folder_path, folder)
            if os.path.isdir(full_path):
                version_sets.append(folder)
                logging.info(f"Found version set: {folder}")
    else:
        logging.error(f"BIBLE_PPT_FOLDER does not exist: {folder_path}")
    logging.info(f"Total version sets found: {len(version_sets)}")
    return version_sets

def get_template_version_sets(folder_path):
    version_sets = []
    if os.path.isdir(folder_path):
        for ppt in os.listdir(folder_path):
            full_path = os.path.join(folder_path, ppt)
            if os.path.isfile(full_path):
                version_sets.append(ppt)
                logging.info(f"Found version set: {ppt}")
    else:
        logging.error(f"TEMPLATES does not exist: {folder_path}")

    sorted(version_sets,key=len)
    logging.info(f"Total version sets found: {len(version_sets)}")
    return version_sets

# Google Cloud 대신 로컬에 파일 저장하는 함수
def save_file_locally(source_file_path, destination_filename):
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    destination = os.path.join(app.config['OUTPUT_FOLDER'], destination_filename)
    os.rename(source_file_path, destination)
    return url_for('download_file', filename=destination_filename, _external=True)

# 루트 라우트: 파일 업로드 및 입력 폼
@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

# 말씀 PPT 생성 메뉴 라우트
@app.route('/create_bible_menu', methods=['GET'])
def create_bible_menu():
    return render_template('create_bible_menu.html')

# 말씀 PPT 생성 작업별 라우트
@app.route('/create_bible_format/<operation>', methods=['GET', 'POST'])
def create_bible_format_route(operation):
    allowed_operations = ['create_bible_format', 'other_church', 'YD_church']

    if operation not in allowed_operations:
        flash('알 수 없는 작업이 선택되었습니다.', 'danger')
        return redirect(url_for('create_bible_format_route', operation=operation))

    if request.method == 'POST':
        # 폼 데이터 가져오기
        text = request.form.get('text')
        ppt_title = request.form.get('ppt_title')
        bible_version_set = request.form.get('bible_version_set')
        ppt_version_set = request.form.get('ppt_version_set')

        if text == '':
            flash('텍스트를 입력해주세요.', 'warning')
            return redirect(url_for('create_bible_format_route', operation=operation))
        if not bible_version_set:
            flash('버전을 선택해주세요.', 'warning')
            return redirect(url_for('create_bible_format_route', operation=operation))
        if not ppt_title:
            flash('PPT 제목을 입력해주세요.', 'warning')
            return redirect(url_for('create_bible_format_route', operation=operation))

        try:
            selected_bible_version_path = os.path.join(app.config['BIBLE_PPT_FOLDER'], bible_version_set)
            if not os.path.isdir(selected_bible_version_path):
                flash('선택한 버전의 파일이 존재하지 않습니다.', 'danger')
                return redirect(url_for('create_bible_format_route', operation=operation))

            if operation == 'create_bible_format':
                if not ppt_version_set:
                    flash('PPT 템플릿을 선택해주세요.', 'warning')
                    return redirect(url_for('create_bible_format_route', operation=operation))
                selected_ppt_version_path = os.path.join(
                    app.config['PPT_TEMPLATE_FOLDER'], ppt_version_set
                )
                cbf = ChangeBibleFormat(
                    text=text,
                    bible_path=selected_bible_version_path,
                    ppt_template_path=selected_ppt_version_path,
                    output_path='/tmp',
                    ppt_title=ppt_title
                )
                cbf.create_ppt_file()

            elif operation == 'other_church':
                cbf2 = ChangeBibleFormat2(
                    text=text,
                    bible_path=selected_bible_version_path,
                    output_path='/tmp',
                    ppt_title=ppt_title
                )
                cbf2.create_ppt_file()

            ppt_filename = f"{ppt_title}.pptx"
            local_ppt_path = os.path.join('/tmp', ppt_filename)

            if os.path.exists(local_ppt_path):
                download_url = save_file_locally(local_ppt_path, ppt_filename)
                flash('PPT 파일이 성공적으로 생성되었습니다.', 'success')
                # 다운로드 URL을 파라미터로 함께 전달
                return redirect(url_for('show_result_page', download_url=download_url))
            else:
                flash('PPT 파일 생성에 실패했습니다.', 'danger')
                return redirect(url_for('create_bible_format_route', operation=operation))

        except Exception as e:
            logging.error(f"에러 발생: {e}")
            flash(f"에러 발생: {e}", 'danger')
            return redirect(url_for('create_bible_format_route', operation=operation))

    bible_version_sets = get_bible_version_sets(BIBLE_PPT_FOLDER)
    ppt_version_sets = get_template_version_sets(PPT_TEMPLATE_FOLDER)
    return render_template('create_bible_operations.html', operation=operation, bible_version_sets=bible_version_sets, ppt_version_sets=ppt_version_sets)

# ChangeLyrics 메인 메뉴 라우트
@app.route('/change_lyrics', methods=['GET'])
def change_lyrics_menu():
    return render_template('change_lyrics.html')

# ChangeLyrics 작업별 라우트
@app.route('/change_lyrics/<operation>', methods=['GET', 'POST'])
def change_lyrics_route(operation):
    allowed_operations = [
        'change_lyrics',
        'change_only_korean_lyrics',
        'change_only_english_lyrics',
        'create_lyrics',
        'create_only_korean_lyrics',
        'create_only_english_lyrics',
        'create_subtitle_lyrics',
        'create_seoul_form',
    ]

    if operation not in allowed_operations:
        flash('알 수 없는 작업이 선택되었습니다.', 'danger')
        return redirect(url_for('change_lyrics_menu'))

    if request.method == 'POST':
        # 폼 데이터 가져오기
        text = request.form.get('text')
        ppt_file = request.files.get('ppt_file')  # PPT 파일 업로드 필드 추가
        ppt_title = request.form.get('ppt_title')

        # 유효성 검사
        if text == '':
            flash('텍스트를 입력해주세요.', 'warning')
            return redirect(request.url)

        # 특정 작업에 따라 PPT 파일이 필요한지 확인
        operations_requiring_ppt = [
            'change_lyrics',
            'change_only_korean_lyrics',
            'change_only_english_lyrics'
        ]
        if operation in operations_requiring_ppt:
            if not ppt_file or ppt_file.filename == '':
                flash('기존 PPT 파일을 업로드해주세요.', 'warning')
                return redirect(request.url)
            if not allowed_file(ppt_file.filename, ALLOWED_PPT_EXTENSIONS):
                flash('허용되지 않은 파일 형식입니다. .pptx 파일만 업로드할 수 있습니다.', 'warning')
                return redirect(url_for('change_lyrics_route', operation=operation))

        if not ppt_title:
            flash('PPT 제목을 입력해주세요.', 'warning')
            return redirect(request.url)

        # 고유한 폴더 이름 생성 (UUID 사용)
        unique_id = str(uuid.uuid4())

        # 업로드된 PPT 파일 저장 (필요 시)
        if operation in operations_requiring_ppt:
            ppt_filename = secure_filename(ppt_file.filename)
            local_ppt_path = os.path.join('/tmp', f"{unique_id}_{ppt_filename}")
            ppt_file.save(local_ppt_path)
        else:
            local_ppt_path = None  # 새로 생성할 경우 기존 PPT가 없으므로 None

        try:
            # ChangeLyrics 인스턴스 초기화
            cl = ChangeLyrics(
                text=text,
                ppt_path=local_ppt_path,  # 기존 PPT 파일 경로 또는 None
                save_path='/tmp',
                ppt_title=ppt_title
            )
            # 선택된 작업 메서드 호출
            getattr(cl, operation)()

            # 생성된 PPT 파일 경로
            ppt_output_filename = f"{ppt_title}.pptx"
            local_ppt_output_path = os.path.join('/tmp', ppt_output_filename)

            if os.path.exists(local_ppt_output_path):
                download_url = save_file_locally(local_ppt_output_path, ppt_output_filename)
                flash('PPT 파일이 성공적으로 생성되었습니다.', 'success')
                # 다운로드 URL을 파라미터로 함께 전달
                return redirect(url_for('show_result_page', download_url=download_url))
            else:
                flash('PPT 파일 생성에 실패했습니다.', 'danger')
                return redirect(url_for('change_lyrics_route', operation=operation))


        except Exception as e:
            logging.error(f"에러 발생: {e}")
            flash(f"에러 발생: {e}", 'danger')
            return redirect(url_for('change_lyrics_route', operation=operation))

    return render_template('change_lyrics_operations.html', operation=operation)

# PPT 파일 다운로드 라우트
@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(directory=app.config['OUTPUT_FOLDER'], path=filename, as_attachment=True)

@app.route('/show_result_page')
def show_result_page():
    # 쿼리 파라미터로 넘어온 download_url 받기
    download_url = request.args.get('download_url')
    return render_template('result_page.html', download_url=download_url)


if __name__ == '__main__':
    # 폴더가 없으면 생성
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(BIBLE_PPT_FOLDER, exist_ok=True)
    os.makedirs(PPT_TEMPLATE_FOLDER, exist_ok=True)
    os.makedirs('/tmp', exist_ok=True)
    app.run(host='0.0.0.0', port=5050, debug=True)
