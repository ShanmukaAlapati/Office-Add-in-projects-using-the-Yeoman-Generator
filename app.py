from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import json
import os
from datetime import datetime
import uuid
from typing import Any, Dict, List, cast

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app, origins="*")

NOTES_FILE = 'saved_notes.json'


def load_notes() -> List[Dict[str, Any]]:
    path = os.path.join(os.getcwd(), NOTES_FILE)
    if os.path.exists(path):
        with open(path, 'r') as f:
            try:
                data = json.load(f)
            except Exception:
                return []
            if isinstance(data, list):
                return cast(List[Dict[str, Any]], data)
    return []


def save_notes(notes: List[Dict[str, Any]]) -> None:
    path = os.path.join(os.getcwd(), NOTES_FILE)
    with open(path, 'w') as f:
        json.dump(notes, f, indent=2)

# Serve ALL HTML/JS/CSS/assets


@app.route('/', defaults={'path': 'taskpane.html'})
@app.route('/<path:path>')
def catch_all(path: str):
    return send_from_directory('.', path)

# API: Save note


@app.route('/api/save-note', methods=['POST'])
def save_note():
    data: Dict[str, Any] = request.get_json(silent=True) or {}
    notes: List[Dict[str, Any]] = load_notes()
    note: Dict[str, Any] = {
        'id': str(uuid.uuid4()),
        'note': data.get('note', ''),
        'emailSubject': data.get('emailSubject', 'Unknown'),
        'timestamp': datetime.now().isoformat(),
        'emailId': data.get('emailId', '')
    }
    notes.append(note)
    save_notes(notes)
    return jsonify({'success': True, 'noteId': note['id']})

# API: Get notes


@app.route('/api/notes')
def get_notes():
    return jsonify(load_notes())

# API: Delete note


@app.route('/api/delete-note/<note_id>', methods=['DELETE'])
def delete_note(note_id: str):
    notes: List[Dict[str, Any]] = load_notes()
    notes = [n for n in notes if n.get('id') != note_id]
    save_notes(notes)
    return jsonify({'success': True})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)


# from flask import Flask, request, jsonify, send_from_directory
# from flask_cors import CORS
# import json
# import os
# from datetime import datetime
# import uuid

# app = Flask(__name__, static_folder='.', static_url_path='')
# CORS(app, origins="*")  # Allows Outlook

# NOTES_FILE = 'saved_notes.json'

# def load_notes():
#     path = os.path.join(os.getcwd(), NOTES_FILE)
#     if os.path.exists(path):
#         with open(path, 'r') as f:
#             return json.load(f)
#     return []

# def save_notes(notes):
#     path = os.path.join(os.getcwd(), NOTES_FILE)
#     with open(path, 'w') as f:
#         json.dump(notes, f, indent=2)

# # Serve ALL your HTML/JS/CSS/assets
# @app.route('/', defaults={'path': 'taskpane.html'})
# @app.route('/<path:path>')
# def catch_all(path):
#     return send_from_directory('.', path)

# # API: Save note from Outlook
# @app.route('/api/save-note', methods=['POST'])
# def save_note():
#     data = request.json or {}
#     notes = load_notes()
#     note = {
#         'id': str(uuid.uuid4()),
#         'note': data.get('note', ''),
#         'emailSubject': data.get('emailSubject', 'Unknown'),
#         'timestamp': datetime.now().isoformat(),
#         'emailId': data.get('emailId', '')
#     }
#     notes.append(note)
#     save_notes(notes)
#     return jsonify({'success': True, 'noteId': note['id']})

# # API: Get notes
# @app.route('/api/notes')
# def get_notes():
#     return jsonify(load_notes())

# # API: Delete note
# @app.route('/api/delete-note/<note_id>', methods=['DELETE'])
# def delete_note(note_id):
#     notes = load_notes()
#     notes = [n for n in notes if n['id'] != note_id]
#     save_notes(notes)
#     return jsonify({'success': True})

# if __name__ == '__main__':
#     port = int(os.environ.get('PORT', 5000))
#     app.run(debug=True, host='0.0.0.0', port=port)
