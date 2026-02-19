import io
import os
import shutil
import tempfile
import threading
import uuid
import zipfile

import fitz  # PyMuPDF
import ghostscript
import openpyxl
from flask import Flask, jsonify, render_template_string, request, send_file
from pdf2docx import Converter
from pdf2docx.converter import ConversionException
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["OUTPUT_FOLDER"] = "outputs"

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf", "docx", "txt", "xlsx", "jpg", "jpeg", "png"}

tasks: dict = {}
tasks_lock = threading.Lock()


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


class SavedFile:
    def __init__(self, path: str):
        self.path = path
        self.filename = os.path.basename(path)

    def save(self, dest: str):
        if os.path.abspath(dest) != os.path.abspath(self.path):
            shutil.copy(self.path, dest)


def set_progress(task_id: str, progress: int, message: str = "", status: str = "processing"):
    with tasks_lock:
        if task_id in tasks:
            tasks[task_id]["progress"] = progress
            tasks[task_id]["message"]  = message
            tasks[task_id]["status"]   = status

HTML_TEMPLATE = (
'\n<!DOCTYPE html>\n<html lang="pt-BR">\n<head>\n    <meta charset="UTF-8">\n    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n    <title>LocalPDF - Ferramentas PDF Corporativas</title>\n    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">\n    <style>\n        * { \n            margin: 0; \n            padding: 0; \n            box-sizing: border-box; \n        }\n        \n        :root {\n            --primary: #0066CC;\n            --primary-dark: #004C99;\n            --primary-light: #3385D6;\n            --secondary: #00A896;\n            --secondary-dark: #008778;\n            --accent: #0099FF;\n            --dark: #1A1A2E;\n            --gray: #64748B;\n            --light-gray: #F1F5F9;\n            --white: #FFFFFF;\n            --success: #10B981;\n            --error: #EF4444;\n            --warning: #F59E0B;\n        }\n        \n        body { \n            font-family: \'Inter\', -apple-system, BlinkMacSystemFont, \'Segoe UI\', sans-serif; \n            background: linear-gradient(135deg, #0066CC 0%, #00A896 100%);\n            min-height: 100vh; \n            position: relative;\n            overflow-x: hidden;\n        }\n        \n        /* Animated background */\n        body::before {\n            content: \'\';\n            position: fixed;\n            top: 0;\n            left: 0;\n            width: 100%;\n            height: 100%;\n            background: \n                radial-gradient(circle at 20% 50%, rgba(0, 168, 150, 0.3) 0%, transparent 50%),\n                radial-gradient(circle at 80% 80%, rgba(0, 102, 204, 0.3) 0%, transparent 50%),\n                radial-gradient(circle at 40% 20%, rgba(0, 153, 255, 0.2) 0%, transparent 50%);\n            animation: gradientShift 15s ease infinite;\n            pointer-events: none;\n            z-index: 0;\n        }\n        \n        @keyframes gradientShift {\n            0%, 100% { opacity: 1; transform: scale(1); }\n            50% { opacity: 0.8; transform: scale(1.1); }\n        }\n        \n        .container { \n            max-width: 1400px; \n            margin: 0 auto; \n            padding: 20px; \n            position: relative;\n            z-index: 1;\n        }\n        \n        /* Header Styles */\n        .header { \n            text-align: center; \n            color: white; \n            margin-bottom: 50px; \n            padding: 60px 20px 40px;\n            animation: fadeInDown 0.8s ease;\n        }\n        \n        @keyframes fadeInDown {\n            from { opacity: 0; transform: translateY(-30px); }\n            to { opacity: 1; transform: translateY(0); }\n        }\n        \n        .header .badge {\n            display: inline-block;\n            background: rgba(255, 255, 255, 0.2);\n            backdrop-filter: blur(20px);\n            -webkit-backdrop-filter: blur(20px);\n            padding: 10px 24px;\n            border-radius: 50px;\n            margin-bottom: 24px;\n            font-size: 0.9rem;\n            font-weight: 600;\n            border: 1px solid rgba(255, 255, 255, 0.3);\n            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);\n            animation: pulse 2s ease-in-out infinite;\n        }\n        \n        @keyframes pulse {\n            0%, 100% { transform: scale(1); }\n            50% { transform: scale(1.05); }\n        }\n        \n        .header h1 { \n            font-size: 4em; \n            margin-bottom: 16px; \n            font-weight: 900;\n            letter-spacing: -2px;\n            text-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);\n            background: linear-gradient(to right, #FFFFFF, #E0F2FE);\n            -webkit-background-clip: text;\n            -webkit-text-fill-color: transparent;\n            background-clip: text;\n        }\n        \n        .header p { \n            font-size: 1.3em; \n            opacity: 0.95; \n            font-weight: 400;\n            text-shadow: 0 2px 10px rgba(0, 0, 0, 0.2);\n            max-width: 600px;\n            margin: 0 auto;\n            line-height: 1.6;\n        }\n        \n        /* Tools Grid */\n        .tools-grid { \n            display: grid; \n            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); \n            gap: 24px; \n            margin-bottom: 50px;\n            animation: fadeInUp 0.8s ease;\n        }\n        \n        @keyframes fadeInUp {\n            from { opacity: 0; transform: translateY(30px); }\n            to { opacity: 1; transform: translateY(0); }\n        }\n        \n        /* Tool Card */\n        .tool-card { \n            background: rgba(255, 255, 255, 0.95);\n            backdrop-filter: blur(20px);\n            -webkit-backdrop-filter: blur(20px);\n            border-radius: 20px; \n            padding: 36px; \n            text-align: center; \n            box-shadow: \n                0 10px 40px rgba(0, 0, 0, 0.1),\n                0 2px 8px rgba(0, 0, 0, 0.06);\n            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1); \n            cursor: pointer; \n            border: 1px solid rgba(255, 255, 255, 0.8);\n            position: relative;\n            overflow: hidden;\n        }\n        \n        .tool-card::before {\n            content: \'\';\n            position: absolute;\n            top: 0;\n            left: -100%;\n            width: 100%;\n            height: 100%;\n            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);\n            transition: left 0.5s;\n        }\n        \n        .tool-card:hover::before {\n            left: 100%;\n        }\n        \n        .tool-card:hover { \n            transform: translateY(-12px) scale(1.02); \n            box-shadow: \n                0 20px 60px rgba(0, 102, 204, 0.3),\n                0 8px 16px rgba(0, 0, 0, 0.1);\n            border-color: var(--primary);\n        }\n        \n        .tool-card .icon {\n            font-size: 3.5em;\n            margin-bottom: 20px;\n            display: inline-block;\n            transition: transform 0.3s ease;\n            filter: drop-shadow(0 4px 8px rgba(0, 0, 0, 0.1));\n        }\n        \n        .tool-card:hover .icon {\n            transform: scale(1.1) rotate(5deg);\n        }\n        \n        .tool-card h3 { \n            color: var(--primary); \n            margin-bottom: 16px; \n            font-size: 1.5em; \n            font-weight: 700;\n            letter-spacing: -0.5px;\n        }\n        \n        .tool-card p { \n            color: var(--gray); \n            line-height: 1.7;\n            font-size: 1.05em;\n            font-weight: 400;\n        }\n        \n        /* Upload Area */\n        .upload-area { \n            border: 3px dashed #CBD5E1; \n            border-radius: 20px; \n            padding: 60px 40px; \n            text-align: center; \n            background: linear-gradient(135deg, #F8FAFC 0%, #F1F5F9 100%);\n            margin: 30px 0; \n            transition: all 0.3s ease; \n            cursor: pointer;\n            position: relative;\n        }\n        \n        .upload-area::before {\n            content: \'📤\';\n            position: absolute;\n            top: 50%;\n            left: 50%;\n            transform: translate(-50%, -50%);\n            font-size: 8em;\n            opacity: 0.05;\n            pointer-events: none;\n        }\n        \n        .upload-area:hover { \n            border-color: var(--primary); \n            background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);\n            border-width: 3px;\n            transform: scale(1.01);\n        }\n        \n        .upload-area.dragover { \n            border-color: var(--secondary); \n            background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%); \n            border-width: 4px;\n            transform: scale(1.02);\n        }\n        \n        .upload-area .upload-icon {\n            font-size: 4em;\n            margin-bottom: 20px;\n            display: block;\n            animation: bounce 2s ease-in-out infinite;\n        }\n        \n        @keyframes bounce {\n            0%, 100% { transform: translateY(0); }\n            50% { transform: translateY(-10px); }\n        }\n        \n        .upload-area p {\n            color: var(--gray);\n            font-size: 1.2em;\n            margin-bottom: 20px;\n            font-weight: 500;\n        }\n        \n        .file-input { display: none; }\n        \n        .upload-btn { \n            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);\n            color: white; \n            padding: 16px 40px; \n            border: none; \n            border-radius: 12px; \n            cursor: pointer; \n            font-size: 1.1em; \n            font-weight: 700;\n            transition: all 0.3s ease; \n            box-shadow: 0 4px 20px rgba(0, 102, 204, 0.3);\n            letter-spacing: 0.5px;\n        }\n        \n        .upload-btn:hover { \n            background: linear-gradient(135deg, var(--primary-dark) 0%, var(--primary) 100%);\n            transform: translateY(-3px);\n            box-shadow: 0 8px 30px rgba(0, 102, 204, 0.4);\n        }\n        \n        .upload-btn:active {\n            transform: translateY(-1px);\n        }\n        \n        .convert-btn { \n            background: linear-gradient(135deg, var(--secondary) 0%, var(--secondary-dark) 100%);\n            color: white; \n            padding: 18px 50px; \n            border: none; \n            border-radius: 12px; \n            cursor: pointer; \n            font-size: 1.3em; \n            font-weight: 700;\n            margin-top: 30px; \n            transition: all 0.3s ease; \n            box-shadow: 0 6px 25px rgba(0, 168, 150, 0.3);\n            letter-spacing: 0.5px;\n        }\n        \n        .convert-btn:hover { \n            background: linear-gradient(135deg, var(--secondary-dark) 0%, #006D5F 100%);\n            transform: translateY(-4px);\n            box-shadow: 0 10px 35px rgba(0, 168, 150, 0.4);\n        }\n        \n        .convert-btn:disabled { \n            background: linear-gradient(135deg, #CBD5E1 0%, #94A3B8 100%);\n            cursor: not-allowed; \n            transform: none;\n            box-shadow: none;\n        }\n        \n        /* File List */\n        .file-list { \n            margin-top: 30px; \n        }\n        \n        .file-item { \n            background: white;\n            padding: 18px 24px; \n            margin: 12px 0; \n            border-radius: 12px; \n            display: flex; \n            justify-content: space-between; \n            align-items: center; \n            border: 2px solid #E2E8F0;\n            transition: all 0.3s ease;\n            animation: slideIn 0.3s ease;\n        }\n        \n        @keyframes slideIn {\n            from { opacity: 0; transform: translateX(-20px); }\n            to { opacity: 1; transform: translateX(0); }\n        }\n        \n        .file-item:hover {\n            border-color: var(--primary);\n            box-shadow: 0 4px 15px rgba(0, 102, 204, 0.1);\n            transform: translateX(5px);\n        }\n        \n        .file-item span {\n            color: var(--dark);\n            font-weight: 600;\n            display: flex;\n            align-items: center;\n            gap: 10px;\n        }\n        \n        .file-item span::before {\n            content: \'📄\';\n            font-size: 1.5em;\n        }\n        \n        /* Progress Bar */\n        .progress { \n            width: 100%; \n            background: #E2E8F0; \n            border-radius: 50px; \n            margin: 30px 0; \n            height: 30px;\n            overflow: hidden;\n            box-shadow: inset 0 2px 8px rgba(0, 0, 0, 0.1);\n        }\n        \n        .progress-bar { \n            height: 100%; \n            background: linear-gradient(90deg, var(--primary) 0%, var(--secondary) 100%);\n            border-radius: 50px; \n            width: 0%; \n            transition: width 0.3s ease;\n            position: relative;\n            overflow: hidden;\n            animation: progressAnimation 1.5s ease infinite;\n        }\n        \n        @keyframes progressAnimation {\n            0% { background-position: 0% 50%; }\n            50% { background-position: 100% 50%; }\n            100% { background-position: 0% 50%; }\n        }\n        \n        .progress-bar::before {\n            content: \'\';\n            position: absolute;\n            top: 0;\n            left: -100%;\n            width: 100%;\n            height: 100%;\n            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);\n            animation: shimmer 2s infinite;\n        }\n        \n        @keyframes shimmer {\n            to { left: 100%; }\n        }\n        \n        /* Result Messages */\n        .result { \n            margin-top: 30px; \n            padding: 24px 28px; \n            background: linear-gradient(135deg, #D1FAE5 0%, #A7F3D0 100%);\n            border-radius: 16px; \n            color: #065F46; \n            border: 2px solid var(--success);\n            animation: slideIn 0.4s ease;\n            box-shadow: 0 4px 20px rgba(16, 185, 129, 0.2);\n        }\n        \n        .result h4 {\n            margin-bottom: 12px;\n            font-size: 1.3em;\n            font-weight: 700;\n            display: flex;\n            align-items: center;\n            gap: 10px;\n        }\n        \n        .result p {\n            font-size: 1.05em;\n            line-height: 1.6;\n        }\n        \n        .error { \n            margin-top: 30px; \n            padding: 24px 28px; \n            background: linear-gradient(135deg, #FEE2E2 0%, #FECACA 100%);\n            border-radius: 16px; \n            color: #991B1B; \n            border: 2px solid var(--error);\n            animation: shake 0.5s ease;\n            box-shadow: 0 4px 20px rgba(239, 68, 68, 0.2);\n        }\n        \n        @keyframes shake {\n            0%, 100% { transform: translateX(0); }\n            25% { transform: translateX(-10px); }\n            75% { transform: translateX(10px); }\n        }\n        \n        .error h4 {\n            margin-bottom: 12px;\n            font-size: 1.3em;\n            font-weight: 700;\n        }\n        \n        .hidden { display: none; }\n        \n        /* Back Button */\n        .back-btn { \n            background: rgba(255, 255, 255, 0.95);\n            color: var(--primary); \n            padding: 12px 28px; \n            border: 2px solid var(--primary);\n            border-radius: 12px; \n            cursor: pointer; \n            margin-bottom: 30px; \n            font-weight: 700;\n            font-size: 1.05em;\n            transition: all 0.3s ease;\n            display: inline-flex;\n            align-items: center;\n            gap: 8px;\n            box-shadow: 0 4px 15px rgba(0, 102, 204, 0.2);\n        }\n        \n        .back-btn:hover { \n            background: var(--primary);\n            color: white;\n            transform: translateX(-5px);\n            box-shadow: 0 6px 20px rgba(0, 102, 204, 0.3);\n        }\n        \n        /* Remove Button */\n        .remove-btn {\n            background: linear-gradient(135deg, var(--error) 0%, #DC2626 100%);\n            color: white;\n            border: none;\n            padding: 8px 20px;\n            border-radius: 8px;\n            cursor: pointer;\n            font-weight: 700;\n            font-size: 0.95em;\n            transition: all 0.3s ease;\n            box-shadow: 0 4px 15px rgba(239, 68, 68, 0.2);\n        }\n        \n        .remove-btn:hover {\n            background: linear-gradient(135deg, #DC2626 0%, #B91C1C 100%);\n            transform: translateY(-2px);\n            box-shadow: 0 6px 20px rgba(239, 68, 68, 0.3);\n        }\n        \n        /* Footer */\n        .footer { \n            text-align: center; \n            color: white; \n            margin-top: 80px; \n            padding: 40px 20px; \n            border-top: 1px solid rgba(255, 255, 255, 0.2);\n            background: rgba(0, 0, 0, 0.1);\n            backdrop-filter: blur(20px);\n            -webkit-backdrop-filter: blur(20px);\n            border-radius: 20px 20px 0 0;\n        }\n        \n        .footer p { \n            margin-bottom: 12px; \n            opacity: 0.95;\n            font-size: 1.05em;\n            font-weight: 500;\n        }\n        \n        .footer a { \n            color: #FFFFFF; \n            text-decoration: none; \n            font-weight: 700;\n            transition: all 0.3s;\n            padding: 4px 8px;\n            border-radius: 6px;\n        }\n        \n        .footer a:hover { \n            background: rgba(255, 255, 255, 0.2);\n            transform: translateY(-2px);\n        }\n        \n        .footer .security-note {\n            margin-top: 20px;\n            padding: 16px 24px;\n            background: rgba(255, 255, 255, 0.15);\n            backdrop-filter: blur(10px);\n            border-radius: 12px;\n            display: inline-block;\n            font-size: 0.95em;\n            border: 1px solid rgba(255, 255, 255, 0.2);\n        }\n        \n        /* Responsive */\n        @media (max-width: 768px) {\n            .header h1 { font-size: 2.5em; }\n            .tools-grid { grid-template-columns: 1fr; }\n            .tool-card { padding: 28px; }\n            .upload-area { padding: 40px 20px; }\n            .container { padding: 15px; }\n        }\n        \n        /* Loading Animation */\n        @keyframes spin {\n            to { transform: rotate(360deg); }\n        }\n        \n        .loading {\n            display: inline-block;\n            width: 20px;\n            height: 20px;\n            border: 3px solid rgba(255, 255, 255, 0.3);\n            border-radius: 50%;\n            border-top-color: white;\n            animation: spin 0.8s linear infinite;\n        }\n\n        /* Compress level selector */\n        .compress-level {\n            display: flex;\n            flex-direction: column;\n            align-items: center;\n            gap: 6px;\n            padding: 18px 24px;\n            border: 2px solid #E2E8F0;\n            border-radius: 16px;\n            cursor: pointer;\n            background: white;\n            transition: all 0.25s ease;\n            min-width: 130px;\n            box-shadow: 0 2px 8px rgba(0,0,0,0.05);\n        }\n        .compress-level strong { color: var(--dark); font-size: 1em; }\n        .compress-level small  { color: var(--gray); font-size: 0.78em; text-align:center; }\n        .compress-level:hover  { border-color: var(--primary); transform: translateY(-3px); box-shadow: 0 6px 20px rgba(0,102,204,0.15); }\n        .compress-level.active { border-color: var(--primary); background: linear-gradient(135deg,#EFF6FF,#DBEAFE); box-shadow: 0 6px 20px rgba(0,102,204,0.2); }\n        .compress-level.active strong { color: var(--primary); }\n    </style>\n</head>\n<body>\n    <div class="container">\n        <div class="header">\n            <div class="badge">🔒 100% Local & Seguro</div>\n            <h1>📄 LocalPDF</h1>\n            <p>Ferramentas PDF corporativas com total privacidade e segurança</p>\n        </div>\n\n        <div id="home-view">\n            <div class="tools-grid">\n                <div class="tool-card" onclick="showTool(\'pdf-to-images\')">\n                    <div class="icon">🖼️</div>\n                    <h3>PDF para Imagens</h3>\n                    <p>Converta páginas PDF em imagens JPG ou PNG</p>\n                </div>\n                <div class="tool-card" onclick="showTool(\'images-to-pdf\')">\n                    <div class="icon">📄</div>\n                    <h3>Imagens para PDF</h3>\n                    <p>Combine várias imagens em um único PDF</p>\n                </div>\n                <div class="tool-card" onclick="showTool(\'merge-pdf\')">\n                    <div class="icon">🔗</div>\n                    <h3>Mesclar PDFs</h3>\n                    <p>Combine vários PDFs em um documento único</p>\n                </div>\n                <div class="tool-card" onclick="showTool(\'split-pdf\')">\n                    <div class="icon">✂️</div>\n                    <h3>Dividir PDF</h3>\n                    <p>Extraia páginas específicas do seu PDF</p>\n                </div>\n                <div class="tool-card" onclick="showTool(\'compress-pdf\')">\n                    <div class="icon">📦</div>\n                    <h3>Comprimir PDF</h3>\n                    <p>Reduza o tamanho do seu arquivo PDF</p>\n                </div>\n                <div class="tool-card" onclick="showTool(\'pdf-to-pdfa\')">\n                    <div class="icon">🔒</div>\n                    <h3>PDF para PDF/A</h3>\n                    <p>Padronize seu PDF para arquivamento</p>\n                </div>\n                <div class="tool-card" onclick="showTool(\'word-to-pdf\')">\n                    <div class="icon">📝</div>\n                    <h3>Word para PDF</h3>\n                    <p>Converta documentos DOCX para PDF</p>\n                </div>\n                <div class="tool-card" onclick="showTool(\'excel-to-pdf\')">\n                    <div class="icon">📊</div>\n                    <h3>Excel para PDF</h3>\n                    <p>Converta planilhas XLSX para PDF</p>\n                </div>\n                <div class="tool-card" onclick="showTool(\'txt-to-pdf\')">\n                    <div class="icon">📃</div>\n                    <h3>TXT para PDF</h3>\n                    <p>Converta arquivos de texto para PDF</p>\n                </div>\n                <div class="tool-card" onclick="showTool(\'pdf-to-word\')">\n                    <div class="icon">🔄</div>\n                    <h3>PDF para Word</h3>\n                    <p>Converta PDF para Word editável</p>\n                </div>\n            </div>\n        </div>\n\n        <!-- Tool Views -->\n        <div id="tool-views" class="hidden">\n            <button class="back-btn" onclick="showHome()">&larr; Voltar</button>\n            <div class="tool-card">\n                <h3 id="tool-title"></h3>\n                <p id="tool-description"></p>\n\n                <div class="upload-area" id="upload-area" onclick="document.getElementById(\'file-input\').click()">\n                    <input type="file" id="file-input" class="file-input" multiple accept=".pdf,.docx,.jpg,.jpeg,.png,.txt,.xlsx">\n                    <span class="upload-icon">📁</span>\n                    <p>Clique aqui ou arraste arquivos para fazer upload</p>\n                    <button class="upload-btn">Escolher Arquivos</button>\n                </div>\n\n                <div id="file-list" class="file-list"></div>\n\n                <button id="convert-btn" class="convert-btn hidden" onclick="convertFiles()">🚀 Converter Agora</button>\n\n                <!-- Seletor de nivel de compressao (so aparece no compress-pdf) -->\n                <div id="compress-options" class="hidden" style="margin:24px 0;">\n                    <p style="font-weight:600;color:var(--dark);margin-bottom:14px;font-size:1.05em;">Nivel de Compressao</p>\n                    <div style="display:flex;gap:12px;flex-wrap:wrap;justify-content:center;">\n                        <label class="compress-level active" data-level="screen">\n                            <input type="radio" name="compress_level" value="screen" checked style="display:none">\n                            <span style="font-size:1.8em;">🔥</span>\n                            <strong>Maxima</strong>\n                            <small>72 dpi - menor tamanho</small>\n                        </label>\n                        <label class="compress-level" data-level="ebook">\n                            <input type="radio" name="compress_level" value="ebook" style="display:none">\n                            <span style="font-size:1.8em;">⚖️</span>\n                            <strong>Balanceada</strong>\n                            <small>150 dpi - qualidade ok</small>\n                        </label>\n                        <label class="compress-level" data-level="printer">\n                            <input type="radio" name="compress_level" value="printer" style="display:none">\n                            <span style="font-size:1.8em;">💎</span>\n                            <strong>Leve</strong>\n                            <small>300 dpi - maior qualidade</small>\n                        </label>\n                    </div>\n                </div>\n\n                <div id="progress" class="progress hidden">\n                    <div id="progress-bar" class="progress-bar"></div>\n                </div>\n                <p id="progress-msg" style="text-align:center;color:var(--gray);margin:8px 0 0;font-size:0.9em;min-height:1.3em;"></p>\n\n                <div id="result" class="hidden"></div>\n            </div>\n        </div>\n\n        <div class="footer">\n            <p><strong>LocalPDF</strong> - Ferramenta Corporativa Interna</p>\n            <p>\n                <a href="mailto:ti-infra@neogenomica.com.br">✉️ ti-infra@neogenomica.com.br</a>\n            </p>\n            <div class="security-note">\n                🛡️ Processamento 100% local • Seus arquivos nunca saem da infraestrutura interna\n            </div>\n        </div>\n    </div>\n\n    <script>\n        let currentTool = \'\';\n        let uploadedFiles = [];\n\n        const tools = {\n            \'pdf-to-images\': { title: \'PDF para Imagens\',    description: \'Converta cada pagina do seu PDF em imagens separadas\', accept: \'.pdf\',              multiple: false },\n            \'images-to-pdf\': { title: \'Imagens para PDF\',    description: \'Combine multiplas imagens em um unico arquivo PDF\',    accept: \'.jpg,.jpeg,.png\',   multiple: true  },\n            \'merge-pdf\':     { title: \'Mesclar PDFs\',        description: \'Combine varios arquivos PDF em um documento unico\',    accept: \'.pdf\',              multiple: true  },\n            \'split-pdf\':     { title: \'Dividir PDF\',         description: \'Extraia paginas especificas do seu PDF\',               accept: \'.pdf\',              multiple: false },\n            \'compress-pdf\':  { title: \'Comprimir PDF\',       description: \'Reduza o tamanho do arquivo PDF mantendo a qualidade\', accept: \'.pdf\',              multiple: false },\n            \'pdf-to-pdfa\':   { title: \'PDF para PDF/A\',      description: \'Converta PDFs para o padrao de arquivamento PDF/A-1b\', accept: \'.pdf\',              multiple: true  },\n            \'word-to-pdf\':   { title: \'Word para PDF\',       description: \'Converta documentos Word (.docx) para PDF\',           accept: \'.docx\',             multiple: true  },\n            \'excel-to-pdf\':  { title: \'Excel para PDF\',      description: \'Converta planilhas Excel (.xlsx) para PDF\',           accept: \'.xlsx\',             multiple: false },\n            \'txt-to-pdf\':    { title: \'TXT para PDF\',        description: \'Converta arquivos de texto simples (.txt) para PDF\',  accept: \'.txt\',              multiple: false },\n            \'pdf-to-word\':   { title: \'PDF para Word\',       description: \'Converta seus documentos PDF para Word (.docx)\',      accept: \'.pdf\',              multiple: false }\n        };\n\n        function showTool(toolName) {\n            currentTool = toolName;\n            const tool = tools[toolName];\n            document.getElementById(\'home-view\').classList.add(\'hidden\');\n            document.getElementById(\'tool-views\').classList.remove(\'hidden\');\n            document.getElementById(\'tool-title\').innerText       = tool.title;\n            document.getElementById(\'tool-description\').innerText = tool.description;\n            document.getElementById(\'file-input\').accept   = tool.accept;\n            document.getElementById(\'file-input\').multiple = tool.multiple;\n            uploadedFiles = [];\n            updateFileList();\n            hideResult();\n\n            // Mostrar seletor de compressao apenas para compress-pdf\n            const compressOpts = document.getElementById(\'compress-options\');\n            if (compressOpts) {\n                compressOpts.classList.toggle(\'hidden\', toolName !== \'compress-pdf\');\n            }\n        }\n\n        function showHome() {\n            document.getElementById(\'home-view\').classList.remove(\'hidden\');\n            document.getElementById(\'tool-views\').classList.add(\'hidden\');\n            uploadedFiles = [];\n        }\n\n        function updateFileList() {\n            const fileList   = document.getElementById(\'file-list\');\n            const convertBtn = document.getElementById(\'convert-btn\');\n            if (uploadedFiles.length === 0) {\n                fileList.innerHTML = \'\';\n                convertBtn.classList.add(\'hidden\');\n                return;\n            }\n            fileList.innerHTML = uploadedFiles.map((file, index) => `\n                <div class="file-item">\n                    <span>${file.name} <small style="opacity:0.7">(${(file.size/1024/1024).toFixed(2)} MB)</small></span>\n                    <button onclick="removeFile(${index})" class="remove-btn">Remover</button>\n                </div>`).join(\'\');\n            convertBtn.classList.remove(\'hidden\');\n        }\n\n        function removeFile(index) {\n            uploadedFiles.splice(index, 1);\n            updateFileList();\n        }\n\n        function hideResult() {\n            document.getElementById(\'result\').classList.add(\'hidden\');\n            document.getElementById(\'progress\').classList.add(\'hidden\');\n            const msg = document.getElementById(\'progress-msg\');\n            if (msg) msg.textContent = \'\';\n        }\n\n        // Seletor de nivel de compressao\n        document.addEventListener(\'click\', function(e) {\n            const label = e.target.closest(\'.compress-level\');\n            if (!label) return;\n            document.querySelectorAll(\'.compress-level\').forEach(el => el.classList.remove(\'active\'));\n            label.classList.add(\'active\');\n            label.querySelector(\'input[type=radio]\').checked = true;\n        });\n\n        document.getElementById(\'file-input\').addEventListener(\'change\', function(e) {\n            const files = Array.from(e.target.files);\n            uploadedFiles = tools[currentTool].multiple ? uploadedFiles.concat(files) : files.slice(0,1);\n            updateFileList();\n        });\n\n        const uploadArea = document.getElementById(\'upload-area\');\n        uploadArea.addEventListener(\'dragover\',  e => { e.preventDefault(); uploadArea.classList.add(\'dragover\'); });\n        uploadArea.addEventListener(\'dragleave\', e => { e.preventDefault(); uploadArea.classList.remove(\'dragover\'); });\n        uploadArea.addEventListener(\'drop\', function(e) {\n            e.preventDefault();\n            uploadArea.classList.remove(\'dragover\');\n            const files = Array.from(e.dataTransfer.files);\n            uploadedFiles = tools[currentTool].multiple ? uploadedFiles.concat(files) : files.slice(0,1);\n            updateFileList();\n        });\n\n        async function convertFiles() {\n            if (uploadedFiles.length === 0) return;\n\n            const formData = new FormData();\n            uploadedFiles.forEach(file => formData.append(\'files\', file));\n            formData.append(\'tool\', currentTool);\n\n            // Adiciona nivel de compressao se for compress-pdf\n            if (currentTool === \'compress-pdf\') {\n                const selected = document.querySelector(\'input[name=compress_level]:checked\');\n                formData.append(\'compress_level\', selected ? selected.value : \'ebook\');\n            }\n\n            const progressBar = document.getElementById(\'progress-bar\');\n            const progressMsg = document.getElementById(\'progress-msg\');\n            document.getElementById(\'progress\').classList.remove(\'hidden\');\n            document.getElementById(\'convert-btn\').disabled = true;\n            hideResult();\n\n            progressBar.style.width = \'5%\';\n            if (progressMsg) progressMsg.textContent = \'Enviando arquivos...\';\n\n            try {\n                const response = await fetch(\'/convert\', { method: \'POST\', body: formData });\n                if (!response.ok) {\n                    const err = await response.json();\n                    throw new Error(err.error || \'Erro ao iniciar conversao\');\n                }\n\n                const { task_id } = await response.json();\n\n                await new Promise((resolve, reject) => {\n                    const interval = setInterval(async () => {\n                        try {\n                            const res = await fetch(`/progress/${task_id}`);\n                            if (!res.ok) { clearInterval(interval); reject(new Error(\'Erro ao verificar progresso\')); return; }\n                            const prog = await res.json();\n                            progressBar.style.width = prog.progress + \'%\';\n                            if (progressMsg) progressMsg.textContent = prog.message;\n                            if (prog.status === \'done\')  { clearInterval(interval); progressBar.style.width = \'100%\'; resolve(); }\n                            if (prog.status === \'error\') { clearInterval(interval); reject(new Error(prog.message)); }\n                        } catch(e) { clearInterval(interval); reject(e); }\n                    }, 600);\n                });\n\n                if (progressMsg) progressMsg.textContent = \'Baixando arquivo...\';\n                const a = document.createElement(\'a\');\n                a.href = `/download/${task_id}`;\n                a.download = \'\';\n                document.body.appendChild(a); a.click(); document.body.removeChild(a);\n\n                document.getElementById(\'result\').className = \'result\';\n                document.getElementById(\'result\').innerHTML  = \'<h4>Sucesso!</h4><p>Arquivo convertido e baixado com sucesso!</p>\';\n                document.getElementById(\'result\').classList.remove(\'hidden\');\n\n            } catch (error) {\n                document.getElementById(\'result\').className = \'error\';\n                document.getElementById(\'result\').innerHTML  = `<h4>Erro!</h4><p>${error.message || \'Ocorreu um erro. Tente novamente.\'}</p>`;\n                document.getElementById(\'result\').classList.remove(\'hidden\');\n            } finally {\n                setTimeout(() => {\n                    document.getElementById(\'progress\').classList.add(\'hidden\');\n                    progressBar.style.width = \'0%\';\n                    if (progressMsg) progressMsg.textContent = \'\';\n                }, 2000);\n                document.getElementById(\'convert-btn\').disabled = false;\n            }\n        }\n    </script>\n</body>\n</html>\n'
)


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/convert", methods=["POST"])
def convert():
    if "files" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    files = request.files.getlist("files")
    tool  = request.form.get("tool")

    if not files or files[0].filename == "":
        return jsonify({"error": "Nenhum arquivo selecionado"}), 400

    for f in files:
        if not allowed_file(f.filename):
            return jsonify({"error": f"Extensao nao permitida: {f.filename}"}), 400

    task_id  = str(uuid.uuid4())
    temp_dir = tempfile.mkdtemp()

    saved_paths = []
    for f in files:
        path = os.path.join(temp_dir, secure_filename(f.filename))
        f.save(path)
        saved_paths.append(path)

    # Parametros extras
    extra = {}
    if tool == "compress-pdf":
        level = request.form.get("compress_level", "ebook")
        if level not in ("screen", "ebook", "printer"):
            level = "ebook"
        extra["compress_level"] = level

    with tasks_lock:
        tasks[task_id] = {
            "progress":    0,
            "status":      "processing",
            "message":     "Aguardando inicio...",
            "result_path": None,
            "temp_dir":    temp_dir,
        }

    threading.Thread(
        target=_process_in_background,
        args=(task_id, tool, saved_paths, temp_dir, extra),
        daemon=True,
    ).start()

    return jsonify({"task_id": task_id})


@app.route("/progress/<task_id>")
def get_progress(task_id):
    with tasks_lock:
        task = tasks.get(task_id)
    if not task:
        return jsonify({"error": "Task nao encontrada"}), 404
    return jsonify({
        "progress": task["progress"],
        "status":   task["status"],
        "message":  task["message"],
    })


@app.route("/download/<task_id>")
def download_file(task_id):
    with tasks_lock:
        task = tasks.get(task_id)
    if not task:
        return jsonify({"error": "Task nao encontrada"}), 404
    if task["status"] != "done":
        return jsonify({"error": "Arquivo ainda nao esta pronto"}), 202
    if not task["result_path"] or not os.path.exists(task["result_path"]):
        return jsonify({"error": "Arquivo de resultado nao encontrado"}), 500

    result_path = task["result_path"]
    temp_dir    = task["temp_dir"]
    filename    = os.path.basename(result_path)

    with open(result_path, "rb") as fh:
        data = fh.read()

    def _cleanup():
        import time; time.sleep(10)
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
        with tasks_lock:
            tasks.pop(task_id, None)

    threading.Thread(target=_cleanup, daemon=True).start()
    return send_file(io.BytesIO(data), as_attachment=True, download_name=filename)


def _process_in_background(task_id: str, tool: str, saved_paths: list, temp_dir: str, extra: dict = None):
    if extra is None:
        extra = {}
    try:
        files = [SavedFile(p) for p in saved_paths]
        set_progress(task_id, 8, "Iniciando processamento...")

        compress_level = extra.get("compress_level", "ebook")

        dispatch = {
            "pdf-to-images": lambda: pdf_to_images(files[0], temp_dir, task_id),
            "images-to-pdf": lambda: images_to_pdf(files, temp_dir, task_id),
            "merge-pdf":     lambda: merge_pdfs(files, temp_dir, task_id),
            "split-pdf":     lambda: split_pdf(files[0], temp_dir, task_id),
            "compress-pdf":  lambda: compress_pdf(files[0], temp_dir, task_id, level=compress_level),
            "pdf-to-pdfa":   lambda: pdf_to_pdfa(files, temp_dir, task_id),
            "word-to-pdf":   lambda: word_to_pdf(files, temp_dir, task_id),
            "excel-to-pdf":  lambda: excel_to_pdf(files[0], temp_dir, task_id),
            "txt-to-pdf":    lambda: txt_to_pdf(files[0], temp_dir, task_id),
            "pdf-to-word":   lambda: pdf_to_word(files[0], temp_dir, task_id),
        }

        if tool not in dispatch:
            raise ValueError(f"Ferramenta nao suportada: {tool}")

        output_files = dispatch[tool]()
        set_progress(task_id, 95, "Preparando arquivo para download...")
        result_path = _build_result(output_files, temp_dir)

        with tasks_lock:
            tasks[task_id]["progress"]    = 100
            tasks[task_id]["status"]      = "done"
            tasks[task_id]["message"]     = "Concluido com sucesso!"
            tasks[task_id]["result_path"] = result_path

    except Exception as exc:
        with tasks_lock:
            if task_id in tasks:
                tasks[task_id]["status"]   = "error"
                tasks[task_id]["message"]  = str(exc)
                tasks[task_id]["progress"] = 0


def _build_result(output_files, temp_dir):
    if not isinstance(output_files, list):
        output_files = [output_files]
    if len(output_files) == 1:
        return output_files[0]
    zip_path = os.path.join(temp_dir, "converted_files.zip")
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for fp in output_files:
            zipf.write(fp, os.path.basename(fp))
    return zip_path


def pdf_to_images(file, temp_dir, task_id=None):
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    doc = fitz.open(pdf_path)
    total = len(doc)
    output_files = []
    for page_num in range(total):
        if task_id:
            set_progress(task_id, 10 + int(page_num / total * 80),
                         f"Convertendo pagina {page_num + 1} de {total}...")
        page = doc.load_page(page_num)
        pix  = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_path = os.path.join(temp_dir, f"page_{page_num + 1}.png")
        pix.save(img_path)
        output_files.append(img_path)
    doc.close()
    return output_files


def images_to_pdf(files, temp_dir, task_id=None):
    total  = len(files)
    images = []
    for i, file in enumerate(files):
        if task_id:
            set_progress(task_id, 10 + int(i / total * 70),
                         f"Processando imagem {i + 1} de {total}...")
        img_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(img_path)
        img = Image.open(img_path)
        if img.mode != "RGB":
            img = img.convert("RGB")
        images.append(img)
    if task_id:
        set_progress(task_id, 85, "Gerando PDF...")
    pdf_path = os.path.join(temp_dir, "images_to_pdf.pdf")
    images[0].save(pdf_path, save_all=True, append_images=images[1:])
    return [pdf_path]


def merge_pdfs(files, temp_dir, task_id=None):
    total      = len(files)
    merged_doc = fitz.open()
    for i, file in enumerate(files):
        if task_id:
            set_progress(task_id, 10 + int(i / total * 80),
                         f"Mesclando arquivo {i + 1} de {total}...")
        pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(pdf_path)
        doc = fitz.open(pdf_path)
        merged_doc.insert_pdf(doc)
        doc.close()
    if task_id:
        set_progress(task_id, 90, "Salvando PDF final...")
    output_path = os.path.join(temp_dir, "merged.pdf")
    merged_doc.save(output_path)
    merged_doc.close()
    return [output_path]


def split_pdf(file, temp_dir, task_id=None):
    pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    doc   = fitz.open(pdf_path)
    total = len(doc)
    output_files = []
    for page_num in range(total):
        if task_id:
            set_progress(task_id, 10 + int(page_num / total * 80),
                         f"Extraindo pagina {page_num + 1} de {total}...")
        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
        output_path = os.path.join(temp_dir, f"page_{page_num + 1}.pdf")
        new_doc.save(output_path)
        new_doc.close()
        output_files.append(output_path)
    doc.close()
    return output_files


def compress_pdf(file, temp_dir, task_id=None, level="ebook"):
    # level: screen=72dpi  ebook=150dpi  printer=300dpi
    LEVEL_NAMES = {"screen": "Maxima (72dpi)", "ebook": "Balanceada (150dpi)", "printer": "Leve (300dpi)"}
    pdf_path    = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    output_path = os.path.join(temp_dir, "compressed.pdf")

    if task_id:
        set_progress(task_id, 20, f"Compressao {LEVEL_NAMES.get(level, level)} - analisando...")

    gs_args = [
        "gs", "-dBATCH", "-dNOPAUSE", "-dQUIET", "-dSAFER",
        f"-dPDFSETTINGS=/{level}",
        f"-dColorImageResolution={72 if level == 'screen' else 150 if level == 'ebook' else 300}",
        f"-dGrayImageResolution={72 if level == 'screen' else 150 if level == 'ebook' else 300}",
        f"-dMonoImageResolution={72 if level == 'screen' else 150 if level == 'ebook' else 300}",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-sOutputFile={output_path}",
        pdf_path,
    ]
    gs_bytes = [a.encode("utf-8") if isinstance(a, str) else a for a in gs_args]

    if task_id:
        set_progress(task_id, 40, f"Comprimindo com Ghostscript ({LEVEL_NAMES.get(level, level)})...")

    try:
        ghostscript.Ghostscript(*gs_bytes)
    except Exception as e:
        raise RuntimeError(f"Erro ao comprimir PDF: {e}") from e

    if task_id:
        set_progress(task_id, 85, "Verificando resultado...")

    orig_size = os.path.getsize(pdf_path)
    comp_size = os.path.getsize(output_path)
    if comp_size >= orig_size:
        return [pdf_path]
    return [output_path]


def pdf_to_pdfa(files, temp_dir, task_id=None):
    if not isinstance(files, list):
        files = [files]
    total        = len(files)
    output_files = []
    for i, file in enumerate(files):
        if task_id:
            set_progress(task_id, 10 + int(i / total * 80),
                         f"Convertendo {file.filename} para PDF/A ({i + 1}/{total})...")
        input_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(input_path)
        base_name, _ = os.path.splitext(os.path.basename(input_path))
        output_path  = os.path.join(temp_dir, f"{base_name}_pdfa.pdf")
        gs_args = [
            "gs", "-dPDFA=1", "-dBATCH", "-dNOPAUSE", "-dNOOUTERSAVE",
            "-dUseCIEColor", "-sProcessColorModel=DeviceRGB", "-sDEVICE=pdfwrite",
            "-sColorConversionStrategy=UseDeviceIndependentColor",
            "-dPDFACompatibilityPolicy=1",
            f"-sOutputFile={output_path}", input_path,
        ]
        gs_args = [a.encode("utf-8") if isinstance(a, str) else a for a in gs_args]
        try:
            ghostscript.Ghostscript(*gs_args)
        except Exception as e:
            raise RuntimeError(f"Erro ao converter {file.filename} para PDF/A: {e}") from e
        output_files.append(output_path)
    return output_files


def word_to_pdf(files, temp_dir, task_id=None):
    from docx import Document

    if not isinstance(files, list):
        files = [files]

    pdf_path      = os.path.join(temp_dir, "word_to_pdf.pdf")
    c             = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter
    y_position    = height - 50
    total         = len(files)

    for file_idx, file in enumerate(files):
        if task_id:
            set_progress(task_id, 10 + int(file_idx / total * 80),
                         f"Convertendo {file.filename} ({file_idx + 1}/{total})...")
        docx_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(docx_path)
        doc = Document(docx_path)

        if file_idx > 0:
            c.showPage(); y_position = height - 50
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y_position, "=" * 60); y_position -= 20
            c.drawString(50, y_position, f"Documento: {file.filename}"); y_position -= 20
            c.drawString(50, y_position, "=" * 60); y_position -= 30
            c.setFont("Helvetica", 11)

        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                chars_per_line = int((width - 100) / 6)
                words, lines, current_line = paragraph.text.split(), [], []
                for word in words:
                    if len(" ".join(current_line + [word])) <= chars_per_line:
                        current_line.append(word)
                    else:
                        if current_line: lines.append(" ".join(current_line))
                        current_line = [word]
                if current_line: lines.append(" ".join(current_line))
                for line in lines:
                    if y_position < 50: c.showPage(); y_position = height - 50
                    c.drawString(50, y_position, line); y_position -= 20

        for table in doc.tables:
            y_position -= 10
            if y_position < 100: c.showPage(); y_position = height - 50
            c.setFont("Helvetica", 9)
            for row in table.rows:
                row_text = " | ".join(cell.text for cell in row.cells)
                if len(row_text) > 100: row_text = row_text[:97] + "..."
                if y_position < 50: c.showPage(); y_position = height - 50
                c.drawString(50, y_position, row_text); y_position -= 15
            y_position -= 10; c.setFont("Helvetica", 11)

    c.save()
    return [pdf_path]


def excel_to_pdf(file, temp_dir, task_id=None):
    xlsx_path     = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(xlsx_path)
    pdf_path      = os.path.join(temp_dir, "excel_to_pdf.pdf")
    c             = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter
    y_position    = height - 50

    if task_id: set_progress(task_id, 20, "Lendo planilha...")

    try:
        workbook     = openpyxl.load_workbook(xlsx_path)
        total_sheets = len(workbook.sheetnames)
        for sheet_idx, sheet_name in enumerate(workbook.sheetnames):
            if task_id:
                set_progress(task_id, 20 + int(sheet_idx / total_sheets * 65),
                             "Processando aba " + sheet_name + "...")
            sheet = workbook[sheet_name]
            c.setFont("Helvetica", 10)
            c.drawString(50, y_position, f"--- Planilha: {sheet_name} ---"); y_position -= 20
            for row in sheet.iter_rows():
                row_data  = [str(cell.value) if cell.value is not None else "" for cell in row]
                line_text = " | ".join(row_data)
                max_chars = int((width - 100) / 6)
                if len(line_text) > max_chars: line_text = line_text[:max_chars] + "..."
                if y_position < 50:
                    c.showPage(); y_position = height - 50; c.setFont("Helvetica", 10)
                c.drawString(50, y_position, line_text); y_position -= 15
            y_position -= 30
            if y_position < 50 and sheet_name != workbook.sheetnames[-1]:
                c.showPage(); y_position = height - 50
    except Exception as e:
        c.drawString(50, y_position - 20, f"Erro ao ler planilha: {e}")

    c.save()
    return [pdf_path]


def txt_to_pdf(file, temp_dir, task_id=None):
    txt_path      = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(txt_path)
    pdf_path      = os.path.join(temp_dir, "text_to_pdf.pdf")
    c             = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter
    y_position    = height - 50
    c.setFont("Helvetica", 12)

    if task_id: set_progress(task_id, 20, "Lendo arquivo de texto...")

    try:
        with open(txt_path, "r", encoding="utf-8") as fh:
            lines = fh.readlines()
        total = len(lines)
        for i, line in enumerate(lines):
            if task_id and i % 50 == 0:
                set_progress(task_id, 20 + int(i / max(total, 1) * 70),
                             f"Processando linha {i + 1} de {total}...")
            text_line      = line.strip()
            chars_per_line = int((width - 100) / 7)
            chunks = (
                [text_line[j:j + chars_per_line] for j in range(0, len(text_line), chars_per_line)]
                if text_line else [""]
            )
            for chunk in chunks:
                if y_position < 50:
                    c.showPage(); y_position = height - 50; c.setFont("Helvetica", 12)
                c.drawString(50, y_position, chunk); y_position -= 15
    except Exception as e:
        c.drawString(50, y_position - 20, f"Erro ao ler arquivo: {e}")

    c.save()
    return [pdf_path]


def pdf_to_word(file, temp_dir, task_id=None):
    pdf_path      = os.path.join(temp_dir, secure_filename(file.filename))
    file.save(pdf_path)
    docx_filename = os.path.splitext(secure_filename(file.filename))[0] + ".docx"
    docx_path     = os.path.join(temp_dir, docx_filename)

    if task_id: set_progress(task_id, 20, "Analisando PDF...")

    cv = None
    try:
        cv = Converter(pdf_path)
        if task_id: set_progress(task_id, 40, "Convertendo para Word (pode demorar)...")
        cv.convert(docx_path)
        if task_id: set_progress(task_id, 85, "Finalizando...")
    except ValueError as e:
        raise RuntimeError(f"Erro no arquivo PDF: {e}") from e
    except ConversionException as e:
        raise RuntimeError(f"Erro interno na conversao: {e}") from e
    except Exception as e:
        raise RuntimeError(f"Erro ao converter {file.filename} para Word: {e}") from e
    finally:
        if cv: cv.close()

    return [docx_path]


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
