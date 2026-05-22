#!/usr/bin/env python3
"""
별따오기 이름표 생성기
"""

import sys, os, re, csv, requests, tempfile, subprocess, shutil
from pathlib import Path

# ── 버전 / 업데이트 설정 ──────────────────────────────────
APP_VERSION = "1.0.0"
# GitHub raw URL — 레포 만든 뒤 아래 두 줄을 실제 주소로 교체
GITHUB_USER   = "Damiyu"
GITHUB_REPO   = "nametag-making-exe"
VERSION_URL   = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main/version.json"
DOWNLOAD_URL  = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/latest/download/별따오기_이름표생성기.exe"
# ────────────────────────────────────────────────────────

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QListWidget, QListWidgetItem,
    QSpinBox, QDoubleSpinBox, QColorDialog, QComboBox, QCheckBox,
    QGroupBox, QLineEdit, QMessageBox, QProgressDialog, QSizePolicy,
    QDialog, QDialogButtonBox, QCompleter, QFrame,
    QStyledItemDelegate, QStyle, QScrollArea, QRubberBand, QDesktopWidget
)
from PyQt5.QtCore import Qt, pyqtSignal, QTimer, QSize, QThread, QObject, QRect, QPoint
from PyQt5.QtGui import QPixmap, QImage, QColor, QFont, QFontDatabase, QCursor, QPainter, QPen

from PIL import Image, ImageDraw, ImageFont
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader
import io


# ═══════════════════════════════════════════════════════
# 시스템 폰트 스캐너
# ═══════════════════════════════════════════════════════

def scan_system_fonts():
    """설치된 폰트 전부 스캔 → {표시이름: 경로} 딕셔너리"""
    dirs = []
    if sys.platform == 'win32':
        win_fonts = os.path.join(os.environ.get('WINDIR', 'C:/Windows'), 'Fonts')
        user_fonts = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Windows', 'Fonts')
        dirs = [win_fonts, user_fonts]
    elif sys.platform == 'darwin':
        dirs = ['/System/Library/Fonts', '/Library/Fonts',
                os.path.expanduser('~/Library/Fonts')]
    else:
        dirs = ['/usr/share/fonts', '/usr/local/share/fonts',
                os.path.expanduser('~/.fonts'), os.path.expanduser('~/.local/share/fonts')]

    fonts = {}
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for root, _, files in os.walk(d):
            for f in files:
                if f.lower().endswith(('.ttf', '.otf')):
                    full = os.path.join(root, f)
                    # 폰트 패밀리 이름 읽기
                    display = _read_font_name(full) or Path(f).stem
                    # 중복이면 파일명 suffix 추가
                    key = display
                    if key in fonts:
                        key = f"{display} ({Path(f).stem})"
                    fonts[key] = full
    return dict(sorted(fonts.items()))


def _read_font_name(path):
    """TTF/OTF name 테이블에서 폰트 패밀리명 읽기"""
    try:
        with open(path, 'rb') as f:
            data = f.read()
        # OTF/TTF offset table
        import struct
        sfVersion = data[:4]
        numTables = struct.unpack('>H', data[4:6])[0]
        # table directory
        for i in range(numTables):
            off = 12 + i * 16
            tag = data[off:off+4]
            if tag == b'name':
                table_off = struct.unpack('>I', data[off+8:off+12])[0]
                count = struct.unpack('>H', data[table_off+2:table_off+4])[0]
                string_off = struct.unpack('>H', data[table_off+4:table_off+6])[0]
                for j in range(count):
                    r = table_off + 6 + j * 12
                    platformID = struct.unpack('>H', data[r:r+2])[0]
                    nameID = struct.unpack('>H', data[r+6:r+8])[0]
                    length = struct.unpack('>H', data[r+8:r+10])[0]
                    soff = struct.unpack('>H', data[r+10:r+12])[0]
                    if nameID == 1:  # Font Family
                        raw = data[table_off + string_off + soff:
                                   table_off + string_off + soff + length]
                        if platformID == 3:
                            return raw.decode('utf-16-be', errors='ignore').strip()
                        else:
                            return raw.decode('latin-1', errors='ignore').strip()
    except Exception:
        pass
    return None


# ═══════════════════════════════════════════════════════
# 유틸
# ═══════════════════════════════════════════════════════

def pil_to_qpixmap(pil_img):
    pil_img = pil_img.convert("RGBA")
    data = pil_img.tobytes("raw", "RGBA")
    qimg = QImage(data, pil_img.width, pil_img.height, QImage.Format_RGBA8888)
    return QPixmap.fromImage(qimg)


def detect_group_number(filename):
    name = Path(filename).stem
    m = re.search(r'(\d+)\s*[조組]', name)
    if m:
        return int(m.group(1))
    m = re.search(r'(\d+)', name)
    if m:
        return int(m.group(1))
    return None


def load_names_from_file(filepath):
    groups = {}
    ext = Path(filepath).suffix.lower()
    rows = []
    if ext == '.csv':
        with open(filepath, encoding='utf-8-sig') as f:
            rows = list(csv.reader(f))
    elif ext in ('.xlsx', '.xls'):
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        rows = [[str(c.value or '').strip() for c in row] for row in ws.iter_rows()]
    if not rows:
        return groups
    header = [c.lower() for c in rows[0]]
    start = 1 if any(k in header for k in ['조', '이름', 'group', 'name']) else 0
    for row in rows[start:]:
        if len(row) < 2:
            continue
        m = re.search(r'(\d+)', str(row[0]).strip())
        name = str(row[1]).strip()
        if m and name:
            groups.setdefault(int(m.group(1)), []).append(name)
    return groups


def download_font(url):
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        suffix = '.woff2' if 'woff2' in url else '.woff' if 'woff' in url else '.ttf'
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(r.content)
        tmp.close()
        return tmp.name
    except Exception:
        return None


# ═══════════════════════════════════════════════════════
# 자동 폰트 크기 계산
# ═══════════════════════════════════════════════════════

def auto_font_size(img_w, img_h, name, font_path,
                   target_w_ratio=0.55, target_h_ratio=0.12):
    """
    이미지 크기 기준으로 이름이 적당히 차지하도록 폰트 크기 자동 결정.
    - 가로: 이미지 너비의 target_w_ratio 이하
    - 세로: 이미지 높이의 target_h_ratio 이하
    두 조건 중 작은 값 채택.
    """
    target_w = img_w * target_w_ratio
    target_h = img_h * target_h_ratio

    size = 200  # 위에서부터 내려오며 탐색
    while size > 10:
        try:
            if font_path and os.path.exists(font_path):
                font = ImageFont.truetype(font_path, size)
            else:
                # 기본 폰트는 크기 제한이 있어서 스킵
                return max(int(img_h * 0.10), 40)
        except Exception:
            return max(int(img_h * 0.10), 40)

        dummy = Image.new('RGBA', (1, 1))
        draw = ImageDraw.Draw(dummy)
        bbox = draw.textbbox((0, 0), name, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]

        if tw <= target_w and th <= target_h:
            return size
        size -= 4

    return max(size, 20)


# ═══════════════════════════════════════════════════════
# 텍스트 스타일
# ═══════════════════════════════════════════════════════

class TextStyle:
    def __init__(self):
        self.font_path = None
        self.font_size = 0          # 0 = 자동
        self.auto_size = True
        self.color = [255, 220, 80]
        self.outline = True
        self.outline_color = [255, 255, 255]
        self.outline_width = 8
        self.shadow = True
        self.shadow_color = [0, 80, 160]
        self.shadow_angle = 135        # 도 (0=오른쪽, 90=아래, 135=오른쪽아래)
        self.shadow_distance = 12      # px
        self.shadow_blur = 6           # blur 반경
        self.shadow_opacity = 180      # 0~255
        self.align = 'center'
        self.pos_x_ratio = 0.5
        self.pos_y_ratio = 0.55


def render_name_on_image(pil_img, name, style: TextStyle):
    import math
    from PIL import ImageFilter
    img = pil_img.copy().convert("RGBA")
    iw, ih = img.size

    # 폰트 크기 결정
    if style.auto_size or style.font_size <= 0:
        size = auto_font_size(iw, ih, name, style.font_path)
    else:
        size = style.font_size

    try:
        if style.font_path and os.path.exists(style.font_path):
            font = ImageFont.truetype(style.font_path, size)
        else:
            font = ImageFont.load_default()
    except Exception:
        font = ImageFont.load_default()

    # 텍스트 위치
    dummy_draw = ImageDraw.Draw(img)
    bbox = dummy_draw.textbbox((0, 0), name, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    cx = int(iw * style.pos_x_ratio)
    cy = int(ih * style.pos_y_ratio)
    tx = cx - tw // 2 if style.align == 'center' else (cx - tw if style.align == 'right' else cx)
    ty = cy - th // 2

    # ── STEP 1: 외곽선 포함 텍스트 레이어 생성 (그림자 소스로 사용) ──
    ow = style.outline_width if style.outline else 0
    offsets = set()
    if ow > 0:
        steps = max(32, ow * 8)
        for i in range(steps):
            angle = 2 * math.pi * i / steps
            for r in range(1, ow + 1):
                offsets.add((int(round(math.cos(angle) * r)),
                             int(round(math.sin(angle) * r))))

    # 텍스트+외곽선 마스크 레이어 (그림자 형태 추출용)
    text_layer = Image.new("RGBA", (iw, ih), (0, 0, 0, 0))
    tl_draw = ImageDraw.Draw(text_layer)
    if offsets:
        oc = tuple(style.outline_color) + (255,)
        for dx, dy in offsets:
            tl_draw.text((tx + dx, ty + dy), name, font=font, fill=oc)
    tl_draw.text((tx, ty), name, font=font, fill=tuple(style.color) + (255,))

    # ── STEP 2: 그림자 = text_layer 전체를 오프셋+블러+투명도 적용 ──
    if style.shadow:
        ang = math.radians(style.shadow_angle)
        sdx = int(math.cos(ang) * style.shadow_distance)
        sdy = int(math.sin(ang) * style.shadow_distance)
        blur_r = max(1, style.shadow_blur)
        pad = blur_r * 3

        # 그림자용: text_layer를 단색(그림자 색)으로 채운 버전
        sh_mask = Image.new("RGBA", (iw + pad*2, ih + pad*2), (0, 0, 0, 0))
        sh_src = text_layer.copy()
        # 알파 채널만 뽑아서 그림자 색으로 칠하기
        r, g, b = style.shadow_color
        alpha = sh_src.split()[3]  # 원본 알파
        sh_colored = Image.new("RGBA", (iw, ih), (r, g, b, 0))
        # 투명도 조절
        scaled_alpha = alpha.point(lambda p: int(p * style.shadow_opacity / 255))
        sh_colored.putalpha(scaled_alpha)

        sh_mask.paste(sh_colored, (pad + sdx, pad + sdy))
        sh_mask = sh_mask.filter(ImageFilter.GaussianBlur(blur_r))
        shadow_crop = sh_mask.crop((pad, pad, pad + iw, pad + ih))
        img = Image.alpha_composite(img, shadow_crop)

    # ── STEP 3: 외곽선 + 본문 합성 ──
    img = Image.alpha_composite(img, text_layer)
    return img


# ═══════════════════════════════════════════════════════
# 렌더링 Worker (QThread - UI 블로킹 제거)
# ═══════════════════════════════════════════════════════

class RenderWorker(QObject):
    finished = pyqtSignal(object)   # QPixmap

    def __init__(self):
        super().__init__()
        self._task = None            # (pil_img, name, style)
        self._timer = QTimer()
        self._timer.setSingleShot(True)
        self._timer.setInterval(220)
        self._timer.timeout.connect(self._run)

    def request(self, pil_img, name, style):
        self._task = (pil_img, name, style)
        self._timer.start()

    def request_now(self, pil_img, name, style):
        self._timer.stop()
        self._task = (pil_img, name, style)
        self._run()

    def _run(self):
        if self._task is None:
            return
        pil_img, name, style = self._task
        try:
            rendered = render_name_on_image(pil_img, name, style)
            pm = pil_to_qpixmap(rendered)
            self.finished.emit(pm)
        except Exception:
            pass


# ═══════════════════════════════════════════════════════
# 스포이드 오버레이 (전체 화면 캡처 후 픽셀 선택)
# ═══════════════════════════════════════════════════════

class EyedropperOverlay(QWidget):
    color_picked = pyqtSignal(QColor)

    def __init__(self, parent=None):
        super().__init__(parent, Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._screen_pixmap = None
        self._cursor_pos = QPoint()

    def activate(self):
        screen = QApplication.primaryScreen()
        self._screen_pixmap = screen.grabWindow(0)
        geo = QApplication.desktop().screenGeometry()
        self.setGeometry(geo)
        self.setCursor(Qt.CrossCursor)
        self.showFullScreen()
        self.raise_()
        self.activateWindow()

    def mouseMoveEvent(self, e):
        self._cursor_pos = e.pos()
        self.update()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton and self._screen_pixmap:
            img = self._screen_pixmap.toImage()
            c = QColor(img.pixel(e.pos().x(), e.pos().y()))
            self.hide()
            self.color_picked.emit(c)
        elif e.button() == Qt.RightButton:
            self.hide()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape:
            self.hide()

    def paintEvent(self, e):
        if not self._screen_pixmap:
            return
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self._screen_pixmap)
        # 돋보기 원
        mx, my = self._cursor_pos.x(), self._cursor_pos.y()
        r = 60
        painter.setPen(QPen(QColor(255, 255, 255), 2))
        painter.drawEllipse(mx - r, my - r, r*2, r*2)
        painter.setPen(QPen(QColor(0, 0, 0), 1))
        painter.drawLine(mx - r, my, mx + r, my)
        painter.drawLine(mx, my - r, mx, my + r)
        # 현재 픽셀 색 미리보기
        img = self._screen_pixmap.toImage()
        if 0 <= mx < img.width() and 0 <= my < img.height():
            c = QColor(img.pixel(mx, my))
            painter.fillRect(mx + r + 4, my - 14, 28, 28, c)
            painter.setPen(QPen(QColor(80, 80, 80), 1))
            painter.drawRect(mx + r + 4, my - 14, 28, 28)


# ═══════════════════════════════════════════════════════
# 실시간 컬러피커 다이얼로그 (미리보기 포함)
# ═══════════════════════════════════════════════════════

class LiveColorDialog(QDialog):
    """색 변경 시 콜백으로 실시간 미리보기 업데이트"""
    color_changed = pyqtSignal(QColor)   # 실시간 변경 시그널

    def __init__(self, initial: QColor, on_change, parent=None):
        super().__init__(parent)
        self.setWindowTitle("색상 선택")
        self.setFixedSize(420, 120)
        self._color = initial
        self._on_change = on_change

        v = QVBoxLayout(self)
        v.setSpacing(10)

        row = QHBoxLayout()
        # 색상 버튼
        self.btn_pick = QPushButton("🎨  색 선택 (팔레트)")
        self.btn_pick.setMinimumHeight(38)
        self.btn_pick.clicked.connect(self._open_palette)
        row.addWidget(self.btn_pick)

        # 스포이드 버튼
        self.btn_eye = QPushButton("💉  스포이드")
        self.btn_eye.setMinimumHeight(38)
        self.btn_eye.clicked.connect(self._open_eyedropper)
        row.addWidget(self.btn_eye)

        # 현재 색 미리보기 박스
        self.lbl_preview = QLabel()
        self.lbl_preview.setFixedSize(60, 38)
        self.lbl_preview.setStyleSheet(f"background:{initial.name()}; border:2px solid #888; border-radius:4px;")
        row.addWidget(self.lbl_preview)
        v.addLayout(row)

        # HEX 입력
        hex_row = QHBoxLayout()
        hex_row.addWidget(QLabel("HEX:"))
        self.txt_hex = QLineEdit(initial.name())
        self.txt_hex.setMaxLength(7)
        self.txt_hex.setFixedWidth(90)
        self.txt_hex.textChanged.connect(self._on_hex)
        hex_row.addWidget(self.txt_hex)
        hex_row.addStretch()
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        hex_row.addWidget(bb)
        v.addLayout(hex_row)

        # 스포이드 오버레이
        self._eyedropper = EyedropperOverlay()
        self._eyedropper.color_picked.connect(self._on_eyedrop)

    def _open_palette(self):
        c = QColorDialog.getColor(self._color, self)
        if c.isValid():
            self._apply(c)

    def _open_eyedropper(self):
        self.hide()
        QTimer.singleShot(150, self._eyedropper.activate)

    def _on_eyedrop(self, c):
        self.show()
        self._apply(c)

    def _on_hex(self, text):
        if QColor.isValidColor(text):
            self._apply(QColor(text), update_hex=False)

    def _apply(self, c: QColor, update_hex=True):
        self._color = c
        self.lbl_preview.setStyleSheet(f"background:{c.name()}; border:2px solid #888; border-radius:4px;")
        if update_hex:
            self.txt_hex.blockSignals(True)
            self.txt_hex.setText(c.name())
            self.txt_hex.blockSignals(False)
        self._on_change(c)      # 실시간 미리보기 업데이트

    def get_color(self):
        return self._color




class DraggablePreview(QLabel):
    position_changed = pyqtSignal(float, float)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setCursor(QCursor(Qt.CrossCursor))
        self._dragging = False
        self._base_pixmap = None

    def set_base_pixmap(self, pm):
        self._base_pixmap = pm
        self.setPixmap(pm.scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._dragging = True
            self._update_pos(e.pos())

    def mouseMoveEvent(self, e):
        if self._dragging:
            self._update_pos(e.pos())

    def mouseReleaseEvent(self, e):
        self._dragging = False

    def _update_pos(self, qpos):
        if not self._base_pixmap:
            return
        lw, lh = self.width(), self.height()
        iw, ih = self._base_pixmap.width(), self._base_pixmap.height()
        scale = min(lw / iw, lh / ih)
        rw, rh = iw * scale, ih * scale
        ox, oy = (lw - rw) / 2, (lh - rh) / 2
        rx = max(0.0, min(1.0, (qpos.x() - ox) / rw))
        ry = max(0.0, min(1.0, (qpos.y() - oy) / rh))
        self.position_changed.emit(rx, ry)

    def resizeEvent(self, e):
        if self._base_pixmap:
            self.setPixmap(self._base_pixmap.scaled(
                self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))


# ═══════════════════════════════════════════════════════
# 폰트 미리보기 Delegate
# ═══════════════════════════════════════════════════════

class FontPreviewDelegate(QStyledItemDelegate):
    """드롭다운 각 항목을 해당 폰트로 렌더링"""
    def __init__(self, font_dict, parent=None):
        super().__init__(parent)
        self.font_dict = font_dict      # {이름: 경로}
        self._cache = {}                # {이름: QFont}

    def _get_qfont(self, name):
        if name in self._cache:
            return self._cache[name]
        path = self.font_dict.get(name)
        if path:
            fid = QFontDatabase.addApplicationFont(path)
            families = QFontDatabase.applicationFontFamilies(fid)
            if families:
                qf = QFont(families[0], 13)
                self._cache[name] = qf
                return qf
        qf = QFont()
        qf.setPointSize(13)
        self._cache[name] = qf
        return qf

    def sizeHint(self, option, index):
        return QSize(200, 32)

    def paint(self, painter, option, index):
        name = index.data(Qt.DisplayRole) or ''
        if option.state & QStyle.State_Selected:
            painter.fillRect(option.rect, QColor('#d0d8f8'))
        else:
            painter.fillRect(option.rect, QColor('#ffffff'))
        qf = self._get_qfont(name)
        painter.setFont(qf)
        painter.setPen(QColor('#222233'))
        painter.drawText(option.rect.adjusted(8, 0, -4, 0),
                         Qt.AlignVCenter | Qt.AlignLeft, name)


# ═══════════════════════════════════════════════════════
# 폰트 선택 콤보박스 (검색 가능 + 미리보기)
# ═══════════════════════════════════════════════════════

class FontComboBox(QComboBox):
    def __init__(self, font_dict, parent=None):
        super().__init__(parent)
        self.font_dict = font_dict
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)

        names = ['[ 자동 (시스템 기본) ]'] + list(font_dict.keys())
        self.addItems(names)

        # 미리보기 delegate
        self.setItemDelegate(FontPreviewDelegate(font_dict, self))

        completer = QCompleter(names, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.setCompleter(completer)
        self.setMinimumWidth(200)
        self.setMaxVisibleItems(12)

    def selected_font_path(self):
        name = self.currentText()
        return self.font_dict.get(name, None)


# ═══════════════════════════════════════════════════════
# 메인 윈도우
# ═══════════════════════════════════════════════════════

class NametagMaker(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("⭐ 별따오기 이름표 생성기")
        self.setMinimumSize(900, 620)

        self.group_images = {}
        self.group_names = {}
        self.style = TextStyle()
        self.preview_gnum = None
        self.preview_name_idx = 0
        self.system_fonts = {}

        # 렌더 워커 (비동기 렌더링)
        self._worker = RenderWorker()
        self._worker.finished.connect(self._on_render_done)

        # 폰트 스캔 (시작 시)
        self.system_fonts = scan_system_fonts()

        self._build_ui()
        self._apply_stylesheet()

    # ────────────────────────────────────────────────
    # UI 구성
    # ────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(14)

        # ── 왼쪽 설정 패널 ──────────────────────────
        left = QWidget()
        left.setFixedWidth(460)
        lv = QVBoxLayout(left)
        lv.setSpacing(10)

        # ▸ 이미지
        img_grp = QGroupBox("📁  조별 배경 이미지")
        ig = QVBoxLayout(img_grp)
        ig.setSpacing(6)

        self.img_list = QListWidget()
        self.img_list.setMaximumHeight(120)
        self.img_list.currentRowChanged.connect(self._on_image_select)
        ig.addWidget(self.img_list)

        br = QHBoxLayout()
        self.btn_add_img = QPushButton("＋ 이미지 추가")
        self.btn_add_img.clicked.connect(self._add_images)
        self.btn_clear_img = QPushButton("✕ 전체 제거")
        self.btn_clear_img.clicked.connect(self._clear_images)
        br.addWidget(self.btn_add_img)
        br.addWidget(self.btn_clear_img)
        ig.addLayout(br)

        hint = QLabel("💡 파일명에 '1조', '2조' 포함 시 자동 매핑")
        hint.setObjectName("hint")
        ig.addWidget(hint)
        lv.addWidget(img_grp)

        # ▸ 명단
        csv_grp = QGroupBox("📋  조별 명단  (CSV / 엑셀)")
        cg = QVBoxLayout(csv_grp)
        cg.setSpacing(6)
        self.btn_load_csv = QPushButton("📂  명단 파일 불러오기")
        self.btn_load_csv.clicked.connect(self._load_names)
        cg.addWidget(self.btn_load_csv)
        self.lbl_csv_status = QLabel("파일 미선택")
        self.lbl_csv_status.setObjectName("hint")
        self.lbl_csv_status.setWordWrap(True)
        cg.addWidget(self.lbl_csv_status)
        fmt_hint = QLabel("형식 예시 →  A열: 1조   B열: 홍길동")
        fmt_hint.setObjectName("hint2")
        cg.addWidget(fmt_hint)
        lv.addWidget(csv_grp)

        # ▸ 폰트
        font_grp = QGroupBox("🔤  폰트 설정")
        fg = QVBoxLayout(font_grp)
        fg.setSpacing(10)
        fg.setContentsMargins(12, 16, 12, 12)

        # ① 시스템 폰트 드롭다운 (전체 너비)
        lbl_sys = QLabel("설치된 폰트 선택  (이름으로 검색 가능):")
        lbl_sys.setObjectName("hint2")
        fg.addWidget(lbl_sys)
        self.cmb_font = FontComboBox(self.system_fonts)
        self.cmb_font.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.cmb_font.setMinimumHeight(34)
        self.cmb_font.currentIndexChanged.connect(self._on_font_combo)
        fg.addWidget(self.cmb_font)

        # 구분선
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("border: none; background: #dde; max-height: 1px;")
        fg.addWidget(line)

        # ② 파일 직접 선택 (한 줄)
        lbl_or = QLabel("또는 폰트 파일 직접 선택  (.ttf / .otf):")
        lbl_or.setObjectName("hint2")
        fg.addWidget(lbl_or)
        self.btn_font_file = QPushButton("📂  파일에서 폰트 선택")
        self.btn_font_file.setMinimumHeight(34)
        self.btn_font_file.clicked.connect(self._pick_font_file)
        fg.addWidget(self.btn_font_file)

        # ③ 웹폰트 URL (라벨 + 입력 + 버튼을 각각 한 줄씩)
        lbl_url = QLabel("또는 웹폰트 URL 입력  (눈누 등):")
        lbl_url.setObjectName("hint2")
        fg.addWidget(lbl_url)
        url_row = QHBoxLayout()
        url_row.setSpacing(6)
        self.txt_font_url = QLineEdit()
        self.txt_font_url.setPlaceholderText("https://... .ttf  URL 붙여넣기")
        self.txt_font_url.setMinimumHeight(34)
        url_row.addWidget(self.txt_font_url)
        self.btn_font_url = QPushButton("다운로드")
        self.btn_font_url.setFixedWidth(90)
        self.btn_font_url.setMinimumHeight(34)
        self.btn_font_url.clicked.connect(self._download_font)
        url_row.addWidget(self.btn_font_url)
        fg.addLayout(url_row)

        # 상태 표시
        self.lbl_font_status = QLabel("선택된 폰트: 시스템 기본")
        self.lbl_font_status.setObjectName("hint")
        self.lbl_font_status.setWordWrap(True)
        fg.addWidget(self.lbl_font_status)

        # 구분선
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setStyleSheet("border: none; background: #dde; max-height: 1px;")
        fg.addWidget(line2)

        # ④ 크기 설정 (한 줄에 체크박스 + 수동 입력)
        sz_row = QHBoxLayout()
        sz_row.setSpacing(12)
        self.chk_auto_size = QCheckBox("크기 자동  (이미지 비율 기준)")
        self.chk_auto_size.setChecked(True)
        self.chk_auto_size.stateChanged.connect(self._on_autosize_toggle)
        sz_row.addWidget(self.chk_auto_size)
        sz_row.addStretch()
        lbl_manual = QLabel("수동 크기:")
        sz_row.addWidget(lbl_manual)
        self.spn_size = QSpinBox()
        self.spn_size.setRange(10, 500)
        self.spn_size.setValue(120)
        self.spn_size.setEnabled(False)
        self.spn_size.setFixedWidth(75)
        self.spn_size.setMinimumHeight(30)
        self.spn_size.valueChanged.connect(self._on_style_change)
        sz_row.addWidget(self.spn_size)
        fg.addLayout(sz_row)

        lv.addWidget(font_grp)

        # ▸ 스타일
        style_grp = QGroupBox("🎨  텍스트 스타일")
        sg = QVBoxLayout(style_grp)
        sg.setSpacing(8)

        # 글자색
        cr = QHBoxLayout()
        cr.addWidget(QLabel("글자 색:"))
        self.btn_color = QPushButton()
        self._set_btn_color(self.btn_color, QColor(255, 220, 80))
        self.btn_color.clicked.connect(lambda: self._pick_color('color', self.btn_color))
        cr.addWidget(self.btn_color)
        cr.addStretch()
        sg.addLayout(cr)

        # 외곽선
        ol_row = QHBoxLayout()
        self.chk_outline = QCheckBox("외곽선")
        self.chk_outline.setChecked(True)
        self.chk_outline.stateChanged.connect(self._on_style_change)
        ol_row.addWidget(self.chk_outline)
        self.btn_ol_color = QPushButton()
        self._set_btn_color(self.btn_ol_color, QColor(255, 255, 255))
        self.btn_ol_color.clicked.connect(lambda: self._pick_color('outline_color', self.btn_ol_color))
        ol_row.addWidget(self.btn_ol_color)
        ol_row.addWidget(QLabel("두께:"))
        self.spn_ol = QSpinBox()
        self.spn_ol.setRange(0, 30)
        self.spn_ol.setValue(4)
        self.spn_ol.valueChanged.connect(self._on_style_change)
        ol_row.addWidget(self.spn_ol)
        ol_row.addStretch()
        sg.addLayout(ol_row)

        # 그림자
        sh_grp_row = QHBoxLayout()
        self.chk_shadow = QCheckBox("그림자")
        self.chk_shadow.setChecked(True)
        self.chk_shadow.stateChanged.connect(self._on_style_change)
        sh_grp_row.addWidget(self.chk_shadow)
        self.btn_sh_color = QPushButton()
        self._set_btn_color(self.btn_sh_color, QColor(0, 80, 160))
        self.btn_sh_color.clicked.connect(lambda: self._pick_color('shadow_color', self.btn_sh_color))
        sh_grp_row.addWidget(self.btn_sh_color)
        sh_grp_row.addStretch()
        sg.addLayout(sh_grp_row)

        # 그림자 세부 (2행 그리드)
        sh_detail1 = QHBoxLayout()
        sh_detail1.addWidget(QLabel("방향(°):"))
        self.spn_sh_angle = QSpinBox()
        self.spn_sh_angle.setRange(0, 359)
        self.spn_sh_angle.setValue(135)
        self.spn_sh_angle.setWrapping(True)
        self.spn_sh_angle.setFixedWidth(65)
        self.spn_sh_angle.valueChanged.connect(self._on_style_change)
        sh_detail1.addWidget(self.spn_sh_angle)
        sh_detail1.addSpacing(8)
        sh_detail1.addWidget(QLabel("거리:"))
        self.spn_sh_dist = QSpinBox()
        self.spn_sh_dist.setRange(0, 100)
        self.spn_sh_dist.setValue(12)
        self.spn_sh_dist.setFixedWidth(60)
        self.spn_sh_dist.valueChanged.connect(self._on_style_change)
        sh_detail1.addWidget(self.spn_sh_dist)
        sh_detail1.addStretch()
        sg.addLayout(sh_detail1)

        sh_detail2 = QHBoxLayout()
        sh_detail2.addWidget(QLabel("흐림:"))
        self.spn_sh_blur = QSpinBox()
        self.spn_sh_blur.setRange(0, 40)
        self.spn_sh_blur.setValue(6)
        self.spn_sh_blur.setFixedWidth(60)
        self.spn_sh_blur.valueChanged.connect(self._on_style_change)
        sh_detail2.addWidget(self.spn_sh_blur)
        sh_detail2.addSpacing(8)
        sh_detail2.addWidget(QLabel("투명도:"))
        self.spn_sh_opacity = QSpinBox()
        self.spn_sh_opacity.setRange(0, 255)
        self.spn_sh_opacity.setValue(180)
        self.spn_sh_opacity.setFixedWidth(60)
        self.spn_sh_opacity.valueChanged.connect(self._on_style_change)
        sh_detail2.addWidget(self.spn_sh_opacity)
        sh_detail2.addStretch()
        sg.addLayout(sh_detail2)

        # 정렬
        al_row = QHBoxLayout()
        al_row.addWidget(QLabel("정렬:"))
        self.cmb_align = QComboBox()
        self.cmb_align.addItems(['가운데', '왼쪽', '오른쪽'])
        self.cmb_align.currentIndexChanged.connect(self._on_style_change)
        al_row.addWidget(self.cmb_align)
        al_row.addStretch()
        sg.addLayout(al_row)
        lv.addWidget(style_grp)

        lv.addStretch()

        self.btn_generate = QPushButton("🖨️   이름표 전체 생성")
        self.btn_generate.setFixedHeight(54)
        self.btn_generate.clicked.connect(self._generate_all)
        self.btn_generate.setObjectName("generateBtn")
        lv.addWidget(self.btn_generate)

        root.addWidget(left)

        # 왼쪽 패널을 스크롤 영역으로 감싸기
        from PyQt5.QtWidgets import QScrollArea
        scroll = QScrollArea()
        scroll.setWidget(left)
        scroll.setWidgetResizable(True)
        scroll.setFixedWidth(480)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        root.addWidget(scroll)

        # ── 오른쪽 미리보기 ──────────────────────────
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setSpacing(8)

        ph = QHBoxLayout()
        lbl_title = QLabel("📌  미리보기   (클릭·드래그로 텍스트 위치 조정)")
        lbl_title.setObjectName("sectionTitle")
        ph.addWidget(lbl_title)
        ph.addStretch()
        self.lbl_preview_info = QLabel("—")
        self.lbl_preview_info.setObjectName("previewInfo")
        ph.addWidget(self.lbl_preview_info)
        self.btn_prev = QPushButton("◀")
        self.btn_prev.setFixedWidth(36)
        self.btn_prev.clicked.connect(self._prev_name)
        self.btn_next = QPushButton("▶")
        self.btn_next.setFixedWidth(36)
        self.btn_next.clicked.connect(self._next_name)
        ph.addWidget(self.btn_prev)
        ph.addWidget(self.btn_next)
        rv.addLayout(ph)

        self.preview = DraggablePreview()
        self.preview.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview.setStyleSheet("background: #e8e8e8; border-radius: 8px; border: 2px solid #bbb;")
        self.preview.position_changed.connect(self._on_pos_drag)
        rv.addWidget(self.preview)

        pos_row = QHBoxLayout()
        pos_row.addWidget(QLabel("X 위치:"))
        self.spn_px = QDoubleSpinBox()
        self.spn_px.setRange(0.0, 1.0)
        self.spn_px.setSingleStep(0.01)
        self.spn_px.setValue(0.5)
        self.spn_px.setDecimals(2)
        self.spn_px.valueChanged.connect(self._on_pos_spinbox)
        pos_row.addWidget(self.spn_px)
        pos_row.addSpacing(20)
        pos_row.addWidget(QLabel("Y 위치:"))
        self.spn_py = QDoubleSpinBox()
        self.spn_py.setRange(0.0, 1.0)
        self.spn_py.setSingleStep(0.01)
        self.spn_py.setValue(0.55)
        self.spn_py.setDecimals(2)
        self.spn_py.valueChanged.connect(self._on_pos_spinbox)
        pos_row.addWidget(self.spn_py)
        pos_row.addStretch()
        rv.addLayout(pos_row)

        root.addWidget(right)

    def _apply_stylesheet(self):
        # 밝은 테마
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background: #f4f6fb;
                color: #222233;
                font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
                font-size: 14px;
            }
            QGroupBox {
                background: #ffffff;
                border: 1.5px solid #d0d4e8;
                border-radius: 8px;
                margin-top: 10px;
                padding: 10px 10px 10px 10px;
                font-size: 15px;
                font-weight: bold;
                color: #3344aa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 4px;
            }
            QPushButton {
                background: #eef0fa;
                border: 1.5px solid #aab0dd;
                border-radius: 6px;
                padding: 6px 14px;
                font-size: 14px;
                color: #333366;
            }
            QPushButton:hover {
                background: #d8dcf8;
                border-color: #667acc;
            }
            QPushButton:pressed {
                background: #c4caf0;
            }
            QPushButton#generateBtn {
                background: #27ae60;
                border: 2px solid #1a8048;
                color: #ffffff;
                font-size: 17px;
                font-weight: bold;
                border-radius: 8px;
            }
            QPushButton#generateBtn:hover {
                background: #2ecc71;
            }
            QListWidget {
                background: #fafbff;
                border: 1.5px solid #ccd;
                border-radius: 6px;
                font-size: 14px;
            }
            QListWidget::item { padding: 4px; }
            QListWidget::item:selected { background: #d0d8f8; color: #222; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {
                background: #fafbff;
                border: 1.5px solid #bbc;
                border-radius: 5px;
                padding: 4px 8px;
                font-size: 14px;
            }
            QComboBox QAbstractItemView {
                background: white;
                selection-background-color: #d0d8f8;
                font-size: 14px;
            }
            QCheckBox { font-size: 14px; spacing: 6px; }
            QCheckBox::indicator {
                width: 18px; height: 18px;
                border-radius: 4px;
                border: 1.5px solid #99a;
                background: white;
            }
            QCheckBox::indicator:checked {
                background: #5566ee;
                border-color: #4455cc;
            }
            QLabel#hint { color: #778; font-size: 12px; }
            QLabel#hint2 { color: #559; font-size: 12px; }
            QLabel#sectionTitle { font-size: 15px; font-weight: bold; color: #334; }
            QLabel#previewInfo { font-size: 14px; color: #446; font-weight: bold; }
            QScrollBar:vertical { background: #eee; width: 10px; border-radius: 5px; }
            QScrollBar::handle:vertical { background: #aab; border-radius: 5px; }
        """)

    # ────────────────────────────────────────────────
    # 헬퍼
    # ────────────────────────────────────────────────

    def _set_btn_color(self, btn, qcolor):
        btn.setFixedSize(36, 28)
        btn.setStyleSheet(f"background:{qcolor.name()}; border:1.5px solid #888; border-radius:4px;")
        btn._color = [qcolor.red(), qcolor.green(), qcolor.blue()]

    def _pick_color(self, attr, btn):
        old = QColor(*getattr(self.style, attr))

        def on_live_change(c: QColor):
            setattr(self.style, attr, [c.red(), c.green(), c.blue()])
            self._set_btn_color(btn, c)
            self._refresh_preview()

        dlg = LiveColorDialog(old, on_live_change, self)
        if dlg.exec_() == QDialog.Accepted:
            c = dlg.get_color()
            setattr(self.style, attr, [c.red(), c.green(), c.blue()])
            self._set_btn_color(btn, c)
        else:
            # 취소 시 원래 색으로 복구
            setattr(self.style, attr, [old.red(), old.green(), old.blue()])
            self._set_btn_color(btn, old)
        self._refresh_preview()

    def _sync_style(self):
        self.style.auto_size = self.chk_auto_size.isChecked()
        self.style.font_size = 0 if self.style.auto_size else self.spn_size.value()
        self.style.outline = self.chk_outline.isChecked()
        self.style.outline_width = self.spn_ol.value()
        self.style.shadow = self.chk_shadow.isChecked()
        self.style.shadow_angle = self.spn_sh_angle.value()
        self.style.shadow_distance = self.spn_sh_dist.value()
        self.style.shadow_blur = self.spn_sh_blur.value()
        self.style.shadow_opacity = self.spn_sh_opacity.value()
        self.style.align = {0: 'center', 1: 'left', 2: 'right'}[self.cmb_align.currentIndex()]

    # ────────────────────────────────────────────────
    # 이벤트
    # ────────────────────────────────────────────────

    def _on_autosize_toggle(self, state):
        self.spn_size.setEnabled(not self.chk_auto_size.isChecked())
        self._refresh_preview()

    def _on_font_combo(self):
        path = self.cmb_font.selected_font_path()
        self.style.font_path = path
        name = self.cmb_font.currentText()
        self.lbl_font_status.setText(f"선택된 폰트: {name}")
        self._refresh_preview()

    def _add_images(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "이미지 파일 선택", "",
            "이미지 파일 (*.png *.jpg *.jpeg *.webp *.bmp)")
        for f in files:
            gnum = detect_group_number(f) or (len(self.group_images) + 1)
            try:
                img = Image.open(f).convert("RGBA")
                self.group_images[gnum] = img
                self.img_list.addItem(f"  {gnum}조  ←  {Path(f).name}")
                if self.preview_gnum is None:
                    self.preview_gnum = gnum
            except Exception as e:
                QMessageBox.warning(self, "오류", f"이미지 로드 실패:\n{e}")
        self._refresh_preview()

    def _clear_images(self):
        self.group_images.clear()
        self.img_list.clear()
        self.preview_gnum = None
        self.preview.clear()

    def _on_image_select(self, row):
        if row < 0:
            return
        gnums = sorted(self.group_images.keys())
        if row < len(gnums):
            self.preview_gnum = gnums[row]
            self.preview_name_idx = 0
            self._refresh_preview()

    def _load_names(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "명단 파일 선택", "", "CSV/엑셀 (*.csv *.xlsx *.xls)")
        if not f:
            return
        try:
            self.group_names = load_names_from_file(f)
            summary = ", ".join(f"{g}조 {len(n)}명" for g, n in sorted(self.group_names.items()))
            self.lbl_csv_status.setText(f"✅ {Path(f).name}\n{summary}")
            self.lbl_csv_status.setStyleSheet("color: #228833; font-size: 12px;")
        except Exception as e:
            QMessageBox.warning(self, "오류", f"명단 로드 실패:\n{e}")
        self._refresh_preview()

    def _pick_font_file(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "폰트 파일 선택", "", "폰트 파일 (*.ttf *.otf)")
        if f:
            self.style.font_path = f
            self.lbl_font_status.setText(f"선택된 폰트: {Path(f).name}")
            self.lbl_font_status.setStyleSheet("color: #226633; font-size: 12px;")
            self._refresh_preview()

    def _download_font(self):
        url = self.txt_font_url.text().strip()
        if not url:
            return
        self.lbl_font_status.setText("⏳ 다운로드 중...")
        QApplication.processEvents()
        path = download_font(url)
        if path:
            self.style.font_path = path
            self.lbl_font_status.setText("✅ 웹폰트 다운로드 완료")
            self.lbl_font_status.setStyleSheet("color: #226633; font-size: 12px;")
        else:
            self.lbl_font_status.setText("❌ 다운로드 실패 — URL을 확인해주세요")
            self.lbl_font_status.setStyleSheet("color: #cc2222; font-size: 12px;")
        self._refresh_preview()

    def _on_style_change(self):
        # 연속 변경 시 디바운스: Worker 내부 타이머가 220ms 후 실행
        self._sync_style()
        gnum = self.preview_gnum
        if gnum is None or gnum not in self.group_images:
            return
        img = self.group_images[gnum]
        names = self.group_names.get(gnum, ["홍길동"])
        name = names[min(self.preview_name_idx, len(names)-1)]
        self._worker.request(img, name, self.style)

    # ────────────────────────────────────────────────
    # 미리보기 렌더
    # ────────────────────────────────────────────────

    def _refresh_preview(self):
        """즉시 렌더 (폰트 선택, 드래그 등 단발성 이벤트)"""
        self._sync_style()
        gnum = self.preview_gnum
        if gnum is None or gnum not in self.group_images:
            return
        img = self.group_images[gnum]
        names = self.group_names.get(gnum, ["홍길동"])
        idx = min(self.preview_name_idx, len(names) - 1)
        name = names[idx]
        self.lbl_preview_info.setText(f"{gnum}조  {idx+1}/{len(names)}  〔{name}〕")
        self._worker.request_now(img, name, self.style)

    def _on_render_done(self, pm):
        self.preview.set_base_pixmap(pm)

    def _on_pos_drag(self, rx, ry):
        self.style.pos_x_ratio = rx
        self.style.pos_y_ratio = ry
        self.spn_px.blockSignals(True)
        self.spn_py.blockSignals(True)
        self.spn_px.setValue(rx)
        self.spn_py.setValue(ry)
        self.spn_px.blockSignals(False)
        self.spn_py.blockSignals(False)
        self._refresh_preview()

    def _on_pos_spinbox(self):
        self.style.pos_x_ratio = self.spn_px.value()
        self.style.pos_y_ratio = self.spn_py.value()
        self._refresh_preview()

    def _prev_name(self):
        gnum = self.preview_gnum
        if gnum and gnum in self.group_names:
            names = self.group_names[gnum]
            self.preview_name_idx = (self.preview_name_idx - 1) % len(names)
            self._refresh_preview()

    def _next_name(self):
        gnum = self.preview_gnum
        if gnum and gnum in self.group_names:
            names = self.group_names[gnum]
            self.preview_name_idx = (self.preview_name_idx + 1) % len(names)
            self._refresh_preview()

    # ────────────────────────────────────────────────
    # 전체 생성
    # ────────────────────────────────────────────────

    def _generate_all(self):
        if not self.group_images:
            QMessageBox.warning(self, "오류", "배경 이미지를 먼저 추가해주세요.")
            return
        if not self.group_names:
            QMessageBox.warning(self, "오류", "명단 파일을 먼저 불러와주세요.")
            return
        out_dir = QFileDialog.getExistingDirectory(self, "저장 폴더 선택")
        if not out_dir:
            return

        fmt_dlg = FormatDialog(self)
        if fmt_dlg.exec_() != QDialog.Accepted:
            return
        save_png = fmt_dlg.chk_png.isChecked()
        save_pdf = fmt_dlg.chk_pdf.isChecked()

        self._sync_style()
        total = sum(len(v) for v in self.group_names.values())
        prog = QProgressDialog("이름표 생성 중...", "취소", 0, total, self)
        prog.setWindowTitle("생성 중")
        prog.setWindowModality(Qt.WindowModal)
        prog.show()

        done = 0
        pdf_images = []
        available = sorted(self.group_images.keys())

        for gnum, names in sorted(self.group_names.items()):
            base_img = self.group_images.get(gnum, self.group_images[available[0]])
            for name in names:
                if prog.wasCanceled():
                    break
                rendered = render_name_on_image(base_img, name, self.style)
                safe = re.sub(r'[\\/:*?"<>|]', '_', name)
                filename = f"{gnum}조_{safe}"
                if save_png:
                    rendered.save(os.path.join(out_dir, filename + ".png"))
                if save_pdf:
                    pdf_images.append((rendered, f"{gnum}조 {name}"))
                done += 1
                prog.setValue(done)
                QApplication.processEvents()

        if save_pdf and pdf_images:
            self._save_pdf(pdf_images, os.path.join(out_dir, "이름표_전체.pdf"))

        prog.close()
        QMessageBox.information(self, "완료", f"✅ {done}장 생성 완료!\n\n저장 위치:\n{out_dir}")

    def _save_pdf(self, images_labels, path):
        """A4 한 장에 이름표 4개 (2×2) 배치 — 인쇄 시 4장"""
        c = rl_canvas.Canvas(path, pagesize=A4)
        pw, ph = A4          # 595 × 842 pt
        margin = 18          # 페이지 여백
        gap = 8              # 칸 사이 간격

        cell_w = (pw - margin * 2 - gap) / 2
        cell_h = (ph - margin * 2 - gap) / 2

        # 4개씩 묶어서 페이지 생성
        for page_start in range(0, len(images_labels), 4):
            batch = images_labels[page_start:page_start + 4]
            # 배치 좌표: [좌상, 우상, 좌하, 우하]  (reportlab: y=0이 아래)
            positions = [
                (margin,              margin + gap + cell_h),   # 좌상
                (margin + gap + cell_w, margin + gap + cell_h), # 우상
                (margin,              margin),                  # 좌하
                (margin + gap + cell_w, margin),                # 우하
            ]
            for (img, _), (cx, cy) in zip(batch, positions):
                buf = io.BytesIO()
                img.save(buf, format='PNG')
                buf.seek(0)
                ir = ImageReader(buf)
                iw, ih = img.size
                scale = min(cell_w / iw, cell_h / ih)
                rw, rh = iw * scale, ih * scale
                # 셀 내 가운데 정렬
                ox = cx + (cell_w - rw) / 2
                oy = cy + (cell_h - rh) / 2
                c.drawImage(ir, ox, oy, rw, rh, mask='auto')
            # 칸 구분선 (옅은 점선)
            c.setStrokeColorRGB(0.75, 0.75, 0.75)
            c.setDash(4, 4)
            c.setLineWidth(0.5)
            mid_x = margin + cell_w + gap / 2
            mid_y = margin + cell_h + gap / 2
            c.line(mid_x, margin, mid_x, ph - margin)
            c.line(margin, mid_y, pw - margin, mid_y)
            c.setDash()
            c.showPage()
        c.save()


# ═══════════════════════════════════════════════════════
# 업데이트 체커 (백그라운드 스레드)
# ═══════════════════════════════════════════════════════

class UpdateChecker(QObject):
    update_available = pyqtSignal(str, str)   # (latest_version, changelog)
    no_update        = pyqtSignal()
    check_failed     = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._thread = QThread()
        self.moveToThread(self._thread)
        self._thread.started.connect(self._run)

    def start(self):
        self._thread.start()

    def _run(self):
        try:
            r = requests.get(VERSION_URL, timeout=6)
            r.raise_for_status()
            data = r.json()
            latest  = data.get("version", "0.0.0")
            changelog = data.get("changelog", "업데이트 내용 없음")
            if self._is_newer(latest, APP_VERSION):
                self.update_available.emit(latest, changelog)
            else:
                self.no_update.emit()
        except Exception:
            self.check_failed.emit()
        finally:
            self._thread.quit()

    @staticmethod
    def _is_newer(a: str, b: str) -> bool:
        """a > b 이면 True"""
        def parts(v): return [int(x) for x in v.split(".")]
        try:
            return parts(a) > parts(b)
        except Exception:
            return False


# ═══════════════════════════════════════════════════════
# 업데이트 다이얼로그 (패치노트 + 다운로드)
# ═══════════════════════════════════════════════════════

class UpdateDialog(QDialog):
    def __init__(self, latest_ver: str, changelog: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🎉  업데이트 알림")
        self.setFixedSize(520, 380)
        self._latest_ver = latest_ver

        v = QVBoxLayout(self)
        v.setSpacing(12)
        v.setContentsMargins(20, 20, 20, 16)

        # 제목
        lbl_title = QLabel(f"새 버전  <b>v{latest_ver}</b>  이 출시됐어요!")
        lbl_title.setStyleSheet("font-size:17px; color:#2255cc;")
        v.addWidget(lbl_title)

        lbl_cur = QLabel(f"현재 버전: v{APP_VERSION}")
        lbl_cur.setStyleSheet("color:#888; font-size:12px;")
        v.addWidget(lbl_cur)

        # 패치노트
        v.addWidget(QLabel("📋  업데이트 내용:"))
        from PyQt5.QtWidgets import QTextEdit
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMarkdown(changelog)
        self.txt_log.setStyleSheet(
            "background:#f8f9ff; border:1px solid #ccd; border-radius:6px; font-size:13px;")
        v.addWidget(self.txt_log)

        # 진행바
        from PyQt5.QtWidgets import QProgressBar
        self.prog = QProgressBar()
        self.prog.setVisible(False)
        self.prog.setRange(0, 100)
        self.prog.setStyleSheet(
            "QProgressBar{border:1px solid #bbc;border-radius:5px;background:#eef;}"
            "QProgressBar::chunk{background:#5566ee;border-radius:5px;}")
        v.addWidget(self.prog)

        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("color:#557; font-size:12px;")
        v.addWidget(self.lbl_status)

        # 버튼
        btn_row = QHBoxLayout()
        self.btn_update = QPushButton("⬇️  지금 업데이트")
        self.btn_update.setMinimumHeight(40)
        self.btn_update.setStyleSheet(
            "background:#2ecc71;border:2px solid #27ae60;color:white;"
            "font-size:15px;font-weight:bold;border-radius:7px;")
        self.btn_update.clicked.connect(self._start_download)
        btn_row.addWidget(self.btn_update)

        btn_skip = QPushButton("나중에")
        btn_skip.setMinimumHeight(40)
        btn_skip.clicked.connect(self.reject)
        btn_row.addWidget(btn_skip)
        v.addLayout(btn_row)

    def _start_download(self):
        self.btn_update.setEnabled(False)
        self.prog.setVisible(True)
        self.lbl_status.setText("다운로드 중...")

        self._dl_thread = QThread()
        self._dl_worker = _DownloadWorker(DOWNLOAD_URL)
        self._dl_worker.moveToThread(self._dl_thread)
        self._dl_thread.started.connect(self._dl_worker.run)
        self._dl_worker.progress.connect(self.prog.setValue)
        self._dl_worker.finished.connect(self._on_downloaded)
        self._dl_worker.failed.connect(self._on_fail)
        self._dl_thread.start()

    def _on_downloaded(self, tmp_path: str):
        self._dl_thread.quit()
        self.lbl_status.setText("✅  다운로드 완료! 재시작 중...")
        QTimer.singleShot(800, lambda: self._replace_and_restart(tmp_path))

    def _on_fail(self, msg: str):
        self._dl_thread.quit()
        self.lbl_status.setText(f"❌  실패: {msg}")
        self.btn_update.setEnabled(True)

    def _replace_and_restart(self, tmp_path: str):
        """현재 exe를 새 파일로 교체 후 재시작"""
        current = sys.executable
        backup  = current + ".bak"
        try:
            # 현재 실행 중인 exe는 덮어쓸 수 없으므로 배치 스크립트로 지연 교체
            if sys.platform == "win32":
                bat = tempfile.NamedTemporaryFile(delete=False, suffix=".bat",
                                                  mode="w", encoding="cp949")
                bat.write(f"""@echo off
timeout /t 2 /nobreak >nul
move /y "{tmp_path}" "{current}"
start "" "{current}"
del "%~f0"
""")
                bat.close()
                subprocess.Popen(["cmd", "/c", bat.name],
                                 creationflags=subprocess.CREATE_NO_WINDOW)
            else:
                shutil.move(tmp_path, current)
                os.chmod(current, 0o755)
                subprocess.Popen([current])
        except Exception as e:
            QMessageBox.warning(self, "오류", f"교체 실패:\n{e}\n\n수동으로 교체해주세요:\n{tmp_path}")
            return
        QApplication.quit()


class _DownloadWorker(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)   # tmp file path
    failed   = pyqtSignal(str)

    def __init__(self, url):
        super().__init__()
        self._url = url

    def run(self):
        try:
            r = requests.get(self._url, stream=True, timeout=30)
            r.raise_for_status()
            total = int(r.headers.get("content-length", 0))
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".exe")
            downloaded = 0
            for chunk in r.iter_content(chunk_size=65536):
                tmp.write(chunk)
                downloaded += len(chunk)
                if total:
                    self.progress.emit(int(downloaded / total * 100))
            tmp.close()
            self.finished.emit(tmp.name)
        except Exception as e:
            self.failed.emit(str(e))




class FormatDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("저장 형식 선택")
        self.setFixedSize(320, 180)
        v = QVBoxLayout(self)
        v.setSpacing(12)
        v.addWidget(QLabel("어떤 형식으로 저장할까요?"))
        self.chk_png = QCheckBox("PNG 이미지  (조별 개별 파일)")
        self.chk_png.setChecked(True)
        self.chk_pdf = QCheckBox("PDF  (전체 이름표 한 파일)")
        self.chk_pdf.setChecked(True)
        v.addWidget(self.chk_png)
        v.addWidget(self.chk_pdf)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)


# ═══════════════════════════════════════════════════════
# 진입점
# ═══════════════════════════════════════════════════════

def main():
    # HiDPI 지원
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("별따오기 이름표 생성기")

    f = app.font()
    f.setPointSize(14)
    app.setFont(f)

    win = NametagMaker()
    win.resize(1180, 720)
    win.show()

    # ── 백그라운드 업데이트 체크 ──────────────────────
    checker = UpdateChecker()

    def _on_update(latest_ver, changelog):
        dlg = UpdateDialog(latest_ver, changelog, win)
        dlg.exec_()

    checker.update_available.connect(_on_update)
    # 창이 완전히 뜬 뒤 1.5초 후에 체크 시작 (첫 화면 렌더 방해 안 하도록)
    QTimer.singleShot(1500, checker.start)

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
