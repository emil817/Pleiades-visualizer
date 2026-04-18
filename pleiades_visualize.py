import csv
import io
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET

import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QKeySequence
from PyQt5.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


plt.rcParams["font.family"] = "DejaVu Sans"

XLSX_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def parse_table_text(text):
    text = text.replace("\r\n", "\n").strip("\n")
    if not text.strip():
        return []

    if "\t" in text:
        return [line.split("\t") for line in text.split("\n")]

    sample = "\n".join(text.split("\n")[:5])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,")
        reader = csv.reader(io.StringIO(text), dialect)
        return [row for row in reader]
    except csv.Error:
        return [[line] for line in text.split("\n")]


def normalize_table_data(rows):
    normalized = []
    for row in rows:
        normalized.append(
            [str(cell).strip() if cell is not None else "" for cell in row]
        )

    while normalized and not any(normalized[-1]):
        normalized.pop()

    if not normalized:
        return []

    max_cols = max(len(row) for row in normalized)
    while max_cols > 0 and all(
        len(row) <= max_cols - 1 or not row[max_cols - 1] for row in normalized
    ):
        max_cols -= 1

    if max_cols <= 0:
        return []

    return [row[:max_cols] + [""] * (max_cols - len(row)) for row in normalized]


def read_csv_table(file_path):
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "windows-1251"):
        try:
            with open(file_path, "r", encoding=encoding, newline="") as source:
                return normalize_table_data(parse_table_text(source.read()))
        except UnicodeDecodeError:
            continue

    raise ValueError("Не удалось прочитать CSV-файл: неизвестная кодировка.")


def cell_reference_to_index(cell_reference):
    letters = "".join(char for char in cell_reference if char.isalpha()).upper()
    if not letters:
        return 0

    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def extract_xlsx_text(text_parent):
    if text_parent is None:
        return ""

    return "".join(
        node.text or "" for node in text_parent.iter() if node.tag.endswith("}t")
    )


def get_first_sheet_path(archive):
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    workbook_rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    first_sheet = workbook_root.find("main:sheets/main:sheet", XLSX_NS)
    if first_sheet is None:
        raise ValueError("В XLSX-файле не найден ни один лист.")

    relation_id = first_sheet.attrib.get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    )
    if not relation_id:
        raise ValueError("Не удалось определить первый лист XLSX-файла.")

    for relation in workbook_rels_root.findall("pkgrel:Relationship", XLSX_NS):
        if relation.attrib.get("Id") != relation_id:
            continue

        target = relation.attrib.get("Target", "")
        if not target:
            break

        normalized = posixpath.normpath(posixpath.join("xl", target))
        return normalized

    raise ValueError("Не удалось найти XML первого листа в XLSX-файле.")


def read_xlsx_shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    shared_strings_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        extract_xlsx_text(string_item)
        for string_item in shared_strings_root.findall("main:si", XLSX_NS)
    ]


def read_xlsx_table(file_path):
    try:
        with zipfile.ZipFile(file_path) as archive:
            shared_strings = read_xlsx_shared_strings(archive)
            sheet_path = get_first_sheet_path(archive)
            sheet_root = ET.fromstring(archive.read(sheet_path))
    except KeyError as exc:
        raise ValueError("XLSX-файл имеет неподдерживаемую структуру.") from exc
    except zipfile.BadZipFile as exc:
        raise ValueError("Файл не является корректным XLSX-архивом.") from exc

    rows = []
    for row_element in sheet_root.findall(".//main:sheetData/main:row", XLSX_NS):
        row_index = max(int(row_element.attrib.get("r", "1")) - 1, 0)
        while len(rows) <= row_index:
            rows.append([])

        row_values = rows[row_index]
        for cell_element in row_element.findall("main:c", XLSX_NS):
            col_index = cell_reference_to_index(cell_element.attrib.get("r", "A1"))
            while len(row_values) <= col_index:
                row_values.append("")

            cell_type = cell_element.attrib.get("t")
            value_node = cell_element.find("main:v", XLSX_NS)

            if cell_type == "inlineStr":
                value = extract_xlsx_text(cell_element.find("main:is", XLSX_NS))
            elif cell_type == "s":
                shared_index = int(value_node.text) if value_node is not None else -1
                if 0 <= shared_index < len(shared_strings):
                    value = shared_strings[shared_index]
                else:
                    value = ""
            elif cell_type == "b":
                value = (
                    "TRUE"
                    if value_node is not None and value_node.text == "1"
                    else "FALSE"
                )
            else:
                value = value_node.text if value_node is not None else ""
                if not value:
                    value = extract_xlsx_text(cell_element.find("main:is", XLSX_NS))

            row_values[col_index] = value

    return normalize_table_data(rows)


def read_table_file(file_path):
    lowered = file_path.lower()
    if lowered.endswith(".csv"):
        rows = read_csv_table(file_path)
    elif lowered.endswith(".xlsx"):
        rows = read_xlsx_table(file_path)
    else:
        raise ValueError("Поддерживаются только файлы CSV и XLSX.")

    if not rows:
        raise ValueError("Файл не содержит данных.")

    return rows


def draw_final_pleiad(left_params, right_params, edges):
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.set_xlim(-2.0, 4.0)

    shared_node_names = set(left_params) & set(right_params)
    left_only = [name for name in left_params if name not in shared_node_names]
    shared_nodes = [name for name in left_params if name in shared_node_names]
    right_only = [name for name in right_params if name not in shared_node_names]

    max_nodes = max(len(left_only), len(shared_nodes), len(right_only), 1)
    y_start = max_nodes * 2
    ax.set_ylim(-1.5, y_start + 1)

    node_artists = {}
    edge_artists = []
    top_y = y_start
    bottom_y = y_start - max(max_nodes - 1, 0) * 2
    default_node_face = "#d9ecff"
    active_node_face = "#bfdbfe"
    default_node_edge = "#5b6c7a"
    active_node_edge = "#2563eb"
    active_node_name = None
    drag_offset = (0.0, 0.0)

    def draw_node(text, x, y):
        box_style = dict(
            boxstyle="round,pad=0.8",
            facecolor=default_node_face,
            edgecolor=default_node_edge,
            linewidth=1.6,
        )
        return ax.text(
            x,
            y,
            text,
            ha="center",
            va="center",
            fontsize=12,
            fontweight="medium",
            bbox=box_style,
            zorder=3,
            multialignment="center",
        )

    def calculate_y_positions(items):
        if not items:
            return []
        if len(items) == 1:
            return [(top_y + bottom_y) / 2]

        step = (top_y - bottom_y) / (len(items) - 1)
        return [top_y - index * step for index in range(len(items))]

    def relayout_edge_labels():
        if not edge_artists:
            return

        positions = [list(edge_data["base_label_position"]) for edge_data in edge_artists]
        min_dx = 0.28
        min_dy = 0.52

        for _ in range(18):
            moved = False
            for left_index in range(len(positions)):
                for right_index in range(left_index + 1, len(positions)):
                    dx = positions[right_index][0] - positions[left_index][0]
                    dy = positions[right_index][1] - positions[left_index][1]
                    if abs(dx) >= min_dx or abs(dy) >= min_dy:
                        continue

                    overlap_x = min_dx - abs(dx)
                    overlap_y = min_dy - abs(dy)
                    shift_x = (overlap_x / 2 + 0.02) * 0.25
                    shift_y = overlap_y / 2 + 0.03

                    if dx >= 0:
                        positions[left_index][0] -= shift_x
                        positions[right_index][0] += shift_x
                    else:
                        positions[left_index][0] += shift_x
                        positions[right_index][0] -= shift_x

                    if dy >= 0:
                        positions[left_index][1] -= shift_y
                        positions[right_index][1] += shift_y
                    else:
                        positions[left_index][1] += shift_y
                        positions[right_index][1] -= shift_y

                    moved = True

            if not moved:
                break

        for edge_data, position in zip(edge_artists, positions):
            edge_data["label"].set_position(tuple(position))

    def update_edge_geometry(edge_data):
        start_x, start_y = node_artists[edge_data["start"]].get_position()
        end_x, end_y = node_artists[edge_data["end"]].get_position()
        edge_data["line"].set_data([start_x, end_x], [start_y, end_y])
        edge_data["base_label_position"] = (
            (start_x + end_x) * 0.5,
            (start_y + end_y) * 0.5,
        )

    def update_edges_for_node(node_name):
        for edge_data in edge_artists:
            if edge_data["start"] == node_name or edge_data["end"] == node_name:
                update_edge_geometry(edge_data)
        relayout_edge_labels()

    def set_node_active_state(node_name, is_active):
        node_patch = node_artists[node_name].get_bbox_patch()
        if node_patch is None:
            return

        node_patch.set_facecolor(active_node_face if is_active else default_node_face)
        node_patch.set_edgecolor(active_node_edge if is_active else default_node_edge)
        node_patch.set_linewidth(2.2 if is_active else 1.6)

    def find_node_under_cursor(event):
        for node_name, node_artist in reversed(list(node_artists.items())):
            contains, _ = node_artist.contains(event)
            if contains:
                return node_name
        return None

    def on_press(event):
        nonlocal active_node_name, drag_offset

        if (
            event.button != 1
            or event.inaxes != ax
            or event.xdata is None
            or event.ydata is None
        ):
            return

        node_name = find_node_under_cursor(event)
        if node_name is None:
            return

        active_node_name = node_name
        node_x, node_y = node_artists[node_name].get_position()
        drag_offset = (node_x - event.xdata, node_y - event.ydata)
        set_node_active_state(node_name, True)
        fig.canvas.draw_idle()

    def on_motion(event):
        if active_node_name is None or event.inaxes != ax:
            return
        if event.xdata is None or event.ydata is None:
            return

        node_artists[active_node_name].set_position(
            (event.xdata + drag_offset[0], event.ydata + drag_offset[1])
        )
        update_edges_for_node(active_node_name)
        fig.canvas.draw_idle()

    def on_release(event):
        nonlocal active_node_name

        if active_node_name is None:
            return

        set_node_active_state(active_node_name, False)
        active_node_name = None
        fig.canvas.draw_idle()

    for node_name, y_pos in zip(left_only, calculate_y_positions(left_only)):
        node_artists[node_name] = draw_node(node_name, 0, y_pos)

    for node_name, y_pos in zip(shared_nodes, calculate_y_positions(shared_nodes)):
        node_artists[node_name] = draw_node(node_name, 1, y_pos)

    for node_name, y_pos in zip(right_only, calculate_y_positions(right_only)):
        node_artists[node_name] = draw_node(node_name, 2, y_pos)

    for start_node, end_node, weight in edges:
        x1, y1 = node_artists[start_node].get_position()
        x2, y2 = node_artists[end_node].get_position()

        style = "dashed" if weight < 0 else "solid"
        width = 2 + abs(weight) * 3

        line_artist, = ax.plot(
            [x1, x2],
            [y1, y2],
            color="#334155",
            linestyle=style,
            linewidth=width,
            zorder=1,
        )

        label_artist = ax.text(
            (x1 + x2) * 0.5,
            (y1 + y2) * 0.5,
            f"{weight:g}",
            fontsize=11,
            fontweight="bold",
            bbox=dict(facecolor="white", edgecolor="none", alpha=0.92, pad=2),
            ha="center",
            va="center",
            zorder=4,
        )
        edge_data = {
            "start": start_node,
            "end": end_node,
            "line": line_artist,
            "label": label_artist,
            "base_label_position": ((x1 + x2) * 0.5, (y1 + y2) * 0.5),
        }
        edge_artists.append(edge_data)

    relayout_edge_labels()

    legend_elements = [
        Line2D(
            [0],
            [0],
            color="#334155",
            lw=2,
            linestyle="solid",
            label="Положительная корреляция",
        ),
        Line2D(
            [0],
            [0],
            color="#334155",
            lw=2,
            linestyle="dashed",
            label="Отрицательная корреляция",
        ),
    ]

    ax.legend(
        handles=legend_elements,
        loc="lower center",
        bbox_to_anchor=(0.5, 0.02),
        ncol=1,
        fontsize=12,
        frameon=True,
        facecolor="white",
        edgecolor="#94a3b8",
    )

    manager = getattr(fig.canvas, "manager", None)
    if manager is not None and hasattr(manager, "set_window_title"):
        manager.set_window_title("Корреляционная плеяда - ЛКМ: перетащить узел")

    plt.title("Корреляционная плеяда", fontsize=18, fontweight="bold", pad=20)
    plt.axis("off")
    plt.tight_layout()
    fig.canvas.draw()
    fig.canvas.mpl_connect("button_press_event", on_press)
    fig.canvas.mpl_connect("motion_notify_event", on_motion)
    fig.canvas.mpl_connect("button_release_event", on_release)
    plt.show()


class PasteTableWidget(QTableWidget):
    dimensionsChanged = pyqtSignal(int, int)

    def __init__(
        self,
        rows,
        cols,
        *,
        highlight_first_row=True,
        highlight_first_col=True,
        show_horizontal_header=False,
        horizontal_labels=None,
    ):
        super().__init__(rows, cols)
        self.highlight_first_row = highlight_first_row
        self.highlight_first_col = highlight_first_col
        self.data_row_offset = 1 if highlight_first_row else 0
        self.data_col_offset = 1 if highlight_first_col else 0

        self.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setAlternatingRowColors(False)
        self.setCornerButtonEnabled(False)

        self.horizontalHeader().setVisible(show_horizontal_header)
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setDefaultSectionSize(170)
        self.verticalHeader().setDefaultSectionSize(40)

        if horizontal_labels:
            self.setColumnCount(len(horizontal_labels))
            self.setHorizontalHeaderLabels(horizontal_labels)

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Paste):
            self.paste_from_clipboard()
            return
        super().keyPressEvent(event)

    def parse_clipboard_text(self, text):
        return parse_table_text(text)

    def paste_from_clipboard(self):
        clipboard_text = QApplication.clipboard().text()
        rows = self.parse_clipboard_text(clipboard_text)
        if not rows:
            return

        start_row = max(self.currentRow(), 0)
        start_col = max(self.currentColumn(), 0)

        max_cols = max(len(row) for row in rows)
        target_rows = max(self.rowCount(), start_row + len(rows))
        target_cols = max(self.columnCount(), start_col + max_cols)
        self.resize_with_preserved_values(target_rows, target_cols)

        for row_offset, row_data in enumerate(rows):
            for col_offset, value in enumerate(row_data):
                item = self.item(start_row + row_offset, start_col + col_offset)
                if item is None:
                    item = QTableWidgetItem("")
                    self.setItem(start_row + row_offset, start_col + col_offset, item)
                item.setText(value.strip())

        self.apply_table_chrome()

    def set_text_matrix(self, matrix):
        if not matrix:
            self.clear_contents_only()
            return

        row_count = len(matrix)
        col_count = max(len(row) for row in matrix)
        self.resize_with_preserved_values(row_count, col_count)

        for row_index in range(row_count):
            for col_index in range(col_count):
                item = self.item(row_index, col_index)
                if item is None:
                    item = QTableWidgetItem("")
                    self.setItem(row_index, col_index, item)

                value = (
                    matrix[row_index][col_index]
                    if col_index < len(matrix[row_index])
                    else ""
                )
                item.setText(value)

        self.apply_table_chrome()

    def get_text_matrix(self):
        matrix = []
        for row in range(self.rowCount()):
            current_row = []
            for col in range(self.columnCount()):
                item = self.item(row, col)
                current_row.append(item.text() if item else "")
            matrix.append(current_row)
        return matrix

    def resize_with_preserved_values(self, rows, cols):
        existing = self.get_text_matrix()
        self.setRowCount(rows)
        self.setColumnCount(cols)

        for row in range(rows):
            for col in range(cols):
                item = self.item(row, col)
                if item is None:
                    item = QTableWidgetItem("")
                    self.setItem(row, col, item)

                if row < len(existing) and col < len(existing[row]):
                    item.setText(existing[row][col])
                else:
                    item.setText("")

        self.apply_table_chrome()
        self.dimensionsChanged.emit(
            max(rows - self.data_row_offset, 1),
            max(cols - self.data_col_offset, 1),
        )

    def clear_contents_only(self):
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if item is None:
                    item = QTableWidgetItem("")
                    self.setItem(row, col, item)
                item.setText("")
        self.apply_table_chrome()

    def apply_table_chrome(self):
        header_bg = QColor("#e8eef7")
        header_fg = QColor("#243447")
        cell_bg = QColor("#ffffff")
        cell_fg = QColor("#1f2937")
        bold_font = QFont("Segoe UI", 10, QFont.Bold)
        regular_font = QFont("Segoe UI", 10)

        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if item is None:
                    item = QTableWidgetItem("")
                    self.setItem(row, col, item)

                item.setTextAlignment(Qt.AlignCenter)
                is_header_cell = (
                    self.highlight_first_row and row == 0
                ) or (
                    self.highlight_first_col and col == 0
                )
                if is_header_cell:
                    item.setBackground(header_bg)
                    item.setForeground(header_fg)
                    item.setFont(bold_font)
                else:
                    item.setBackground(cell_bg)
                    item.setForeground(cell_fg)
                    item.setFont(regular_font)


class AppGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Визуализация корреляционной плеяды")
        self.resize(1240, 760)

        self.rows_count = 7
        self.cols_count = 7

        self.create_widgets()
        self.apply_styles()

    def create_widgets(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(24, 24, 24, 24)
        root_layout.setSpacing(16)

        header_card = QFrame()
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(22, 18, 22, 18)
        header_layout.setSpacing(6)

        title = QLabel("Корреляционная таблица")
        title.setObjectName("titleLabel")
        subtitle = QLabel(
            "Можно работать в двух форматах: как с матрицей корреляций или как со списком "
            "связей из 3 столбцов. Обе вкладки поддерживают вставку таблицы из Excel или "
            "Google Sheets через Ctrl+V."
        )
        subtitle.setWordWrap(True)
        subtitle.setObjectName("subtitleLabel")

        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        root_layout.addWidget(header_card)

        self.input_tabs = QTabWidget()
        self.input_tabs.addTab(self.create_matrix_tab(), "Матрица")
        self.input_tabs.addTab(self.create_edge_list_tab(), "Список связей")
        root_layout.addWidget(self.input_tabs, 1)

        draw_button = QPushButton("Построить плеяду")
        draw_button.setObjectName("primaryButton")
        draw_button.clicked.connect(self.process_data)
        root_layout.addWidget(draw_button)

    def create_matrix_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(16)

        controls_card = QFrame()
        card_layout = QVBoxLayout(controls_card)
        card_layout.setContentsMargins(18, 14, 18, 14)
        card_layout.setSpacing(12)

        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(12)

        rows_label = QLabel("Строк:")
        self.rows_spin = QSpinBox()
        self.rows_spin.setRange(1, 200)
        self.rows_spin.setValue(self.rows_count)

        cols_label = QLabel("Столбцов:")
        self.cols_spin = QSpinBox()
        self.cols_spin.setRange(1, 200)
        self.cols_spin.setValue(self.cols_count)

        resize_button = QPushButton("Изменить размер")
        resize_button.clicked.connect(self.resize_table)

        import_button = QPushButton("Открыть CSV/XLSX")
        import_button.clicked.connect(self.import_matrix_file)

        paste_button = QPushButton("Вставить из буфера")
        paste_button.clicked.connect(self.table_paste)

        clear_button = QPushButton("Очистить")
        clear_button.clicked.connect(self.clear_table)

        hint = QLabel(
            "Верхняя строка задаёт правую группу, первый столбец задаёт левую группу, "
            "внутри таблицы находятся коэффициенты корреляции."
        )
        hint.setObjectName("hintLabel")
        hint.setWordWrap(True)

        card_layout.addWidget(hint)
        controls_layout.addWidget(rows_label)
        controls_layout.addWidget(self.rows_spin)
        controls_layout.addWidget(cols_label)
        controls_layout.addWidget(self.cols_spin)
        controls_layout.addWidget(resize_button)
        controls_layout.addWidget(import_button)
        controls_layout.addWidget(paste_button)
        controls_layout.addWidget(clear_button)
        controls_layout.addStretch(1)
        card_layout.addLayout(controls_layout)
        layout.addWidget(controls_card)

        self.table = PasteTableWidget(self.rows_count + 1, self.cols_count + 1)
        self.table.dimensionsChanged.connect(self.sync_size_controls)
        self.table.apply_table_chrome()
        layout.addWidget(self.table, 1)
        return tab

    def create_edge_list_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(16)

        controls_card = QFrame()
        card_layout = QVBoxLayout(controls_card)
        card_layout.setContentsMargins(18, 14, 18, 14)
        card_layout.setSpacing(12)

        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(12)

        add_rows_button = QPushButton("Добавить 5 строк")
        add_rows_button.clicked.connect(self.add_edge_rows)

        import_button = QPushButton("Открыть CSV/XLSX")
        import_button.clicked.connect(self.import_edge_file)

        paste_button = QPushButton("Вставить из буфера")
        paste_button.clicked.connect(self.edge_table_paste)

        clear_button = QPushButton("Очистить")
        clear_button.clicked.connect(self.clear_edge_table)

        hint = QLabel(
            "Каждая строка описывает одну связь: параметр 1, параметр 2 и корреляция "
            "между ними. Первый столбец станет левой группой, второй столбец правой."
        )
        hint.setObjectName("hintLabel")
        hint.setWordWrap(True)

        card_layout.addWidget(hint)
        controls_layout.addWidget(add_rows_button)
        controls_layout.addWidget(import_button)
        controls_layout.addWidget(paste_button)
        controls_layout.addWidget(clear_button)
        controls_layout.addStretch(1)
        card_layout.addLayout(controls_layout)
        layout.addWidget(controls_card)

        self.edge_table = PasteTableWidget(
            12,
            3,
            highlight_first_row=False,
            highlight_first_col=False,
            show_horizontal_header=True,
            horizontal_labels=["Параметр 1", "Параметр 2", "Корреляция"],
        )
        self.edge_table.apply_table_chrome()
        layout.addWidget(self.edge_table, 1)
        return tab

    def apply_styles(self):
        self.setStyleSheet(
            """
            QWidget {
                background: #f4f7fb;
                color: #1f2937;
                font-family: "Segoe UI";
                font-size: 10.5pt;
            }
            QFrame {
                background: white;
                border: 1px solid #dbe3ef;
                border-radius: 18px;
            }
            QLabel {
                background: transparent;
            }
            QLabel#titleLabel {
                font-size: 20pt;
                font-weight: 700;
                color: #10233a;
            }
            QLabel#subtitleLabel, QLabel#hintLabel {
                color: #516274;
            }
            QSpinBox {
                min-width: 70px;
                padding: 7px 10px;
                border: 1px solid #cad5e3;
                border-radius: 10px;
                background: white;
            }
            QPushButton {
                padding: 9px 16px;
                border: none;
                border-radius: 12px;
                background: #dce8f7;
                color: #14304d;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #cfe0f5;
            }
            QPushButton#primaryButton {
                background: #1e6bb8;
                color: white;
                font-size: 11pt;
                padding: 12px 18px;
            }
            QPushButton#primaryButton:hover {
                background: #175898;
            }
            QTableWidget {
                background: white;
                border: 1px solid #dbe3ef;
                border-radius: 18px;
                gridline-color: #dbe3ef;
                selection-background-color: #b7d4f5;
                selection-color: #10233a;
            }
            QHeaderView::section {
                background: #e8eef7;
                color: #243447;
                border: none;
                border-bottom: 1px solid #dbe3ef;
                padding: 10px;
                font-weight: 700;
            }
            QTabWidget::pane {
                border: none;
            }
            QTabBar::tab {
                background: #dce8f7;
                color: #14304d;
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                min-width: 150px;
                padding: 10px 22px;
                margin-right: 6px;
                font-weight: 600;
            }
            QTabBar::tab:selected {
                background: white;
                color: #10233a;
            }
            """
        )

    def import_table_into_widget(self, target_widget):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть таблицу",
            "",
            "Таблицы (*.csv *.xlsx);;CSV (*.csv);;Excel (*.xlsx)",
        )
        if not file_path:
            return

        try:
            target_widget.set_text_matrix(read_table_file(file_path))
        except ValueError as error:
            QMessageBox.warning(self, "Ошибка импорта", str(error))

    def import_matrix_file(self):
        self.import_table_into_widget(self.table)

    def import_edge_file(self):
        self.import_table_into_widget(self.edge_table)

    def resize_table(self):
        rows = self.rows_spin.value() + 1
        cols = self.cols_spin.value() + 1
        self.table.resize_with_preserved_values(rows, cols)

    def table_paste(self):
        self.table.paste_from_clipboard()

    def clear_table(self):
        self.table.clear_contents_only()

    def add_edge_rows(self):
        self.edge_table.resize_with_preserved_values(self.edge_table.rowCount() + 5, 3)

    def edge_table_paste(self):
        self.edge_table.paste_from_clipboard()

    def clear_edge_table(self):
        self.edge_table.clear_contents_only()

    def sync_size_controls(self, rows, cols):
        self.rows_spin.setValue(rows)
        self.cols_spin.setValue(cols)

    def process_data(self):
        try:
            if self.input_tabs.currentIndex() == 0:
                left_params, right_params, edges = self.collect_matrix_data()
            else:
                left_params, right_params, edges = self.collect_edge_list_data()
        except ValueError as error:
            QMessageBox.warning(self, "Ошибка ввода", str(error))
            return

        draw_final_pleiad(left_params, right_params, edges)

    def collect_matrix_data(self):
        matrix = self.table.get_text_matrix()
        row_count = len(matrix)
        col_count = len(matrix[0]) if matrix else 0

        if row_count < 2 or col_count < 2:
            raise ValueError(
                "Таблица должна содержать хотя бы одну строку и один столбец данных."
            )

        left_params = [
            matrix[row][0].strip() for row in range(1, row_count) if matrix[row][0].strip()
        ]
        right_params = [
            matrix[0][col].strip() for col in range(1, col_count) if matrix[0][col].strip()
        ]

        if not left_params or not right_params:
            raise ValueError(
                "Заполните хотя бы одно название строки и одно название столбца."
            )

        edges = []
        for row in range(1, row_count):
            row_name = matrix[row][0].strip()
            for col in range(1, col_count):
                value_text = matrix[row][col].strip()
                if not value_text:
                    continue

                col_name = matrix[0][col].strip()
                if not row_name or not col_name:
                    raise ValueError(
                        f"Для значения '{value_text}' нужна подпись строки и столбца."
                    )
                if row_name == col_name:
                    continue

                try:
                    weight = float(value_text.replace(",", "."))
                except ValueError as exc:
                    raise ValueError(
                        f"Некорректное число: '{value_text}' в строке {row + 1}, "
                        f"столбце {col + 1}."
                    ) from exc

                edges.append((row_name, col_name, weight))

        return left_params, right_params, edges

    def collect_edge_list_data(self):
        rows = self.edge_table.get_text_matrix()
        data_rows = [row[:3] for row in rows if any(cell.strip() for cell in row[:3])]

        if not data_rows:
            raise ValueError("Заполните хотя бы одну строку списка связей.")

        if self.looks_like_edge_header(data_rows[0]):
            data_rows = data_rows[1:]

        if not data_rows:
            raise ValueError("После строки заголовков не найдено ни одной связи.")

        left_params = []
        right_params = []
        edges = []

        for index, row in enumerate(data_rows, start=1):
            param_1 = row[0].strip() if len(row) > 0 else ""
            param_2 = row[1].strip() if len(row) > 1 else ""
            value_text = row[2].strip() if len(row) > 2 else ""

            if not param_1 or not param_2 or not value_text:
                raise ValueError(
                    f"В строке списка связей {index} должны быть заполнены все 3 столбца."
                )

            try:
                weight = float(value_text.replace(",", "."))
            except ValueError as exc:
                raise ValueError(
                    f"Некорректное число в строке списка связей {index}: '{value_text}'."
                ) from exc

            if param_1 not in left_params:
                left_params.append(param_1)
            if param_2 not in right_params:
                right_params.append(param_2)
            edges.append((param_1, param_2, weight))

        return left_params, right_params, edges

    def looks_like_edge_header(self, row):
        normalized = [cell.strip().lower() for cell in row[:3]]
        header_variants = {
            ("параметр 1", "параметр 2", "корреляция"),
            ("1 параметр", "2 параметр", "корреляция"),
            ("parameter 1", "parameter 2", "correlation"),
        }
        if tuple(normalized) in header_variants:
            return True

        if len(normalized) < 3:
            return False

        try:
            float(normalized[2].replace(",", "."))
        except ValueError:
            return "параметр" in normalized[0] or "parameter" in normalized[0]
        return False


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = AppGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
