from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from . import config
from .gui_formatter import format_result_payload
from . import template_types
from .section_rules import BodyRangeRule, SectionPosition, SectionRule
from .template_parser import TemplateParser


def _require_pyside6():
    try:
        from PySide6 import QtCore, QtGui, QtWidgets  # noqa: F401
    except ImportError as exc:
        raise SystemExit("PySide6 未安装，请先安装 PySide6") from exc
    return QtCore, QtGui, QtWidgets


def _split_text_tokens(text: str) -> list[str]:
    tokens = []
    for raw in (
        text.replace("；", ",")
        .replace(";", ",")
        .replace("，", ",")
        .replace("/", ",")
        .replace("|", ",")
        .split(",")
    ):
        token = raw.strip()
        if token:
            tokens.append(token)
    return tokens


def _parse_position_text(text: str) -> SectionPosition:
    normalized = (text or "").strip().lower()
    mapping = {
        "首页": SectionPosition.FIRST_PAGE,
        "首页页": SectionPosition.FIRST_PAGE,
        "first": SectionPosition.FIRST_PAGE,
        "first_page": SectionPosition.FIRST_PAGE,
        "前置": SectionPosition.FRONT,
        "前": SectionPosition.FRONT,
        "front": SectionPosition.FRONT,
        "正文": SectionPosition.BODY,
        "正文中": SectionPosition.BODY,
        "body": SectionPosition.BODY,
        "后置": SectionPosition.BACK,
        "后": SectionPosition.BACK,
        "back": SectionPosition.BACK,
        "尾页": SectionPosition.LAST_PAGE,
        "last": SectionPosition.LAST_PAGE,
        "last_page": SectionPosition.LAST_PAGE,
    }
    return mapping.get(normalized, SectionPosition.BODY)


def _parse_body_range_text(text: str) -> BodyRangeRule:
    normalized = (text or "").strip().lower()
    mapping = {
        "直到下一个标题": BodyRangeRule.UNTIL_NEXT_TITLE,
        "到下一个标题": BodyRangeRule.UNTIL_NEXT_TITLE,
        "until_next_title": BodyRangeRule.UNTIL_NEXT_TITLE,
        "到空行": BodyRangeRule.UNTIL_BLANK,
        "直到空行": BodyRangeRule.UNTIL_BLANK,
        "until_blank": BodyRangeRule.UNTIL_BLANK,
        "固定段数": BodyRangeRule.FIXED_PARAGRAPHS,
        "fixed_paragraphs": BodyRangeRule.FIXED_PARAGRAPHS,
    }
    return mapping.get(normalized, BodyRangeRule.UNTIL_NEXT_TITLE)


def _parse_int(text: str) -> int | None:
    value = (text or "").strip()
    if not value:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def main() -> None:
    QtCore, QtGui, QtWidgets = _require_pyside6()

    class ParserWorker(QtCore.QObject):
        finished = QtCore.Signal(bool, str, str)

        def __init__(
            self,
            template_path: str,
            output_path: str | None,
            allow_fallback: bool,
            template_type: str,
            section_rules: list[template_types.SectionRule] | None,
        ):
            super().__init__()
            self.template_path = template_path
            self.output_path = output_path
            self.allow_fallback = allow_fallback
            self.template_type = template_type
            self.section_rules = section_rules
            self.max_heading_level = config.DEFAULT_MAX_HEADING_LEVEL
            self.required_roles = list(config.DEFAULT_REQUIRED_ROLES)
            self.required_on_presence_map = dict(config.DEFAULT_REQUIRED_ON_PRESENCE_MAP)

        @QtCore.Slot()
        def run(self) -> None:
            parser = TemplateParser(
                max_heading_level=self.max_heading_level,
                required_roles=self.required_roles,
                required_on_presence_map=self.required_on_presence_map,
                allow_fallback=self.allow_fallback,
                strict=False,
                template_type=self.template_type,
                section_rules=self.section_rules,
            )
            try:
                result = parser.parse(self.template_path)
                parser.export_json(result, output_path=self.output_path)
                resolved_output = self.output_path or str(config.DEFAULT_OUTPUT_PATH)
                self.finished.emit(True, resolved_output, "解析完成")
            except Exception as exc:
                self.finished.emit(False, "", str(exc))

    class MainWindow(QtWidgets.QWidget):
        def __init__(self) -> None:
            super().__init__()
            self._thread: QtCore.QThread | None = None
            self._worker: ParserWorker | None = None
            self._user_role: int | None = None
            self._custom_templates: dict[str, template_types.TemplateType] = {}
            self._loading_rules = False
            self._save_timer = None
            self._build_ui(QtWidgets, QtGui, QtCore)

        def _build_ui(self, QtWidgets, QtGui, QtCore) -> None:
            self._QtWidgets = QtWidgets
            self._QtGui = QtGui
            self._QtCore = QtCore
            self._user_role = getattr(QtCore.Qt, "UserRole", 32)
            self.setWindowTitle("论文自动排版工具（里程碑1）")
            self.resize(720, 480)

            self.template_edit = QtWidgets.QLineEdit()
            self.template_button = QtWidgets.QPushButton("选择...")
            self.output_edit = QtWidgets.QLineEdit()
            self.output_button = QtWidgets.QPushButton("选择...")
            self.allow_fallback_check = QtWidgets.QCheckBox("缺失字段自动兜底")
            self.allow_fallback_check.setChecked(True)
            self.template_type_combo = QtWidgets.QComboBox()
            self.add_template_button = QtWidgets.QPushButton("新增模板")
            self.edit_template_button = QtWidgets.QPushButton("编辑模板")
            self.delete_template_button = QtWidgets.QPushButton("删除模板")
            self.section_table = QtWidgets.QTableWidget()
            self.section_table.setColumnCount(8)
            self.section_table.setHorizontalHeaderLabels(
                [
                    "启用",
                    "板块名称",
                    "标题关键词",
                    "内容关键词 (?)",
                    "标题样式名 (?)",
                    "位置 (?)",
                    "正文范围 (?)",
                    "固定段数 (?)",
                ]
            )
            self._set_section_table_header_tooltips()
            header = getattr(self.section_table, "horizontalHeader", None)
            if callable(header):
                section_header = header()
                if section_header is not None:
                    stretch = getattr(section_header, "setStretchLastSection", None)
                    if callable(stretch):
                        stretch(True)
            self.add_section_button = QtWidgets.QPushButton("新增板块")
            self.remove_section_button = QtWidgets.QPushButton("删除选中")
            self.parse_button = QtWidgets.QPushButton("解析模板")
            self.open_button = QtWidgets.QPushButton("打开输出目录")
            self.status_label = QtWidgets.QLabel("就绪")
            self.progress = QtWidgets.QProgressBar()
            self.progress.setRange(0, 1)
            self.result_view = QtWidgets.QTextEdit()
            self.result_view.setReadOnly(True)
            self._reload_custom_templates()
            self._refresh_template_choices()
            self._save_timer = QtCore.QTimer(self)
            self._save_timer.setSingleShot(True)
            self._save_timer.timeout.connect(self._auto_save_current_template)

            self.template_button.clicked.connect(self._choose_template)
            self.output_button.clicked.connect(self._choose_output)
            self.parse_button.clicked.connect(self._start_parse)
            self.open_button.clicked.connect(self._open_output_dir)
            self.add_section_button.clicked.connect(self._add_section_row)
            self.remove_section_button.clicked.connect(self._remove_section_rows)
            self.template_type_combo.currentIndexChanged.connect(self._load_section_rules)
            self.add_template_button.clicked.connect(self._add_template)
            self.edit_template_button.clicked.connect(self._edit_template)
            self.delete_template_button.clicked.connect(self._delete_template)
            self.template_edit.textChanged.connect(self._update_actions)
            item_changed = getattr(self.section_table, "itemChanged", None)
            if callable(item_changed):
                item_changed.connect(self._on_section_table_changed)
            self._load_section_rules()
            self._update_actions()

            form_layout = QtWidgets.QGridLayout()
            form_layout.addWidget(QtWidgets.QLabel("模板文件："), 0, 0)
            form_layout.addWidget(self.template_edit, 0, 1)
            form_layout.addWidget(self.template_button, 0, 2)
            form_layout.addWidget(QtWidgets.QLabel("输出位置："), 1, 0)
            form_layout.addWidget(self.output_edit, 1, 1)
            form_layout.addWidget(self.output_button, 1, 2)
            form_layout.addWidget(QtWidgets.QLabel("模板类型："), 2, 0)
            form_layout.addWidget(self.template_type_combo, 2, 1)
            template_buttons_layout = QtWidgets.QHBoxLayout()
            template_buttons_layout.addWidget(self.add_template_button)
            template_buttons_layout.addWidget(self.edit_template_button)
            template_buttons_layout.addWidget(self.delete_template_button)
            template_buttons_box = QtWidgets.QWidget()
            template_buttons_box.setLayout(template_buttons_layout)
            form_layout.addWidget(template_buttons_box, 2, 2)
            form_layout.addWidget(self.allow_fallback_check, 3, 1)

            section_layout = QtWidgets.QVBoxLayout()
            section_layout.addWidget(QtWidgets.QLabel("板块识别设置："))
            section_layout.addWidget(self.section_table)
            section_buttons = QtWidgets.QHBoxLayout()
            section_buttons.addWidget(self.add_section_button)
            section_buttons.addWidget(self.remove_section_button)
            section_buttons.addStretch(1)
            section_layout.addLayout(section_buttons)

            button_layout = QtWidgets.QHBoxLayout()
            button_layout.addWidget(self.parse_button)
            button_layout.addWidget(self.open_button)
            button_layout.addStretch(1)

            status_layout = QtWidgets.QHBoxLayout()
            status_layout.addWidget(QtWidgets.QLabel("状态："))
            status_layout.addWidget(self.status_label)
            status_layout.addStretch(1)
            status_layout.addWidget(self.progress)

            layout = QtWidgets.QVBoxLayout()
            layout.addLayout(form_layout)
            layout.addLayout(section_layout)
            layout.addLayout(button_layout)
            layout.addLayout(status_layout)
            layout.addWidget(QtWidgets.QLabel("解析结果："))
            layout.addWidget(self.result_view)
            self.setLayout(layout)

        def _update_actions(self) -> None:
            has_template = bool(self.template_edit.text().strip())
            self.parse_button.setEnabled(has_template and (not self._thread_running()))
            self._update_template_actions()

        def _thread_running(self) -> bool:
            if self._thread is None:
                return False
            is_running = getattr(self._thread, "isRunning", None)
            if callable(is_running):
                return bool(is_running())
            return True

        def _reload_custom_templates(self) -> None:
            self._custom_templates = {
                template.key: template
                for template in template_types.load_custom_template_types()
            }

        def _refresh_template_choices(self, selected_key: str | None = None) -> None:
            block = getattr(self.template_type_combo, "blockSignals", None)
            if callable(block):
                block(True)
            try:
                clear = getattr(self.template_type_combo, "clear", None)
                if callable(clear):
                    clear()
                for key, label in template_types.get_template_type_choices():
                    self.template_type_combo.addItem(label, key)
                if selected_key:
                    find = getattr(self.template_type_combo, "findData", None)
                    if callable(find):
                        index = find(selected_key)
                    else:
                        index = -1
                    if index is not None and index >= 0:
                        set_current = getattr(self.template_type_combo, "setCurrentIndex", None)
                        if callable(set_current):
                            set_current(index)
            finally:
                if callable(block):
                    block(False)

        def _update_template_actions(self) -> None:
            key = self._current_template_type()
            is_auto = key.lower() == "auto"
            is_custom = key in self._custom_templates
            if hasattr(self, "edit_template_button"):
                self.edit_template_button.setEnabled(not is_auto)
            if hasattr(self, "delete_template_button"):
                self.delete_template_button.setEnabled(is_custom)

        def _on_section_table_changed(self, *_args) -> None:
            if self._loading_rules:
                return
            self._schedule_auto_save()

        def _schedule_auto_save(self) -> None:
            if self._loading_rules:
                return
            if self._save_timer is None:
                return
            self._save_timer.start(400)

        def _auto_save_current_template(self) -> None:
            if self._loading_rules or self._thread_running():
                return
            template_key = self._current_template_type()
            if template_key.lower() == "auto":
                return
            rules, errors = self._collect_section_rules_for_save()
            if errors:
                self.status_label.setText(f"未保存：{errors[0]}")
                return
            display_name = str(self.template_type_combo.currentText()).strip() or template_key
            try:
                template = template_types.TemplateType(
                    key=template_key,
                    display_name=display_name,
                    section_rules=tuple(rules),
                )
                template.validate()
            except ValueError as exc:
                self.status_label.setText(f"未保存：{exc}")
                return
            self._custom_templates[template_key] = template
            template_types.save_custom_template_types(self._custom_templates.values())
            self.status_label.setText("已保存")

        def _collect_section_rules_for_save(self) -> tuple[list[SectionRule], list[str]]:
            rules: list[SectionRule] = []
            errors: list[str] = []
            checked = getattr(self._QtCore.Qt, "Checked", 2)
            for row in range(self.section_table.rowCount()):
                enable_item = self.section_table.item(row, 0)
                if enable_item is None:
                    continue
                state = enable_item.checkState()
                if state != checked:
                    continue
                name_item = self.section_table.item(row, 1)
                keyword_item = self.section_table.item(row, 2)
                content_item = self.section_table.item(row, 3)
                style_item = self.section_table.item(row, 4)
                position_text = self._read_combo_text(row, 5)
                range_text = self._read_combo_text(row, 6)
                limit_item = self.section_table.item(row, 7)
                display_name = (name_item.text() if name_item else "").strip()
                if not display_name:
                    errors.append(f"第{row + 1}行板块名称为空")
                    continue
                raw_key = name_item.data(self._user_role) if name_item and self._user_role is not None else ""
                key = str(raw_key).strip() if raw_key else ""
                if not key:
                    key = self._slugify(display_name) or self._next_custom_key()
                keywords_text = keyword_item.text() if keyword_item else ""
                keywords = tuple(_split_text_tokens(keywords_text))
                content_text = content_item.text() if content_item else ""
                content_keywords = tuple(_split_text_tokens(content_text))
                if not keywords and not content_keywords:
                    errors.append(f"板块“{display_name}”缺少关键词")
                    continue
                style_text = style_item.text() if style_item else ""
                style_names = tuple(_split_text_tokens(style_text))
                position = _parse_position_text(position_text or "")
                body_range = _parse_body_range_text(range_text or "")
                limit_text = limit_item.text() if limit_item else ""
                body_limit = _parse_int(limit_text)
                try:
                    rule = SectionRule(
                        key=key,
                        display_name=display_name,
                        title_keywords=keywords,
                        content_keywords=content_keywords,
                        title_style_names=style_names,
                        position=position,
                        body_range=body_range,
                        body_paragraph_limit=body_limit,
                    )
                    rule.validate()
                except ValueError as exc:
                    errors.append(str(exc))
                    continue
                rules.append(rule)
            return rules, errors

        def _choose_template(self) -> None:
            path, _ = self._QtWidgets.QFileDialog.getOpenFileName(
                self,
                "选择模板文件",
                "",
                "Word 文件 (*.docx)",
            )
            if path:
                self.template_edit.setText(path)

        def _choose_output(self) -> None:
            path, _ = self._QtWidgets.QFileDialog.getSaveFileName(
                self,
                "选择输出位置",
                str(config.DEFAULT_OUTPUT_PATH),
                "JSON 文件 (*.json)",
            )
            if path:
                self.output_edit.setText(path)

        def _set_busy(self, busy: bool) -> None:
            if busy:
                self.progress.setRange(0, 0)
                self.status_label.setText("解析中...")
            else:
                self.progress.setRange(0, 1)
                self.status_label.setText("就绪")
            self.template_button.setEnabled(not busy)
            self.output_button.setEnabled(not busy)
            self.template_edit.setEnabled(not busy)
            self.output_edit.setEnabled(not busy)
            self.allow_fallback_check.setEnabled(not busy)
            self.template_type_combo.setEnabled(not busy)
            self.add_template_button.setEnabled(not busy)
            self.edit_template_button.setEnabled(not busy)
            self.delete_template_button.setEnabled(not busy)
            self.section_table.setEnabled(not busy)
            self.add_section_button.setEnabled(not busy)
            self.remove_section_button.setEnabled(not busy)
            self.open_button.setEnabled(not busy)
            can_parse = (not busy) and bool(self.template_edit.text().strip())
            if self._thread_running():
                can_parse = False
            self.parse_button.setEnabled(can_parse)

        def _current_template_type(self) -> str:
            data = self.template_type_combo.currentData()
            if data:
                return str(data)
            return str(self.template_type_combo.currentText()).strip()

        def _load_section_rules(self) -> None:
            template_type = self._current_template_type()
            self._loading_rules = True
            try:
                rules, checked_keys = self._resolve_section_rules_for_template(template_type)
                self._set_section_table_rules(rules, checked_keys)
            finally:
                self._loading_rules = False
            self._update_template_actions()

        def _resolve_section_rules_for_template(
            self,
            template_type: str,
        ) -> tuple[list[SectionRule], set[str]]:
            if template_type.lower() == "auto":
                selected = template_types.resolve_template_type("generic")
                checked_keys: set[str] = set()
            else:
                selected = template_types.resolve_template_type(template_type)
                checked_keys = {rule.key for rule in selected.section_rules}
                if template_type not in self._custom_templates:
                    checked_keys.add("cover")

            rules_by_key: dict[str, SectionRule] = {}
            order: list[str] = []

            def _add_rule(rule: SectionRule) -> None:
                if rule.key in rules_by_key:
                    return
                rules_by_key[rule.key] = rule
                order.append(rule.key)

            for rule in selected.section_rules:
                _add_rule(rule)
            _add_rule(self._cover_section_rule())
            for template in template_types.iter_template_types():
                for rule in template.section_rules:
                    _add_rule(rule)

            order_index = {key: idx for idx, key in enumerate(order)}
            rules = [rules_by_key[key] for key in order]
            rules.sort(
                key=lambda rule: (
                    0 if rule.key in checked_keys else 1,
                    order_index.get(rule.key, 0),
                )
            )
            return rules, checked_keys

        def _cover_section_rule(self) -> SectionRule:
            return SectionRule(
                key="cover",
                display_name="首页识别",
                title_keywords=("封面", "扉页", "首页"),
                content_keywords=(
                    "专业",
                    "年级",
                    "姓名",
                    "作者",
                    "学号",
                    "学院",
                    "院系",
                    "系别",
                    "班级",
                    "单位",
                    "学校",
                    "学生",
                    "指导教师",
                    "指导老师",
                    "导师",
                    "日期",
                    "时间",
                ),
                position=SectionPosition.FIRST_PAGE,
                body_range=BodyRangeRule.UNTIL_BLANK,
            )

        def _set_section_table_rules(
            self,
            rules: list[SectionRule],
            checked_keys: set[str],
        ) -> None:
            self.section_table.setRowCount(0)
            for rule in rules:
                row = self.section_table.rowCount()
                self.section_table.insertRow(row)
                enable_item = self._QtWidgets.QTableWidgetItem()
                if self._user_role is not None:
                    enable_item.setData(self._user_role, rule.key)
                self._set_item_checkable(enable_item)
                state = getattr(
                    self._QtCore.Qt,
                    "Checked" if rule.key in checked_keys else "Unchecked",
                    0,
                )
                enable_item.setCheckState(state)
                self.section_table.setItem(row, 0, enable_item)
                name_item = self._QtWidgets.QTableWidgetItem(rule.display_name)
                if self._user_role is not None:
                    name_item.setData(self._user_role, rule.key)
                self.section_table.setItem(row, 1, name_item)
                self.section_table.setItem(
                    row,
                    2,
                    self._QtWidgets.QTableWidgetItem("，".join(rule.title_keywords)),
                )
                self.section_table.setItem(
                    row,
                    3,
                    self._QtWidgets.QTableWidgetItem("，".join(rule.content_keywords)),
                )
                self.section_table.setItem(
                    row,
                    4,
                    self._QtWidgets.QTableWidgetItem("，".join(rule.title_style_names)),
                )
                self._set_position_combo(row, rule.position)
                self._set_body_range_combo(row, rule.body_range)
                limit_text = "" if rule.body_paragraph_limit is None else str(rule.body_paragraph_limit)
                self.section_table.setItem(
                    row,
                    7,
                    self._QtWidgets.QTableWidgetItem(limit_text),
                )

        def _format_position(self, position: SectionPosition) -> str:
            mapping = {
                SectionPosition.FIRST_PAGE: "首页",
                SectionPosition.FRONT: "前置",
                SectionPosition.BODY: "正文",
                SectionPosition.BACK: "后置",
                SectionPosition.LAST_PAGE: "尾页",
            }
            return mapping.get(position, "正文")

        def _format_body_range(self, body_range: BodyRangeRule) -> str:
            mapping = {
                BodyRangeRule.UNTIL_NEXT_TITLE: "直到下一个标题",
                BodyRangeRule.UNTIL_BLANK: "到空行",
                BodyRangeRule.FIXED_PARAGRAPHS: "固定段数",
            }
            return mapping.get(body_range, "直到下一个标题")

        def _set_section_table_header_tooltips(self) -> None:
            tips = {
                0: "勾选后启用该板块规则。",
                1: "用于输出展示的板块名称。",
                2: "用于识别标题的关键词，多个可用逗号/中文逗号/分号/斜杠/竖线分隔，空白会被忽略。",
                3: "用于识别正文内容的关键词（可选）。命中后可作为正文线索。",
                4: "Word 中标题的样式名称（可选）。命中样式名时优先作为标题。",
                5: "标题位置：下拉选择。首页=第一页；前置=文首 1/3(最多 30 段)；正文=全文；后置=文末 1/3(最多 30 段)；尾页=最后一页。",
                6: "正文范围：下拉选择。直到下一个标题 / 到空行 / 固定段数。",
                7: "仅当正文范围为“固定段数”时填写，表示标题后正文段落数。",
            }
            for col, tip in tips.items():
                item = self.section_table.horizontalHeaderItem(col)
                if item is None:
                    continue
                item.setToolTip(tip)

        def _add_section_row(self) -> None:
            row = self.section_table.rowCount()
            self.section_table.insertRow(row)
            enable_item = self._QtWidgets.QTableWidgetItem()
            checked = getattr(self._QtCore.Qt, "Checked", 2)
            self._set_item_checkable(enable_item)
            enable_item.setCheckState(checked)
            self.section_table.setItem(row, 0, enable_item)
            name_item = self._QtWidgets.QTableWidgetItem("")
            if self._user_role is not None:
                name_item.setData(self._user_role, self._next_custom_key())
            self.section_table.setItem(row, 1, name_item)
            self.section_table.setItem(row, 2, self._QtWidgets.QTableWidgetItem(""))
            self.section_table.setItem(row, 3, self._QtWidgets.QTableWidgetItem(""))
            self.section_table.setItem(row, 4, self._QtWidgets.QTableWidgetItem(""))
            self._set_position_combo(row, SectionPosition.BODY)
            self._set_body_range_combo(row, BodyRangeRule.UNTIL_NEXT_TITLE)
            self.section_table.setItem(row, 7, self._QtWidgets.QTableWidgetItem(""))
            self._schedule_auto_save()

        def _remove_section_rows(self) -> None:
            selection_model = getattr(self.section_table, "selectionModel", None)
            if not callable(selection_model):
                return
            model = selection_model()
            if model is None:
                return
            selected_rows = getattr(model, "selectedRows", None)
            if not callable(selected_rows):
                return
            indexes = selected_rows()
            for index in sorted(indexes, key=lambda item: item.row(), reverse=True):
                self.section_table.removeRow(index.row())
            self._schedule_auto_save()

        def _next_custom_key(self) -> str:
            existing: set[str] = set()
            for row in range(self.section_table.rowCount()):
                item = self.section_table.item(row, 1)
                if item is None:
                    continue
                key = item.data(self._user_role) if self._user_role is not None else None
                if isinstance(key, str) and key:
                    existing.add(key)
            index = 1
            while True:
                key = f"custom_{index}"
                if key not in existing:
                    return key
                index += 1

        def _collect_section_rules(self) -> list[SectionRule]:
            rules: list[SectionRule] = []
            unchecked = getattr(self._QtCore.Qt, "Unchecked", 0)
            checked = getattr(self._QtCore.Qt, "Checked", 2)
            for row in range(self.section_table.rowCount()):
                enable_item = self.section_table.item(row, 0)
                if enable_item is None:
                    continue
                state = enable_item.checkState()
                if state != checked:
                    continue
                name_item = self.section_table.item(row, 1)
                keyword_item = self.section_table.item(row, 2)
                content_item = self.section_table.item(row, 3)
                style_item = self.section_table.item(row, 4)
                position_text = self._read_combo_text(row, 5)
                range_text = self._read_combo_text(row, 6)
                limit_item = self.section_table.item(row, 7)
                display_name = (name_item.text() if name_item else "").strip()
                if not display_name:
                    raise ValueError("板块名称不能为空")
                raw_key = name_item.data(self._user_role) if name_item and self._user_role is not None else ""
                key = str(raw_key).strip() if raw_key else ""
                if not key:
                    key = self._slugify(display_name) or self._next_custom_key()
                keywords_text = keyword_item.text() if keyword_item else ""
                keywords = tuple(_split_text_tokens(keywords_text))
                content_text = content_item.text() if content_item else ""
                content_keywords = tuple(_split_text_tokens(content_text))
                if not keywords and not content_keywords:
                    raise ValueError(f"板块“{display_name}”缺少关键词")
                style_text = style_item.text() if style_item else ""
                style_names = tuple(_split_text_tokens(style_text))
                position_text = position_text or ""
                body_range_text = range_text or ""
                limit_text = limit_item.text() if limit_item else ""
                position = _parse_position_text(position_text)
                body_range = _parse_body_range_text(body_range_text)
                body_limit = _parse_int(limit_text)
                rule = SectionRule(
                    key=key,
                    display_name=display_name,
                    title_keywords=keywords,
                    content_keywords=content_keywords,
                    title_style_names=style_names,
                    position=position,
                    body_range=body_range,
                    body_paragraph_limit=body_limit,
                )
                rule.validate()
                rules.append(rule)
            return rules

        def _add_template(self) -> None:
            name, ok = self._QtWidgets.QInputDialog.getText(
                self,
                "新增模板",
                "模板名称：",
            )
            if not ok:
                return
            name = (name or "").strip()
            if not name:
                self._QtWidgets.QMessageBox.warning(self, "模板名称错误", "模板名称不能为空")
                return
            key = self._generate_template_key(name)
            rules, errors = self._collect_section_rules_for_save()
            if errors:
                self._QtWidgets.QMessageBox.warning(self, "模板保存失败", errors[0])
                return
            template = template_types.TemplateType(
                key=key,
                display_name=name,
                section_rules=tuple(rules),
            )
            try:
                template.validate()
            except ValueError as exc:
                self._QtWidgets.QMessageBox.warning(self, "模板保存失败", str(exc))
                return
            self._custom_templates[key] = template
            template_types.save_custom_template_types(self._custom_templates.values())
            self._refresh_template_choices(selected_key=key)
            self._load_section_rules()
            self.status_label.setText("已保存")

        def _edit_template(self) -> None:
            template_key = self._current_template_type()
            if template_key.lower() == "auto":
                return
            current_name = str(self.template_type_combo.currentText()).strip()
            name, ok = self._QtWidgets.QInputDialog.getText(
                self,
                "编辑模板",
                "模板名称：",
                text=current_name,
            )
            if not ok:
                return
            name = (name or "").strip()
            if not name:
                self._QtWidgets.QMessageBox.warning(self, "模板名称错误", "模板名称不能为空")
                return
            rules, errors = self._collect_section_rules_for_save()
            if errors:
                self._QtWidgets.QMessageBox.warning(self, "模板保存失败", errors[0])
                return
            template = template_types.TemplateType(
                key=template_key,
                display_name=name,
                section_rules=tuple(rules),
            )
            try:
                template.validate()
            except ValueError as exc:
                self._QtWidgets.QMessageBox.warning(self, "模板保存失败", str(exc))
                return
            self._custom_templates[template_key] = template
            template_types.save_custom_template_types(self._custom_templates.values())
            self._refresh_template_choices(selected_key=template_key)
            self._load_section_rules()
            self.status_label.setText("已保存")

        def _delete_template(self) -> None:
            template_key = self._current_template_type()
            if template_key.lower() == "auto":
                return
            if template_key not in self._custom_templates:
                self._QtWidgets.QMessageBox.warning(self, "删除失败", "内置模板不可删除")
                return
            self._custom_templates.pop(template_key, None)
            template_types.save_custom_template_types(self._custom_templates.values())
            fallback_key = "generic"
            self._refresh_template_choices(selected_key=fallback_key)
            self._load_section_rules()
            self.status_label.setText("已保存")

        def _generate_template_key(self, name: str) -> str:
            base = self._slugify(name)
            if not base:
                base = "template"
            existing = {
                key.lower()
                for key, _ in template_types.get_template_type_choices()
            }
            candidate = base
            index = 1
            while candidate.lower() in existing:
                candidate = f"{base}_{index}"
                index += 1
            return candidate

        def _set_position_combo(self, row: int, position: SectionPosition) -> None:
            combo = self._QtWidgets.QComboBox()
            combo.addItem("首页", SectionPosition.FIRST_PAGE.value)
            combo.addItem("前置", SectionPosition.FRONT.value)
            combo.addItem("正文", SectionPosition.BODY.value)
            combo.addItem("后置", SectionPosition.BACK.value)
            combo.addItem("尾页", SectionPosition.LAST_PAGE.value)
            self._select_combo_value(combo, position.value)
            combo.currentIndexChanged.connect(self._on_section_table_changed)
            self.section_table.setCellWidget(row, 5, combo)

        def _set_body_range_combo(self, row: int, body_range: BodyRangeRule) -> None:
            combo = self._QtWidgets.QComboBox()
            combo.addItem("直到下一个标题", BodyRangeRule.UNTIL_NEXT_TITLE.value)
            combo.addItem("到空行", BodyRangeRule.UNTIL_BLANK.value)
            combo.addItem("固定段数", BodyRangeRule.FIXED_PARAGRAPHS.value)
            self._select_combo_value(combo, body_range.value)
            combo.currentIndexChanged.connect(self._on_section_table_changed)
            self.section_table.setCellWidget(row, 6, combo)

        def _set_item_checkable(self, item: object) -> None:
            flags = getattr(item, "flags", None)
            set_flags = getattr(item, "setFlags", None)
            if not callable(flags) or not callable(set_flags):
                return
            current = flags()
            checkable = getattr(self._QtCore.Qt, "ItemIsUserCheckable", None)
            enabled = getattr(self._QtCore.Qt, "ItemIsEnabled", None)
            if checkable is not None:
                current |= checkable
            if enabled is not None:
                current |= enabled
            set_flags(current)

        def _select_combo_value(self, combo: object, value: str) -> None:
            index = -1
            count = getattr(combo, "count", None)
            if callable(count):
                total = int(count())
                for i in range(total):
                    data = combo.itemData(i)
                    if str(data) == value:
                        index = i
                        break
            if index < 0:
                index = 0
            combo.setCurrentIndex(index)

        def _read_combo_text(self, row: int, col: int) -> str:
            widget = self.section_table.cellWidget(row, col)
            if widget is not None:
                data = widget.currentData()
                if data is not None:
                    return str(data)
                text = widget.currentText()
                if text:
                    return str(text)
            item = self.section_table.item(row, col)
            return item.text() if item else ""

        def _slugify(self, name: str) -> str:
            normalized = name.strip().lower().replace(" ", "_")
            normalized = re.sub(r"[^a-z0-9_]+", "", normalized)
            return normalized

        def _start_parse(self) -> None:
            template_path = self.template_edit.text().strip()
            if not template_path:
                return
            output_path = self.output_edit.text().strip() or None
            allow_fallback = self.allow_fallback_check.isChecked()
            template_type = self._current_template_type()
            try:
                custom_rules = self._collect_section_rules()
            except ValueError as exc:
                self._QtWidgets.QMessageBox.warning(self, "板块配置错误", str(exc))
                return
            if custom_rules:
                section_rules = list(custom_rules)
            elif template_type.lower() == "auto":
                section_rules = None
            elif self.section_table.rowCount() > 0:
                section_rules = []
            else:
                section_rules = list(
                    template_types.resolve_template_type(template_type).section_rules
                )
            self.result_view.setPlainText("")
            self._set_busy(True)

            self._thread = QtCore.QThread()
            self._worker = ParserWorker(
                template_path,
                output_path,
                allow_fallback,
                template_type,
                section_rules,
            )
            self._worker.moveToThread(self._thread)
            self._thread.started.connect(self._worker.run)
            self._worker.finished.connect(self._handle_finished)
            self._worker.finished.connect(self._thread.quit)
            self._worker.finished.connect(self._worker.deleteLater)
            self._thread.finished.connect(self._cleanup_thread)
            self._thread.start()

        def _handle_finished(self, success: bool, output_path: str, message: str) -> None:
            self._set_busy(False)
            self.status_label.setText("完成" if success else "失败")
            if success:
                if not self.output_edit.text().strip():
                    self.output_edit.setText(output_path)
                self._render_result(output_path)
            else:
                self._QtWidgets.QMessageBox.critical(self, "解析失败", message)

        def _cleanup_thread(self) -> None:
            if self._thread is None:
                return
            self._thread.deleteLater()
            self._thread = None
            self._worker = None
            self._update_actions()

        def _render_result(self, output_path: str) -> None:
            if not output_path:
                self.result_view.setPlainText("")
                return
            try:
                with open(output_path, "r", encoding="utf-8") as handle:
                    payload = json.load(handle)
            except Exception as exc:
                self.result_view.setPlainText("")
                self.status_label.setText(f"失败：{exc}")
                return
            self.result_view.setPlainText(format_result_payload(payload))

        def _open_output_dir(self) -> None:
            path_text = self.output_edit.text().strip()
            if path_text:
                target = Path(path_text).parent
            else:
                target = config.DEFAULT_OUTPUT_PATH.parent
            config.ensure_base_dirs()
            opened = self._QtGui.QDesktopServices.openUrl(
                self._QtCore.QUrl.fromLocalFile(str(target))
            )
            if not opened:
                self._QtWidgets.QMessageBox.warning(
                    self,
                    "打开失败",
                    f"无法打开目录：{target}",
                )

    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
