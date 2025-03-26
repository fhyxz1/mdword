from kivy.app import App
from kivy.core.text import LabelBase
from kivy.resources import resource_add_path, resource_find
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.scrollview import ScrollView
from kivy.lang import Builder
from kivy.properties import BooleanProperty, OptionProperty
from kivy.clock import Clock
import re
import os
import sys
import time
from kivy.utils import platform

# Android剪贴板处理
if platform == 'android':
    from jnius import autoclass
    PythonActivity = autoclass('org.kivy.android.PythonActivity')
    Context = autoclass('android.content.Context')
    ClipData = autoclass('android.content.ClipData')
    clipboard_service = PythonActivity.mActivity.getSystemService(Context.CLIPBOARD_SERVICE)

    def android_copy(text):
        clip = ClipData.newPlainText("text", text)
        clipboard_service.setPrimaryClip(clip)

    def android_paste():
        clip = clipboard_service.getPrimaryClip()
        if clip and clip.getItemCount() > 0:
            item = clip.getItemAt(0)
            return str(item.coerceToText(PythonActivity.mActivity))
        return ""
else:
    from kivy.core.clipboard import Clipboard
    def android_copy(text):
        Clipboard.copy(text)
    def android_paste():
        return Clipboard.paste()


# 字体处理
if hasattr(sys, '_MEIPASS'):
    resource_add_path(os.path.join(sys._MEIPASS, 'fonts'))
else:
    resource_add_path(os.path.abspath('./fonts'))

LabelBase.register('Roboto', resource_find('SourceHanSansSC-Regular-2.otf'))

Builder.load_string('''
#:kivy 2.0.0
#:import hex kivy.utils.get_color_from_hex

<MarkdownTool>:
    orientation: 'vertical'
    spacing: '8dp'
    padding: '8dp'

    BoxLayout:
        orientation: 'vertical'
        size_hint_y: 0.30
        Label:
            text: '输入区'
            size_hint_y: None
            height: '25dp'
            font_size: '14sp'
        ScrollView:
            TextInput:
                id: input_area
                text: ''
                hint_text: '在此输入或粘贴Markdown内容...'
                size_hint_y: None
                height: max(self.minimum_height, dp(200))
                background_color: hex('#FFFFFF')
                foreground_color: hex('#333333')
                font_size: '14sp'
                on_text: root.auto_process_and_update()

    BoxLayout:
        orientation: 'vertical'
        size_hint_y: 0.35
        spacing: '4dp'
        padding: '4dp'
        Label:
            text: '选项区'
            size_hint_y: None
            height: '25dp'
            font_size: '14sp'
        ScrollView:
            GridLayout:
                cols: 2
                size_hint_y: None
                height: self.minimum_height
                row_default_height: '32dp'
                row_force_default: True
                spacing: '8dp'
                padding: '4dp'
                CheckBox:
                    size_hint_x: None
                    width: '32dp'
                    active: root.remove_italic
                    on_active: root.remove_italic = self.active
                Label:
                    text: '去除斜体'
                    font_size: '12sp'
                CheckBox:
                    size_hint_x: None
                    width: '32dp'
                    active: root.remove_strikethrough
                    on_active: root.remove_strikethrough = self.active
                Label:
                    text: '删除线'
                    font_size: '12sp'
                CheckBox:
                    size_hint_x: None
                    width: '32dp'
                    active: root.remove_highlight
                    on_active: root.remove_highlight = self.active
                Label:
                    text: '去除高亮'
                    font_size: '12sp'
                CheckBox:
                    size_hint_x: None
                    width: '32dp'
                    active: root.remove_links
                    on_active: root.remove_links = self.active
                Label:
                    text: '去除链接'
                    font_size: '12sp'
                CheckBox:
                    size_hint_x: None
                    width: '32dp'
                    active: root.remove_unordered_list
                    on_active: root.remove_unordered_list = self.active
                Label:
                    text: '清洗无序列表'
                    font_size: '12sp'
                CheckBox:
                    size_hint_x: None
                    width: '32dp'
                    active: root.remove_ordered_list
                    on_active: root.remove_ordered_list = self.active
                Label:
                    text: '清洗有序列表'
                    font_size: '12sp'
                CheckBox:
                    size_hint_x: None
                    width: '32dp'
                    active: root.table_clean
                    on_active: root.table_clean = self.active
                Label:
                    text: '表格清洁'
                    font_size: '12sp'
                Label:
                    text: '表格转换:'
                    font_size: '12sp'
                Spinner:
                    id: table_spinner
                    text: root.table_conversion
                    values: ['无', '空格', '/t', ',']
                    font_size: '12sp'
                    size_hint_y: None
                    height: '32dp'
                    on_text: root.table_conversion = self.text

    BoxLayout:
        orientation: 'vertical'
        size_hint_y: 0.30
        Label:
            text: '输出区'
            size_hint_y: None
            height: '25dp'
            font_size: '14sp'
        BoxLayout:
            size_hint_y: None
            height: '32dp'
            spacing: '4dp'
            CustomButton:
                text: '清空输入'
                on_press: root.process_reset('input')
                font_size: '8sp'
                size_hint_x: 1
            CustomButton:
                text: '读取剪贴板'
                on_press: root.paste_from_clipboard()
                font_size: '9sp'
                size_hint_x: 1
            CustomButton:
                text: '清空输出'
                on_press: root.process_reset('output')
                font_size: '8sp'
                size_hint_x: 1
            CustomButton:
                text: '复制结果'
                on_press: root.copy_to_clipboard()
                font_size: '8sp'
                size_hint_x: 1
        ScrollView:
            TextInput:
                id: output_area
                text: ''
                hint_text: '处理后的纯净文本...'
                size_hint_y: None
                height: max(self.minimum_height, dp(200))
                background_color: hex('#F5F5F5')
                foreground_color: hex('#333333')
                font_size: '14sp'

<CustomButton@Button>:
    font_size: '12sp'
    background_normal: ''
    background_color: hex('#4CAF50') if self.state == 'normal' else hex('#45a049')
    color: hex('#FFFFFF')
    size_hint: (None, None)
    size: ('80dp', '32dp')
    padding: ('4dp', '4dp')
''')

# [剩余类定义保持不变...]

class MarkdownTool(BoxLayout):
    auto_process = BooleanProperty(True)
    remove_italic = BooleanProperty(False)
    remove_strikethrough = BooleanProperty(False)
    remove_highlight = BooleanProperty(False)
    remove_links = BooleanProperty(False)
    remove_unordered_list = BooleanProperty(False)
    remove_ordered_list = BooleanProperty(False)
    table_clean = BooleanProperty(False)
    table_conversion = OptionProperty("无", options=["无", "空格", "/t", ","])

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.bind(
            remove_italic=lambda inst, val: self._option_changed(),
            remove_strikethrough=lambda inst, val: self._option_changed(),
            remove_highlight=lambda inst, val: self._option_changed(),
            remove_links=lambda inst, val: self._option_changed(),
            remove_unordered_list=lambda inst, val: self._option_changed(),
            remove_ordered_list=lambda inst, val: self._option_changed(),
            table_clean=lambda inst, val: self._option_changed(),
            table_conversion=lambda inst, val: self._option_changed()
        )

    def _option_changed(self):
        if self.auto_process:
            self.process_markdown()

    def paste_from_clipboard(self):
        try:
            self.ids.input_area.text = android_paste().strip()
        except Exception as e:
            self.ids.output_area.text = f"剪贴板错误: {str(e)}"

    def auto_process_and_update(self):
        if self.auto_process:
            self.process_markdown()

    def process_markdown(self):
        try:
            text = self.ids.input_area.text
            text = re.sub(r'^#+\s*', '', text, flags=re.MULTILINE)
            text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
            text = re.sub(r"''(.*?)''", r'\1', text)

            if self.remove_italic:
                text = re.sub(r'(?<!\*)\*(?!\*)(.*?)\*(?!\*)', r'\1', text)
                text = re.sub(r'(?<!_)_(?!_)(.*?)_(?!_)', r'\1', text)
            if self.remove_strikethrough:
                text = re.sub(r'~~(.*?)~~', r'\1', text)
            if self.remove_highlight:
                text = re.sub(r'==(.+?)==', r'\1', text)
            if self.remove_links:
                text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
            if self.remove_unordered_list:
                text = re.sub(r'(?m)^\s*[-*+]\s+', '', text)
            if self.remove_ordered_list:
                text = re.sub(r'(?m)^\s*\d+\.\s+', '', text)

            if self.table_clean:
                text = re.sub(r'(?m)^\s*\|?[\s\-|]+\|?\s*$', '', text)
                text = text.replace("|", "")
            elif self.table_conversion != "无":
                lines = text.splitlines()
                processed_lines = []
                for line in lines:
                    if re.match(r'^\s*\|?[\s\-|]+\|?\s*$', line):
                        continue
                    line = line.strip()
                    if line.startswith("|"):
                        line = line[1:]
                    if line.endswith("|"):
                        line = line[:-1]
                    if self.table_conversion == "空格":
                        line = line.replace("|", "    ")
                    elif self.table_conversion == "/t":
                        line = line.replace("|", "\t")
                    elif self.table_conversion == ",":
                        line = line.replace("|", ",")
                    processed_lines.append(line)
                text = "\n".join(processed_lines)

            text = re.sub(r'(?m)^(?:\s*[-*_]{3,}\s*)$', '', text)
            self.ids.output_area.text = text.strip()
        except Exception as e:
            self.ids.output_area.text = f"处理错误: {str(e)}"

    def copy_to_clipboard(self):
        try:
            android_copy(self.ids.output_area.text)
        except Exception as e:
            self.ids.output_area.text = f"复制失败: {str(e)}"

    def process_reset(self, target):
        getattr(self.ids, f"{target}_area").text = ''

class MarkdownApp(App):
    def build(self):
        self.title = 'mdword'
        return MarkdownTool()

if __name__ == '__main__':
    MarkdownApp().run()