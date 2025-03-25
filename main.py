from kivy.app import App
from kivy.core.text import LabelBase
from kivy.resources import resource_add_path, resource_find
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.core.window import Window
from kivy.lang import Builder
from kivy.properties import BooleanProperty
from kivy.clock import Clock
import re
import threading
import pyperclip  # 更可靠的剪贴板库
import keyboard  # 全局快捷键支持
from pystray import Icon, Menu, MenuItem  # 系统托盘支持
from PIL import Image  # 图标处理
import os
import sys
import time
"""
@Author: ylab
@Date: 2025/3/25
"""
if hasattr(sys, '_MEIPASS'):
    # 打包后的字体目录在 _MEIPASS/fonts 下
    resource_add_path(os.path.join(sys._MEIPASS, 'fonts'))
else:
    resource_add_path(os.path.abspath('./fonts'))

LabelBase.register('Roboto', resource_find('SourceHanSansSC-Regular-2.otf'))

Builder.load_string('''
#:kivy 2.0.0
#:import hex kivy.utils.get_color_from_hex

<MarkdownTool>:
    orientation: 'horizontal'
    spacing: '10sp'
    padding: '10sp'

    BoxLayout:
        orientation: 'vertical'
        size_hint_x: 0.4
        spacing: '5sp'

        BoxLayout:
            size_hint_y: 0.1
            spacing: '5sp'
            CustomButton:
                text: '清空输入'
                on_press: root.process_reset('input')
            CustomButton:
                text: '读取剪贴板'
                on_press: root.paste_from_clipboard()

        TextInput:
            id: input_area
            hint_text: '在此输入或粘贴Markdown内容...'
            background_color: hex('#FFFFFF')
            foreground_color: hex('#333333')
            on_text: root.auto_process_and_update()

    BoxLayout:
        orientation: 'vertical'
        size_hint_x: 0.2
        spacing: '15sp'
        padding: '10sp'

        Label:
            text: '实时统计:'
            font_size: '14sp'
            color: hex('#4CAF50')

        # 处理选项：依次为去除斜体、删除线、高亮、链接
        BoxLayout:
            orientation: 'vertical'
            spacing: '5sp'

            BoxLayout:
                size_hint_y: None
                height: '30sp'
                CheckBox:
                    active: root.remove_italic
                    on_active: root.remove_italic = self.active
                Label:
                    text: '去除斜体'

            BoxLayout:
                size_hint_y: None
                height: '30sp'
                CheckBox:
                    active: root.remove_strikethrough
                    on_active: root.remove_strikethrough = self.active
                Label:
                    text: '删除线'

            BoxLayout:
                size_hint_y: None
                height: '30sp'
                CheckBox:
                    active: root.remove_highlight
                    on_active: root.remove_highlight = self.active
                Label:
                    text: '去除高亮'

            BoxLayout:
                size_hint_y: None
                height: '30sp'
                CheckBox:
                    active: root.remove_links
                    on_active: root.remove_links = self.active
                Label:
                    text: '去除链接'

        BoxLayout:
            orientation: 'vertical'
            Label:
                id: char_count
                text: '字符数: 0'
            Label:
                id: process_time
                text: '处理耗时: 0ms'

    BoxLayout:
        orientation: 'vertical'
        size_hint_x: 0.4
        spacing: '5sp'

        BoxLayout:
            size_hint_y: 0.1
            spacing: '5sp'
            CustomButton:
                text: '清空输出'
                on_press: root.process_reset('output')
            CustomButton:
                text: '复制结果'
                on_press: root.copy_to_clipboard()

        TextInput:
            id: output_area
            hint_text: '处理后的纯净文本...'
            background_color: hex('#F5F5F5')
            foreground_color: hex('#333333')

<CustomButton@Button>:
    font_size: '14sp'
    background_normal: ''
    background_color: hex('#4CAF50') if self.state == 'normal' else hex('#45a049')
    color: hex('#FFFFFF')
    size_hint_y: None
    height: '40sp'
''')

class MarkdownTool(BoxLayout):
    auto_process = BooleanProperty(True)
    remove_italic = BooleanProperty(False)
    remove_strikethrough = BooleanProperty(False)
    remove_highlight = BooleanProperty(False)
    remove_links = BooleanProperty(False)  # 新增去除链接的选项

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self._keyboard = Window.request_keyboard(None, self)
        # 绑定选项变化时动态更新
        self.bind(remove_italic=lambda instance, value: self._option_changed())
        self.bind(remove_strikethrough=lambda instance, value: self._option_changed())
        self.bind(remove_highlight=lambda instance, value: self._option_changed())
        self.bind(remove_links=lambda instance, value: self._option_changed())

    def _option_changed(self):
        if self.auto_process:
            self.process_markdown()

    def paste_from_clipboard(self):
        try:
            self.ids.input_area.text = pyperclip.paste().strip()
        except Exception as e:
            self.ids.output_area.text = f"剪贴板错误: {str(e)}"

    def auto_process_and_update(self):
        if self.auto_process:
            self.process_markdown()

    def process_markdown(self):
        start_time = time.perf_counter()
        try:
            text = self.ids.input_area.text

            # 移除 Markdown 标题
            text = re.sub(r'^#+\s*', '', text, flags=re.MULTILINE)

            # 移除加粗语法 **text**
            text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)

            # 根据选项去除斜体（处理 *text* 和 _text_ ）
            if self.remove_italic:
                text = re.sub(r'(?<!\*)\*(?!\*)(.*?)\*(?!\*)', r'\1', text)
                text = re.sub(r'(?<!_)_(?!_)(.*?)_(?!_)', r'\1', text)

            # 根据选项去除删除线语法 ~~text~~
            if self.remove_strikethrough:
                text = re.sub(r'~~(.*?)~~', r'\1', text)

            # 根据选项去除高亮语法（例如 ==text==）
            if self.remove_highlight:
                text = re.sub(r'==(.+?)==', r'\1', text)

            # 根据选项去除链接，匹配 Markdown 格式链接 [文本](链接)
            if self.remove_links:
                text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)

            # 默认去除 Markdown 分割线（如 ---、***、___ 独占一行的情况）
            text = re.sub(r'(?m)^(?:\s*[-*_]{3,}\s*)$', '', text)

            self.ids.output_area.text = text.strip()

            # 更新统计信息
            char_count = len(self.ids.input_area.text)
            self.ids.char_count.text = f'字符数: {char_count}'
            process_time = (time.perf_counter() - start_time) * 1000
            self.ids.process_time.text = f'处理耗时: {process_time:.2f}ms'
        except Exception as e:
            self.ids.output_area.text = f"处理错误: {str(e)}"

    def copy_to_clipboard(self):
        try:
            pyperclip.copy(self.ids.output_area.text)
        except Exception as e:
            self.ids.output_area.text = f"复制失败: {str(e)}"

    def process_reset(self, target):
        getattr(self.ids, f"{target}_area").text = ''


class MarkdownApp(App):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.tray_icon = None
        self.is_running = True

    def build(self):
        Window.size = (800, 500)  # 优化窗口大小
        Window.bind(on_request_close=self.on_request_close)
        self.setup_tray_icon()
        self.register_hotkey()
        self.title = 'mdword'
        return MarkdownTool()

    def setup_tray_icon(self):
        def create_image():
            from PIL import Image, ImageDraw
            image = Image.new('RGB', (64, 64), 'white')
            dc = ImageDraw.Draw(image)
            dc.rectangle((16, 16, 48, 48), fill='black')
            return image

        import webbrowser
        menu = Menu(
            MenuItem('打开主界面', lambda: self.restore_window()),
            MenuItem('退出', lambda: self.stop_app()),
            MenuItem('关于项目', lambda: webbrowser.open('https://github.com/fhyxz1/mdword'))
        )
        self.tray_icon = Icon(
            'mdword',
            create_image(),
            menu=menu,
            title="mdword\nN+M快速启动"  # 悬浮提示文字
        )
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def register_hotkey(self):
        def toggle_window():
            def _toggle(dt):
                # 使用 Window.visible 判断窗口是否可见
                if Window.visible:
                    Window.hide()
                else:
                    Window.show()
                    Window.raise_window()
            Clock.schedule_once(_toggle)
        keyboard.add_hotkey('N+M',Clock.schedule_once(lambda dt: (Window.show(), Window.raise_window())))

    def on_request_close(self, *args):
        self.show_confirmation()
        return True  # 阻止默认关闭行为

    def show_confirmation(self):
        content = BoxLayout(orientation='vertical', spacing=10)
        popup = Popup(title='操作确认', content=content, size_hint=(0.6, 0.3))

        btn_layout = BoxLayout(spacing=10, size_hint_y=0.5)
        btn_min = Button(text='最小化到托盘', on_press=lambda x: self.minimize_app(popup))
        btn_exit = Button(text='退出程序', on_press=lambda x: self.exit_app(popup))

        content.add_widget(Label(text='请选择要执行的操作:'))
        btn_layout.add_widget(btn_min)
        btn_layout.add_widget(btn_exit)
        content.add_widget(btn_layout)

        popup.open()

    def minimize_app(self, popup):
        popup.dismiss()
        Clock.schedule_once(lambda dt: Window.hide())

    def exit_app(self, popup):
        popup.dismiss()
        self.stop_app()

    def restore_window(self):
        Clock.schedule_once(lambda dt: (Window.show(), Window.raise_window()))

    def stop_app(self):
        self.is_running = False
        if self.tray_icon:
            self.tray_icon.stop()
        Window.close()
        App.get_running_app().stop()
        os._exit(0)  # 确保完全退出

if __name__ == '__main__':
    MarkdownApp().run()
