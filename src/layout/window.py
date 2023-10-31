import os
from contextlib import contextmanager

from nicegui import ui

from biz.analyse_excel import AnalyseExcel
from biz.excel_with_pic import ExcelWithPic
from component.tips import Tips
from exception.biz_except import BizExcept
from layout.local_file_picker import LocalFilePicker


class Window:
    def __init__(self):
        self.title = 'ExcelWithPic'
        self.file_path = ''
        self.res_path = ''
        self.data_rate = 0.0
        self.data_rate_str = ''

        self.data_num = 0
        self.data_num_str = ''
        self.pic_num_str = ''
        self.pic_downloaded_num_str = ''
        self.pic_compressed_num_str = ''

        self.analyse_file_visible = False
        self.del_file_but_visible = False
        self.success = False

        self.url_prefix = ''
        self.pic_len = 10
        self.pic_width = 1000
        self.pic_height = 1000
        self.compress_rate = 0.8
        self.res_num = 4

        with ui.row().classes('align-items'):
            ui.button('选择文件', on_click=self.pick_file, icon='folder')
            ui.label().bind_text_from(self, 'file_path').classes('self-center')
            ui.button('移除文件',
                      on_click=self.remove_file, icon='remove').bind_visibility_from(self, 'del_file_but_visible')

        with ui.row().classes('align-items').bind_visibility_from(self, 'analyse_file_visible'):
            ui.button('分析', on_click=lambda e: self.analyse_file(e.sender), icon='cached')
            ui.label().bind_text_from(self, 'data_num_str').classes('self-center')
            ui.label().bind_text_from(self, 'pic_num_str').classes('self-center')
            ui.label().bind_text_from(self, 'pic_downloaded_num_str').classes('self-center')
            ui.label().bind_text_from(self, 'pic_compressed_num_str').classes('self-center')

        with ui.row().classes('align-items').bind_visibility_from(self, 'analyse_file_visible'):
            ui.button('删除所有压缩图片', on_click=self.delete_all_compressed_pic, icon='delete')

        with ui.card().classes('w-full'), ui.grid(columns=1).classes('w-full'):
            ui.input('系统地址（固定不改）：', value='http://10.0.0.100/').bind_value_to(self, 'url_prefix')
            ui.number('每个文件条数（小于总条数时会拆分文件）：',
                      value=1000, min=1, format='%.0f').bind_value_to(self, 'res_num', forward=lambda x: int(x))
            ui.number('每张图片最大(kb)（影响文件大小）：',
                      value=10, max=10, min=1, format='%.0f').bind_value_to(self, 'pic_len', forward=lambda x: int(x))
            ui.number('每张图片最宽(px)（影响图片占用的最大宽度）：',
                      value=1000, format='%.0f').bind_value_to(self, 'pic_width', forward=lambda x: int(x))
            ui.number('每张图片最高(px)（影响图片占用的最大高度）：',
                      value=1000, format='%.0f').bind_value_to(self, 'pic_height', forward=lambda x: int(x))
            ui.number('每次压缩比例（小幅影响压缩速度）：', value=0.8, max=1, step=0.01).bind_value_to(self, 'compress_rate')

        with ui.row().classes('w-full justify-end'):
            with ui.row().classes('self-center').bind_visibility_from(self, 'success'):
                ui.label('转换成功，')
                ui.link('打开文件夹', '###').on('click', self.open_dir)
            ui.button('开始转换', on_click=lambda e: self.excel_pic(e.sender), icon='cached')

        with ui.linear_progress(size='20px', show_value=False).bind_value_from(self, 'data_rate'):
            with ui.row().classes('absolute-full flex flex-center'):
                ui.badge(color='white', text_color='dark').bind_text_from(self, 'data_rate_str')

    async def pick_file(self) -> None:
        file_paths = await LocalFilePicker('')

        if len(file_paths) == 0:
            return None

        if not file_paths[0].endswith('.xlsx'):
            raise BizExcept('请选择一个xlsx文件')
        else:
            self.file_path = file_paths[0]
            self.res_path = os.path.dirname(file_paths[0])

        self.analyse_file_visible = True
        self.del_file_but_visible = True
        self.success = False
        self.data_num = 0
        self.data_num_str = '共0条数据'
        self.pic_num_str = '共0张图片'
        self.pic_downloaded_num_str = '已下载0张图片'
        self.pic_compressed_num_str = '已压缩0张图片'

    @contextmanager
    def disable_button(self, button: ui.button) -> None:
        button.disable()
        try:
            yield
        finally:
            button.enable()

    async def analyse_file(self, button: ui.button) -> None:
        with self.disable_button(button):
            async with AnalyseExcel(self.file_path) as analyse:
                self.data_num = await analyse.get_data_num()
                self.data_num_str = '共' + str(self.data_num) + '条数据'
                pic_num = await analyse.get_pic_num()
                self.pic_num_str = '共' + str(pic_num) + '张图片'
                pic_downloaded_num = await analyse.get_pic_downloaded_num()
                self.pic_downloaded_num_str = '已下载' + str(pic_downloaded_num) + '张图片'
                pic_compressed_num = await analyse.get_pic_compressed_num()
                self.pic_compressed_num_str = '已压缩' + str(pic_compressed_num) + '张图片'

    def remove_file(self):
        self.file_path = ''
        self.data_num = 0
        self.data_num_str = ''
        self.pic_num_str = ''
        self.pic_downloaded_num_str = ''
        self.pic_compressed_num_str = ''
        self.analyse_file_visible = False
        self.del_file_but_visible = False
        self.success = False

    def delete_all_compressed_pic(self):
        os.system(f'del /s /q {self.res_path}\\*.compress > nul')
        Tips('删除完成')

    async def excel_pic(self, button: ui.button) -> None:
        with self.disable_button(button):
            async with ExcelWithPic(self.file_path, url_prefix=self.url_prefix, pic_len=self.pic_len,
                                    pic_width=self.pic_width, pic_height=self.pic_height,
                                    compress_rate=self.compress_rate, res_num=self.res_num) as excel_with_pic:
                try:
                    excel_with_pic.set_window(self)
                    await excel_with_pic.run()
                except BizExcept:
                    pass

    def open_dir(self):
        os.system(f'explorer {self.res_path}')

    def run(self):
        # ui.run(title=self.title)
        ui.run(title=self.title, reload=False, native=True, window_size=(800, 800))
