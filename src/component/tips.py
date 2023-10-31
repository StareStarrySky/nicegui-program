from nicegui import ui


class Tips:
    def __init__(self, msg: str = ''):
        with ui.dialog() as dialog, ui.card().classes('w-2/5'):
            with ui.row():
                ui.label('提示').classes('text-lg')
            ui.separator()
            with ui.row():
                ui.label(msg)
            with ui.row().classes('w-full justify-end'):
                ui.button('确定', on_click=dialog.close)
        dialog.open()
