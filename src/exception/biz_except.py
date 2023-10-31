from component.tips import Tips


class BizExcept(BaseException):
    def __init__(self, msg: str):
        super().__init__(msg)
        Tips(msg)
