from flogin import Plugin


class OutlookAgendaPlugin(Plugin):
    def __init__(self) -> None:
        super().__init__()

        from .handlers.get import GetOutlookAgenda

        self.register_search_handlers(
            GetOutlookAgenda(),
        )
