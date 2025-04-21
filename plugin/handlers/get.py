from __future__ import annotations

from flogin import PlainTextCondition, Query, Result, SearchHandler

from ..plugin import OutlookAgendaPlugin

class GetOutlookAgenda(SearchHandler[OutlookAgendaPlugin]):
    def __init__(self):
        super().__init__(PlainTextCondition(""))

    async def callback(self, query: Query):
        assert self.plugin

        yield Result(
                title="Outlook Agenda",
                sub="Gets agenda",
                icon="assets/app.png"
                ),
        )