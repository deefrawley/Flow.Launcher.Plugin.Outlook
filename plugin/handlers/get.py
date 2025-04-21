from __future__ import annotations

from flogin import PlainTextCondition, Query, Result, SearchHandler

from ..plugin import OutlookAgendaPlugin

import win32com.client
import pywintypes
from datetime import datetime, timedelta


class GetOutlookAgenda(SearchHandler[OutlookAgendaPlugin]):
    def __init__(self):
        super().__init__(PlainTextCondition(""))

    async def callback(self, query: Query):
        assert self.plugin

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            if query.text == "":
                yield Result(
                    title=f"{query.keyword} today, tomorrow, week, month, custom (from date, end date in YYYY-MM-DD [HH:MM] format)",
                    sub="Default is today, unless after 5pm then tomorrow",
                    icon="assets/app.png",
                )
            else:
                start_date, end_date = self._get_date_range("today")
                yield Result(
                    title=f"Start date - {start_date}, End date - {end_date}",
                    sub="Default is today, unless after 5pm then tomorrow",
                    icon="assets/app.png",
                )
        except pywintypes.com_error as e:
            yield Result(
                title=f"Error - {e.hresult}",
                sub="This plugin won't work (most likely reason is Outlook not installed)",
                icon="assets/app.png",
            )

    def _get_date_range(self, period):
        """Calculate start and end dates based on period"""
        now = datetime.now()

        period = period.lower()
        if period == "today":
            start = now.replace(hour=0, minute=0, second=0, microsecond=0)
            end = start + timedelta(days=1) - timedelta(seconds=1)
        elif period == "tomorrow":
            start = (now + timedelta(days=1)).replace(
                hour=0, minute=0, second=0, microsecond=0
            )
            end = start + timedelta(days=1) - timedelta(seconds=1)
        elif period == "week":
            start = now - timedelta(days=now.weekday())
            start = start.replace(hour=0, minute=0, second=0, microsecond=0)
            end = start + timedelta(days=6, hours=23, minutes=59, seconds=59)
        elif period == "month":
            start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            if now.month == 12:
                end = start.replace(year=start.year + 1, month=1)
            else:
                end = start.replace(month=start.month + 1)
            end = end - timedelta(seconds=1)

        return start, end
