def get_date_range(period):
    """Calculate start and end dates based on period"""
    now = datetime.now()

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
    elif period == "fromnow":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=365) - timedelta(seconds=1)
    else:
        raise ValueError(f"Unknown period: {period}")

    return start, end


def check_if_installed():
    """Check if Outlook is installed"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return True
    except:
        return False


def get_meetings(
    start_date,
    end_date,
    subject_filter=None,
    organizer_filter=None,
    attendee_filter=None,
):
    try:
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)

        # Format dates for Outlook filter
        start_str = start_date.strftime("%m/%d/%Y %H:%M %p")
        end_str = end_date.strftime("%m/%d/%Y %H:%M %p")

        filter_str = f"[Start] >= '{start_str}' AND [End] <= '{end_str}'"
        appointments = calendar.Items.Restrict(filter_str)
        appointments.Sort("[Start]")

        meetings = []

        for appointment in appointments:
            try:
                # Apply filters
                if (
                    subject_filter
                    and subject_filter.lower() not in appointment.Subject.lower()
                ):
                    continue

                if (
                    organizer_filter
                    and organizer_filter.lower() not in appointment.Organizer.lower()
                ):
                    continue

                if attendee_filter:
                    attendees = appointment.RequiredAttendees or ""
                    if attendee_filter.lower() not in attendees.lower():
                        continue

                meetings.append(
                    {
                        "subject": appointment.Subject,
                        "start": appointment.Start,
                        "end": appointment.End,
                        "organizer": appointment.Organizer,
                        "required_attendees": appointment.RequiredAttendees,
                        "location": appointment.Location,
                        "body": appointment.Body,
                        "is_recurring": appointment.IsRecurring,
                    }
                )
                return meetings
            except:
                pass
    except:
        pass


def get_agenda(
    period: str = "today", subject: str = "", organizer: str = "", attendee: str = ""
):
    start_date, end_date = get_date_range(args.period)

    meetings = get_meetings(
        start_date, end_date, args.subject, args.organizer, args.attendee
    )

    # Filter out past meetings unless argument set
    if not args.past:
        if args.period:
            local_tz = TimeZoneInfo.local()
            now = datetime.now(local_tz)
            meetings = [m for m in meetings if m["end"] >= now]

    print(
        f"\nMeetings from {start_date.strftime('%Y-%m-%d %H:%M')} to {end_date.strftime('%Y-%m-%d %H:%M')}:"
    )
    if args.subject:
        print(f"Subject filter: '{args.subject}'")
    if args.organizer:
        print(f"Organizer filter: '{args.organizer}'")
    if args.attendee:
        print(f"Attendee filter: '{args.attendee}'")
    print("-" * 70)

    if not meetings:
        print("No meetings found matching criteria")
        return

    for idx, meeting in enumerate(meetings, 1):
        print(f"Meeting {idx}:")
        print(f"Subject: {meeting['subject']}")
        print(f"Start: {meeting['start'].strftime('%Y-%m-%d %H:%M')} (Local Time)")
        print(f"End: {meeting['end'].strftime('%Y-%m-%d %H:%M')} (Local Time)")
        print(f"Organizer: {meeting['organizer']}")
        print(f"Attendees: {meeting['required_attendees']}")
        print(f"Location: {meeting['location']}")
        print("-" * 70)
