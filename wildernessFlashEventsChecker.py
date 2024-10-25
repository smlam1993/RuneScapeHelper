from datetime import datetime, timezone, timedelta
import math
import xlsxwriter

seconds_in_minute = 60
rotation_start_time = datetime(2024, 2, 5, 7, 0, 0, 0, timezone.utc)
rotation_interval = 60
rotation_count = 14

events = [
    "Spider Swarm",
    "Unnatural Outcrop",
    "Stryke the Wyrm",
    "Demon Stragglers",
    "Butterfly Swarm",
    "King Black Dragon Rampage",
    "Forgotten Soldiers",
    "Surprising Seedlings",
    "Hellhound Pack",
    "Infernal Star",
    "Lost Souls",
    "Ramokee Incursion",
    "Displaced Energy",
    "Evil Bloodwood Tree",
]


def to_minutes(seconds):
    return seconds / seconds_in_minute


rotation_mins_offset = -to_minutes(rotation_start_time.timestamp())


def rotation_minutes(interval, rotation_count, offset, timestamp):
    minutes_after_utc = to_minutes(timestamp)
    minutes_into_period = (minutes_after_utc + offset) % (
        rotation_interval * rotation_count
    )

    rotation = math.floor(minutes_into_period / interval) + 1
    minutes_until_next_rotation = interval - minutes_into_period % interval
    return rotation, minutes_until_next_rotation


def hour_rounder(t):
    return t.replace(second=0, microsecond=0, minute=0, hour=t.hour) + timedelta(
        hours=t.minute // 30
    )


def get_list(rotation_names, interval, offset, timestamp, length):
    total_rots = len(rotation_names)
    curr_rot, next = rotation_minutes(interval, total_rots, offset, timestamp)
    curr_time = datetime.fromtimestamp(timestamp, timezone.utc)
    print(curr_time)

    times = []

    for i in range(length):
        i_rot = (curr_rot + i) % total_rots

        name = rotation_names[i_rot]
        next_date_time = curr_time + timedelta(hours=i)
        t = hour_rounder(next_date_time).strftime("%m/%d/%Y %H:%M:%S")
        times.append((name, t))

    return times


special_events = [
    "Stryke the Wyrm",
    "King Black Dragon Rampage",
    "Infernal Star",
    "Evil Bloodwood Tree",
]

skilling_events = [
    "Unnatural Outcrop",
    "Butterfly Swarm",
    "Surprising Seedlings",
    "Infernal Star",
    "Displaced Energy",
]

combat_events = [
    "Spider Swarm",
    "Stryke the Wyrm",
    "Demon Stragglers",
    "King Black Dragon Rampage",
    "Forgotten Soldiers",
    "Hellhound Pack",
    "Infernal Star",
    "Lost Souls",
    "Ramokee Incursion",
]

pain_events = [
    "Forgotten Soldiers",
    "Ramokee Incursion",
]


def generate_spreadsheet(times):
    i_events = {}
    for x in times:
        if x[0] not in i_events:
            i_events[x[0]] = []

        i_events[x[0]].append(x[1])

    workbook = xlsxwriter.Workbook("wilderness_flash_events.xlsx")

    header_format = workbook.add_format()
    header_format.set_bold()
    header_format.set_bg_color("#2A2539")
    header_format.set_font_color("red")
    header_format.set_font_size(14)

    event_cell_format = workbook.add_format()
    event_cell_format.set_bg_color("blue")

    skilling_event_format = workbook.add_format()
    skilling_event_format.set_bg_color("#6E718C")

    combat_events_format = workbook.add_format()
    combat_events_format.set_bg_color("#C9B29F")

    pain_events_format = workbook.add_format()
    pain_events_format.set_italic()
    pain_events_format.set_bg_color("#C9B29F")

    special_event_cell_format = workbook.add_format()
    special_event_cell_format.set_bg_color("#A698AA")
    special_event_cell_format.set_bold()
    special_event_cell_format.set_underline()

    time_cell_format = workbook.add_format()
    time_cell_format.set_bg_color("#BEB1AE")

    worksheet = workbook.add_worksheet()

    row_idx = 0
    worksheet.write(row_idx, 0, "Flash Event Name", header_format)

    for x in i_events:
        row_idx += 1
        column_idx = 0
        row_format = event_cell_format
        if x in special_events:
            row_format = special_event_cell_format
        elif x in skilling_events:
            row_format = skilling_event_format
        elif x in pain_events:
            row_format = pain_events_format
        elif x in combat_events:
            row_format = combat_events_format

        worksheet.set_row(row_idx, 20, row_format)
        worksheet.write(row_idx, column_idx, x)
        for y in i_events[x]:
            column_idx += 1
            worksheet.write(row_idx, column_idx, y, time_cell_format)

        worksheet.set_column_pixels(1, column_idx, 140)

    worksheet.autofit()
    workbook.close()


def dothis():
    # timestamp = datetime.now(timezone.utc).timestamp()
    length = 500
    timestamp = datetime(2024, 10, 27, 23, 59, 59, 59, timezone.utc).timestamp()
    times = get_list(events, rotation_interval, rotation_mins_offset, timestamp, length)
    generate_spreadsheet(times)


dothis()
