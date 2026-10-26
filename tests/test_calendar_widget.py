import unittest
from datetime import date

from calendar_widget import (
    COLORS,
    day_visual_state,
    group_events_by_date,
    is_overdue_event,
    month_grid,
)


class CalendarHelperTests(unittest.TestCase):
    def test_month_grid_is_six_monday_first_weeks(self) -> None:
        grid = month_grid(2026, 10)
        self.assertEqual(len(grid), 6)
        self.assertTrue(all(len(week) == 7 for week in grid))
        self.assertEqual(grid[0][0], date(2026, 9, 28))
        self.assertEqual(grid[0][0].weekday(), 0)
        self.assertEqual(grid[-1][-1], date(2026, 11, 8))
        self.assertIn(date(2026, 10, 26), [day for week in grid for day in week])

    def test_day_color_priority_and_today_border(self) -> None:
        today = date(2026, 10, 26)
        ordinary = day_visual_state(date(2026, 10, 27), [], today=today)
        self.assertEqual(ordinary["background"], COLORS["empty"])

        today_state = day_visual_state(today, [], today=today)
        self.assertEqual(today_state["background"], COLORS["today"])
        self.assertTrue(today_state["is_today"])

        event = {"event_date": "2026-10-27", "status": "Scheduled"}
        event_state = day_visual_state(date(2026, 10, 27), [event], today=today)
        self.assertEqual(event_state["background"], COLORS["event"])
        self.assertEqual(event_state["event_count"], 1)

        overdue_event = {"event_date": "2026-10-20", "status": "Scheduled"}
        overdue_state = day_visual_state(date(2026, 10, 20), [overdue_event], today=today)
        self.assertEqual(overdue_state["background"], COLORS["overdue"])
        self.assertTrue(overdue_state["overdue"])

        today_with_event = day_visual_state(today, [{"event_date": today.isoformat(), "status": "Scheduled"}], today=today)
        self.assertEqual(today_with_event["background"], COLORS["event"])
        self.assertEqual(today_with_event["border"], COLORS["today_border"])

    def test_overdue_status_and_event_grouping(self) -> None:
        today = date(2026, 10, 26)
        self.assertTrue(is_overdue_event({"event_date": "2026-10-25", "status": "Scheduled"}, today))
        self.assertTrue(is_overdue_event({"event_date": "2026-10-25", "status": "Overdue"}, today))
        self.assertFalse(is_overdue_event({"event_date": "2026-10-25", "status": "Completed"}, today))
        self.assertFalse(is_overdue_event({"event_date": "2026-10-27", "status": "Scheduled"}, today))
        grouped = group_events_by_date([
            {"event_date": "2026-10-26", "status": "Scheduled"},
            {"event_date": "26/10/2026", "status": "Completed"},
            {"event_date": None, "status": "Scheduled"},
        ])
        self.assertEqual(len(grouped["2026-10-26"]), 2)


if __name__ == "__main__":
    unittest.main()
