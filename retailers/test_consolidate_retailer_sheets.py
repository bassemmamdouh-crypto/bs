import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from consolidate_retailer_sheets import (
    consolidate,
    extract_sheet_rows,
    parse_sheet_name,
)


class ParseSheetNameTests(unittest.TestCase):
    def test_agent_and_supervisor_from_brackets(self):
        agent, supervisor = parse_sheet_name("أحمد علي (خالد حسن)")
        self.assertEqual(agent, "أحمد علي")
        self.assertEqual(supervisor, "خالد حسن")

    def test_spaces_around_brackets(self):
        agent, supervisor = parse_sheet_name("  Sara Nabil(Omar Fadi)  ")
        self.assertEqual(agent, "Sara Nabil")
        self.assertEqual(supervisor, "Omar Fadi")

    def test_no_brackets_uses_full_tab_name(self):
        agent, supervisor = parse_sheet_name("No Supervisor Tab")
        self.assertEqual(agent, "No Supervisor Tab")
        self.assertEqual(supervisor, "")

    def test_csv_file_stem_is_not_treated_as_agent(self):
        agent, supervisor = parse_sheet_name("Massabeh_Retailers_Ameen_Data_2efc")
        self.assertEqual(agent, "")
        self.assertEqual(supervisor, "")


class ExtractSheetRowsTests(unittest.TestCase):
    def test_maps_arabic_headers(self):
        frame = pd.DataFrame(
            {
                "كود": [2777, 2795],
                "رقم الهاتف": ["7901791754", "7705106480"],
                "الاسم التجاري": ["سوق الرافدين", "سوق بغداد"],
            }
        )
        out = extract_sheet_rows(frame, "محمد جاسم (علي كريم)")
        self.assertEqual(list(out.columns), [
            "retailer_id",
            "mobile_number",
            "market_name",
            "agent_name",
            "supervisor_name",
        ])
        self.assertEqual(out.iloc[0].to_dict(), {
            "retailer_id": "2777",
            "mobile_number": "7901791754",
            "market_name": "سوق الرافدين",
            "agent_name": "محمد جاسم",
            "supervisor_name": "علي كريم",
        })
        self.assertEqual(len(out), 2)

    def test_infers_columns_when_headers_are_unreadable(self):
        frame = pd.DataFrame(
            {
                "#": [1, 2],
                "??? ??????": ["Market A BM", "Market B BM"],
                "???????": ["area", "area"],
                "??? ?????? ": ["7901791754", "7705106480"],
                "???": [2777, 2795],
            }
        )
        out = extract_sheet_rows(frame, "Agent One (Supervisor Two)")
        self.assertEqual(out["retailer_id"].tolist(), ["2777", "2795"])
        self.assertEqual(out["mobile_number"].tolist(), ["7901791754", "7705106480"])
        self.assertEqual(out["market_name"].tolist(), ["Market A BM", "Market B BM"])
        self.assertEqual(out["agent_name"].unique().tolist(), ["Agent One"])
        self.assertEqual(out["supervisor_name"].unique().tolist(), ["Supervisor Two"])


class ConsolidateWorkbookTests(unittest.TestCase):
    def test_concatenates_all_tabs(self):
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "agents.xlsx"
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                pd.DataFrame(
                    {
                        "كود": [1001],
                        "رقم الهاتف": ["7700000001"],
                        "الاسم التجاري": ["Market 1"],
                    }
                ).to_excel(writer, index=False, sheet_name="ليث أحمد (سامي يوسف)")
                pd.DataFrame(
                    {
                        "كود": [1002],
                        "رقم الهاتف": ["7700000002"],
                        "الاسم التجاري": ["Market 2"],
                    }
                ).to_excel(writer, index=False, sheet_name="نور علي (سامي يوسف)")
            out = consolidate(path)
            self.assertEqual(len(out), 2)
            self.assertEqual(sorted(out["agent_name"].tolist()), ["ليث أحمد", "نور علي"])
            self.assertEqual(out["supervisor_name"].unique().tolist(), ["سامي يوسف"])
            self.assertEqual(sorted(out["retailer_id"].tolist()), ["1001", "1002"])


if __name__ == "__main__":
    unittest.main()
