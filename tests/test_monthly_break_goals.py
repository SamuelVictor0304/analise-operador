"""Financial break impact must use matching monthly goals and filtered amounts."""
import ast
from pathlib import Path
import unittest
import unicodedata

import numpy as np
import pandas as pd


def break_namespace():
    source = Path(__file__).resolve().parents[1] / "dashboard_operadores.py"
    tree = ast.parse(source.read_text(encoding="utf-8-sig"))
    names = {
        "normalize_text", "normalize_status", "normalize_month_key",
        "selected_months", "result_months_frame", "build_monthly_break_goal_analysis",
        "load_office_goals", "money_fmt", "num_fmt", "pct_fmt", "monthly_break_goal_section",
    }
    nodes = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names]
    nodes += [node for node in tree.body if isinstance(node, ast.Assign)
              and any(isinstance(target, ast.Name) and target.id == "MONTH_NAMES_PT" for target in node.targets)]
    for node in nodes:
        if isinstance(node, ast.FunctionDef):
            node.decorator_list = []
    namespace = dict(pd=pd, np=np, unicodedata=unicodedata, RESULTADOS_FILE=None, file_version=lambda _: None)
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(source), "exec"), namespace)
    return namespace


class MonthlyBreakGoalsTest(unittest.TestCase):
    def setUp(self):
        self.ns = break_namespace()
        self.ns["load_office_goals"] = lambda _: {"JULHO": 1000, "AGOSTO": 2000, "SETEMBRO": 0}
        self.results = pd.DataFrame([
            {"MES_RESULTADO": "AGOSTO", "MES_NUM": 8, "VALOR_NAO_PAGOU": 250, "IS_NAO_PAGOU": True},
            {"MES_RESULTADO": "AGOSTO", "MES_NUM": 8, "VALOR_NAO_PAGOU": 0, "IS_NAO_PAGOU": False},
            {"MES_RESULTADO": "JULHO", "MES_NUM": 7, "VALOR_NAO_PAGOU": 500, "IS_NAO_PAGOU": True},
        ])

    def test_each_month_uses_its_own_goal_in_chronological_order(self):
        rows = self.ns["build_monthly_break_goal_analysis"](self.results)
        self.assertEqual(rows["MES_RESULTADO"].tolist(), ["JULHO", "AGOSTO"])
        self.assertEqual(rows["pct_quebra"].tolist(), [.5, .125])
        self.assertEqual(rows["acordos_nao_pagou"].tolist(), [1, 1])

    def test_filtered_amounts_keep_full_office_goal(self):
        rows = self.ns["build_monthly_break_goal_analysis"](self.results.iloc[1:2])
        self.assertEqual(rows.iloc[0]["valor_quebra"], 0)
        self.assertEqual(rows.iloc[0]["meta_geral_escritorio"], 2000)
        self.assertEqual(rows.iloc[0]["pct_quebra"], 0)

    def test_zero_and_missing_goals_are_unavailable(self):
        rows = self.ns["build_monthly_break_goal_analysis"](self.results, ["SETEMBRO", "OUTUBRO"])
        self.assertTrue(rows["pct_quebra"].isna().all())
        self.assertEqual(rows.iloc[0]["meta_geral_escritorio"], 0)
        self.assertTrue(pd.isna(rows.iloc[1]["meta_geral_escritorio"]))

    def test_empty_filter_keeps_selected_month_and_deduplicates(self):
        rows = self.ns["build_monthly_break_goal_analysis"](self.results.iloc[:0], [" agosto ", "AGOSTO"])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows.iloc[0]["pct_quebra"], 0)
        self.assertEqual(rows.iloc[0]["acordos_nao_pagou"], 0)

    def test_no_results_or_selected_months_is_empty(self):
        self.assertTrue(self.ns["build_monthly_break_goal_analysis"](self.results.iloc[:0]).empty)

    def test_unused_categories_do_not_create_extra_months(self):
        self.results["MES_RESULTADO"] = pd.Categorical(self.results["MES_RESULTADO"], categories=["JULHO", "AGOSTO", "SETEMBRO"])
        rows = self.ns["build_monthly_break_goal_analysis"](self.results, ["AGOSTO"])
        self.assertEqual(rows["MES_RESULTADO"].tolist(), ["AGOSTO"])
        self.assertEqual(rows.iloc[0]["pct_quebra"], .125)


if __name__ == "__main__":
    unittest.main()
