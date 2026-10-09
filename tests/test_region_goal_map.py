"""Regression checks without starting the Streamlit app or loading event files."""
import ast
from pathlib import Path
import unittest
import unicodedata

import numpy as np
import pandas as pd


def map_namespace():
    source = Path(__file__).resolve().parents[1] / "dashboard_operadores.py"
    tree = ast.parse(source.read_text(encoding="utf-8-sig"))
    names = {"normalize_text", "normalize_status", "normalize_month_key", "safe_div", "build_region_goal_map",
             "load_office_goals", "load_region_goals", "office_goal_for_months"}
    functions = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names]
    for node in functions:
        node.decorator_list = []
    namespace = dict(pd=pd, np=np, unicodedata=unicodedata, RESULTADOS_FILE=None, file_version=lambda _: None)
    exec(compile(ast.Module(body=functions, type_ignores=[]), str(source), "exec"), namespace)
    return namespace


class RegionGoalMapTest(unittest.TestCase):
    def setUp(self):
        self.ns = map_namespace()
        self.goals = pd.DataFrame([
            {"REGIÃO": "SUL", "REGIAO_KEY": "SUL", "MES_RESULTADO": "SETEMBRO", "meta_regiao": 1000},
            {"REGIÃO": "SUL", "REGIAO_KEY": "SUL", "MES_RESULTADO": "OUTUBRO", "meta_regiao": 0},
        ])
        self.ns["load_region_goals"] = lambda _: self.goals.copy()
        self.results = pd.DataFrame([
            {"REGIÃO": "Sul", "MES_RESULTADO": "OUTUBRO", "VALOR_PAGO": 250.0},
            {"REGIÃO": "Sul", "MES_RESULTADO": "SETEMBRO", "VALOR_PAGO": 900.0},
        ])

    def test_missing_goal_does_not_borrow_another_month(self):
        result = self.ns["build_region_goal_map"](["OUTUBRO"], self.results)
        self.assertTrue(result.empty)

    def test_registered_goal_uses_selected_month(self):
        self.goals.loc[1, "meta_regiao"] = 2000
        result = self.ns["build_region_goal_map"](["OUTUBRO"], self.results)
        self.assertEqual(result.iloc[0]["pct_meta_regiao"], .125)
        self.assertEqual(result.iloc[0]["valor_pago"], 250)

    def test_empty_filtered_results_do_not_invent_receipts(self):
        self.goals.loc[1, "meta_regiao"] = 2000
        result = self.ns["build_region_goal_map"](["OUTUBRO"], self.results.iloc[:0])
        self.assertEqual(result.iloc[0]["valor_pago"], 0)

    def test_multiple_months_sum_their_own_goals_and_receipts(self):
        self.goals.loc[1, "meta_regiao"] = 2000
        result = self.ns["build_region_goal_map"](["SETEMBRO", "OUTUBRO"], self.results)
        self.assertEqual(result.iloc[0]["meta_regiao"], 3000)
        self.assertEqual(result.iloc[0]["valor_pago"], 1150)

    def test_office_goal_preserves_zero_and_deduplicates_months(self):
        self.ns["load_office_goals"] = lambda _: {"SETEMBRO": 1000, "OUTUBRO": 2000, "NOVEMBRO": 0}
        self.assertEqual(self.ns["office_goal_for_months"](["NOVEMBRO"]), 0)
        self.assertEqual(self.ns["office_goal_for_months"](["SETEMBRO", "OUTUBRO", "OUTUBRO"]), 3000)


if __name__ == "__main__":
    unittest.main()
