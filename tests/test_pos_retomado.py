"""Conversão por retomada: população judicial, datas, parcelas e denominadores."""
import unittest

import pandas as pd

from pos_retomado import prepare_pos_retomado_cases, pos_retomado_summary


def seizure(contract, month="JAN", date="2026-01-02", nature="JUDICIAL", agreement="NÃO", sale="NÃO"):
    return {"CONTRATO": contract, "Mês": month, "DATA DA RETOMADA": date,
            "JUDICIAL / DA": nature, "ACORDO": agreement, "VENDA": sale, "NEGOCIAÇÃO": ""}


def outcome(contract, status="PAGOU", date="2026-01-03", payment=None, campaign="PÓS RETOMADA", value=100):
    return {"Nº CONTRATO": contract, "EMISSÃO": date, "DATA DO PAGAMENTO": payment,
            "STATUS": status, "CAMPANHA": campaign, "VALOR DO BANCO - META": value}


def build(seizures, outcomes=None):
    return prepare_pos_retomado_cases(pd.DataFrame(seizures), pd.DataFrame(outcomes or [], columns=list(outcome(1))))


class PosRetomadoTest(unittest.TestCase):
    def test_da_is_visible_but_excluded_from_judicial_denominators(self):
        cases = build([seizure(1), seizure(2), seizure(3, nature="D.A")], [outcome(1), outcome(3)])
        monthly, statuses, total = pos_retomado_summary(cases)
        self.assertEqual(len(cases), 3)
        self.assertEqual(monthly.iloc[0]["casos"], 2)
        self.assertEqual(monthly.iloc[0]["convertidos"], 1)
        self.assertEqual(monthly.iloc[0]["pct_conversao"], .5)
        self.assertAlmostEqual(statuses["pct_mes"].sum(), 1)
        self.assertAlmostEqual(total["pct_total"].sum(), 1)

    def test_multiple_installments_count_one_case_with_any_payment(self):
        cases = build([seizure(1)], [outcome(1), outcome(1, status="NÃO PAGOU"), outcome(1, status="EM ABERTO")])
        self.assertEqual(cases.iloc[0]["situacao"], "Pago")
        self.assertEqual(cases.iloc[0]["registros_base"], 3)
        self.assertEqual(cases.iloc[0]["valor_pago"], 100)
        self.assertEqual(pos_retomado_summary(cases)[0].iloc[0]["convertidos"], 1)

    def test_monthly_ticket_uses_paid_cases_and_excludes_da_and_unpaid(self):
        cases = build([seizure(1), seizure(2), seizure(3), seizure(4, nature="D.A"),
                       seizure(5, month="FEV", date="2026-02-02")],
                      [outcome(1, value=100), outcome(1, value=300), outcome(2, value=200),
                       outcome(3, status="EM ABERTO", value=1000), outcome(4, value=9000)])
        for converted in ["pago", "tem_acordo"]:
            monthly, _, _ = pos_retomado_summary(cases, converted)
            january, february = monthly.iloc[0], monthly.iloc[1]
            self.assertEqual(january["valor_pago"], 600)
            self.assertEqual(january["pagos"], 2)
            self.assertEqual(january["ticket_medio_pos"], 300)
            self.assertTrue(pd.isna(february["ticket_medio_pos"]))

    def test_only_post_campaign_after_seizure_is_eligible(self):
        cases = build([seizure(1), seizure(2), seizure(3)], [
            outcome(1, date="2026-01-01"), outcome(2, campaign="EXCEÇÃO"), outcome(3)])
        self.assertEqual(cases["pago"].tolist(), [False, False, True])
        self.assertEqual(cases["tem_acordo"].tolist(), [False, False, True])

    def test_later_month_payment_belongs_to_seizure_month(self):
        cases = build([seizure(1, month="FEV", date="2026-02-02"), seizure(2, month="JAN")],
                      [outcome(1, date="2026-03-01", payment="2026-03-02")])
        monthly, _, _ = pos_retomado_summary(cases)
        self.assertEqual(monthly["mes"].tolist(), ["JANEIRO", "FEVEREIRO"])
        self.assertEqual(monthly["pct_conversao"].tolist(), [0, 1])

    def test_agreement_conversion_is_distinct_from_payment(self):
        cases = build([seizure(1, agreement="SIM"), seizure(2), seizure(3)], [outcome(2, status="EM ABERTO")])
        payments = pos_retomado_summary(cases)[0].iloc[0]
        agreements = pos_retomado_summary(cases, "tem_acordo")[0].iloc[0]
        self.assertEqual(payments["convertidos"], 0)
        self.assertEqual(agreements["convertidos"], 1)
        self.assertFalse(cases.iloc[0]["tem_acordo"])
        self.assertEqual(cases.iloc[0]["situacao"], "Sem acordo na BASE")
        self.assertIn("sem pós confirmado na BASE", cases.iloc[0]["conferencia"])
        self.assertIn("sem ACORDO = SIM", cases.iloc[1]["conferencia"])

    def test_apreensao_agreement_flag_does_not_determine_conversion_or_status(self):
        outcomes = [outcome(1), outcome(2, status="NÃO PAGOU")]
        marked = build([seizure(i, agreement="SIM") for i in [1, 2, 3]], outcomes)
        unmarked = build([seizure(i, agreement="NÃO") for i in [1, 2, 3]], outcomes)
        columns = ["tem_acordo", "pago", "situacao", "valor_pago"]
        pd.testing.assert_frame_equal(marked[columns], unmarked[columns])

    def test_missing_date_keeps_case_denominator_without_confirming_payment(self):
        cases = build([seizure(1, date="26/102026", month="OUT")], [outcome(1, date="2026-10-27")])
        self.assertEqual(cases.iloc[0]["mes"], "OUTUBRO")
        self.assertFalse(cases.iloc[0]["pago"])
        self.assertIn("Data de retomada inválida", cases.iloc[0]["conferencia"])
        self.assertEqual(pos_retomado_summary(cases)[0].iloc[0]["casos"], 1)

    def test_month_field_is_preserved_and_conflict_is_flagged(self):
        cases = build([seizure(1, month="ABR", date="2026-03-31")])
        self.assertEqual(cases.iloc[0]["mes"], "ABRIL")
        self.assertIn("Mês diferente", cases.iloc[0]["conferencia"])

    def test_repeated_seizure_does_not_reuse_old_outcome(self):
        cases = build([seizure(1), seizure(1, month="FEV", date="2026-02-02")], [outcome(1)])
        self.assertEqual(cases["pago"].tolist(), [True, False])

    def test_duplicate_seizure_rows_are_not_counted_twice(self):
        cases = build([seizure(1), seizure(1)], [outcome(1)])
        self.assertEqual(len(cases), 1)

    def test_payment_date_and_open_status_are_reconciled(self):
        cases = build([seizure(1), seizure(2)], [
            outcome(1, status="EM ABERTO", payment="2026-01-04"),
            outcome(2, status="EM ABERTO"), outcome(2, status="NÃO PAGOU")])
        self.assertEqual(cases["situacao"].tolist(), ["Pago", "Em aberto"])

    def test_sale_without_agreement_is_not_conversion(self):
        cases = build([seizure(1, sale="SIM")])
        self.assertEqual(cases.iloc[0]["situacao"], "Vendido")
        self.assertFalse(cases.iloc[0]["tem_acordo"])

    def test_empty_and_da_only_filters_have_no_judicial_percentages(self):
        cases = build([seizure(1, nature="D.A")])
        for subset in [cases, cases.iloc[:0]]:
            self.assertTrue(all(frame.empty for frame in pos_retomado_summary(subset)))


if __name__ == "__main__":
    unittest.main()
