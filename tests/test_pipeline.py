"""Fast tests that do not require re-downloading the 233 MiB input archive."""
import json
import math
import unittest
import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from src.analysis import (KEY,FOOD7,FOOD30,CSQ30,CSQ365,DGQ365,
                          CODE_BUCKET,ALL_CODES,EXCLUDE,MONTH_FACTOR)
ROOT=Path(__file__).resolve().parents[1]
class TestHCESPipeline(unittest.TestCase):
    def test_categories_use_subtotals_only(self):
        self.assertEqual(CODE_BUCKET[129],'staples')
        self.assertEqual(CODE_BUCKET[199],'protein')
        self.assertEqual(CODE_BUCKET[289],'processed')
        self.assertEqual(CODE_BUCKET[349],'fuel')
        self.assertEqual(CODE_BUCKET[499],'services')
        self.assertIn(539,EXCLUDE)
    def test_recall_period_factors(self):
        self.assertEqual(MONTH_FACTOR[129],1)
        self.assertAlmostEqual(MONTH_FACTOR[199],30/7)
        self.assertAlmostEqual(MONTH_FACTOR[409],30/365)
        self.assertAlmostEqual(MONTH_FACTOR[379],30/365)
        self.assertEqual(MONTH_FACTOR[349],1)
    def test_distinct_code_sets(self):
        coll=[FOOD7,FOOD30,CSQ30,CSQ365,DGQ365]
        for i,a in enumerate(coll):
            for b in coll[i+1:]:self.assertFalse(a & b)
        self.assertEqual(set().union(*coll),ALL_CODES)
    def test_join_id_has_required_household_fields(self):
        self.assertEqual(len(KEY),6)
        self.assertIn('fsu',KEY)
        self.assertIn('b1q1pt12',KEY)
    def test_saved_outputs_match_calculations(self):
        s=json.loads((ROOT/'results/run_summary.json').read_text())
        self.assertEqual(s['data']['eligible_households'],261744)
        self.assertEqual(sum(s['cluster_counts'].values()),s['data']['eligible_households'])
        self.assertEqual(s['best_k'],5)
        self.assertAlmostEqual(s['best_silhouette'],0.2216,places=4)
        self.assertLess(s['ablation_excluding_processed_share_ari'],0.5)
        self.assertEqual(len(list((ROOT/'results/figures').glob('*.png'))),12)
if __name__=='__main__':unittest.main()
