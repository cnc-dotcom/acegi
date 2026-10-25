import unittest
from app import Part, shelf_nest, optimize_lengths

class TestAlgorithms(unittest.TestCase):
    def test_plate_nesting_fits_and_counts(self):
        parts=[Part("A",500,300,2,"10"),Part("B",250,300,1,"10")]
        sheets,util=shelf_nest(parts,1000,1000,5)
        self.assertGreaterEqual(len(sheets),1)
        self.assertEqual(sum(len(s["placements"]) for s in sheets),3)
        self.assertGreater(util,0)
        for s in sheets:
            for p in s["placements"]:
                self.assertGreaterEqual(p["x"],0)
                self.assertGreaterEqual(p["y"],0)
                self.assertLessEqual(p["x"]+p["w"],1000+1e-6)
                self.assertLessEqual(p["y"]+p["h"],1000+1e-6)
    def test_length_optimizer_counts_and_scrap(self):
        bars,util=optimize_lengths([{"name":"A","length":4000,"qty":2},{"name":"B","length":2000,"qty":1}],6000,3,0)
        self.assertEqual(sum(len(b["pieces"]) for b in bars),3)
        self.assertGreater(util,0)
        self.assertTrue(all(b["scrap"]>=0 for b in bars))
    def test_invalid_part_rejected(self):
        with self.assertRaises(ValueError):
            shelf_nest([Part("Bad",0,10,1)],100,100)
    def test_oversize_length_rejected(self):
        with self.assertRaises(ValueError):
            optimize_lengths([{"name":"Long","length":7000,"qty":1}],6000)
if __name__=="__main__": unittest.main()
