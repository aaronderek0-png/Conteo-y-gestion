import json, os, sys, tempfile, unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


class TestSellos(unittest.TestCase):
    def test_variantes_sucias(self):
        import invoice_parser as ip
        for t, esperado in (("CERV. RECIBIDO", "Congreso"), ("NUÑEZ", "Chimidoha"), ("MISHIGUENE CAB", "Cabello"),
                            ("mishiguene cerv - recibido", "Congreso")):
            self.assertEqual(ip.detectar_local(sello=t)["local"], esperado, t)

    def test_sin_sello_sigue_indeterminado(self):
        import invoice_parser as ip
        self.assertIsNone(ip.detectar_local(sello=None, razon_social="LAFINUR 3368 SRL")["local"])
        self.assertIsNone(ip.local_por_patron_sello("FACTURA A ORIGINAL")[0])


class TestGrupos(unittest.TestCase):
    def test_vincular_y_archivo(self):
        import grupos_telegram as g
        with tempfile.TemporaryDirectory() as d:
            g.ARCHIVO = Path(d) / "chats_locales.json"
            self.assertEqual(g.vincular("-100123456", "cabello"), "Cabello")
            self.assertEqual(g.local_de_archivo("-100123456_77_tg_x.jpg"), "Cabello")
            self.assertIsNone(g.local_de_archivo("-999999999_1_tg_x.jpg"))
            with self.assertRaises(ValueError):
                g.vincular("-1", "Narnia")


class TestROI(unittest.TestCase):
    def test_modelo(self):
        import telemetria_roi as t
        # 45 s base + 5 renglones * 25 s = 170 s
        self.assertAlmostEqual(t.minutos_modelo("Factura", 5), round(170 / 60, 2))
        self.assertAlmostEqual(t.minutos_modelo("Remito", 2), round((45 + 50 + 60) / 60, 2))
        self.assertAlmostEqual(t.minutos_modelo("NC", 1), round((45 + 25 + 60) / 60, 2))

    def test_registro_idempotente(self):
        import telemetria_roi as t
        with tempfile.TemporaryDirectory() as d:
            t.ARCHIVO = Path(d) / "t.json"
            c = {"tipo": "Factura", "sha256": "x", "numero_factura": "1", "proveedor_detectado": "P", "archivo": "a.jpg",
                 "fecha_factura": "2026-10-01", "items": [{}, {}]}
            t.registrar_lote([c], {}, {"push_drive": {}}, 3.0)
            t.registrar_lote([c], {}, {"push_drive": {}}, 3.0)
            r = t.resumen()
            self.assertEqual(r["comprobantes_procesados"], 1)
            self.assertEqual(r["renglones_totales"], 2)
            self.assertEqual(r["seg_maquina_promedio"], 3.0)


if __name__ == "__main__":
    unittest.main()
