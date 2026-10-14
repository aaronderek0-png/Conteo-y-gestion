import os
import tempfile as _tf
os.environ.setdefault('MISHI_ESTADO_DIR', _tf.mkdtemp())
os.environ.setdefault('MISHI_LOG_DIR', os.environ['MISHI_ESTADO_DIR'])
import tempfile
import unittest
from pathlib import Path

import coordinador as co
import conocimiento as kn


class TestCola(unittest.TestCase):
    def setUp(self):
        self.d = Path(tempfile.mkdtemp())
        self.c = co.Cola(self.d / "trabajos.json")

    def test_idempotente_y_error_no_consume(self):
        self.c.registrar("abc", "a.jpg")
        self.c.registrar("abc", "a.jpg")
        self.assertEqual(self.c.resumen()["RECIBIDO"], 1)
        self.c.transicion("abc", "LEIDO")
        self.c.transicion("abc", "ERROR", error="APIError 429")
        j = self.c.obtener("abc")
        self.assertEqual((j["estado"], j["intentos"]), ("ERROR", 1))
        self.assertEqual(len(self.c.pendientes()), 1)          # sigue en cola
        self.c.transicion("abc", "VALIDADO")
        self.c.transicion("abc", "ESCRITO", {"rango": "A10:I12"})
        self.assertEqual(self.c.pendientes(), [])
        with self.assertRaises(ValueError):
            self.c.transicion("abc", "LEIDO")                  # ESCRITO no se reabre

    def test_no_salta_a_escrito_desde_recibido(self):
        self.c.registrar("x")
        with self.assertRaises(ValueError):
            self.c.transicion("x", "ESCRITO")

    def test_un_solo_escritor_por_destino(self):
        with co._bloqueo("destino_T", timeout=0.3, base=self.d):
            with self.assertRaises(co.Bloqueado):
                with co._bloqueo("destino_T", timeout=0.3, base=self.d):
                    pass
        with co._bloqueo("destino_T", timeout=0.3, base=self.d):    # liberado
            pass


class TestConocimiento(unittest.TestCase):
    def setUp(self):
        self.r = Path(tempfile.mkdtemp()) / "estado.json"

    def test_invalidacion_por_dependencias(self):
        kn.registrar_fuente("precios_vigentes", "h1", self.r)
        kn.marcar_calculado("costos_bom", {"precios_vigentes": "h1"}, ruta=self.r)
        self.assertTrue(kn.es_valido("costos_bom", self.r))
        inval = kn.registrar_fuente("precios_vigentes", "h2", self.r)
        self.assertEqual(inval, ["filas_factura", "costos_bom", "gestion_margenes", "comparativa_precios", "reporte_web"])
        self.assertFalse(kn.es_valido("costos_bom", self.r))

    def test_inferencia_no_pisa_validado(self):
        f = Path(tempfile.mkdtemp()) / "f.json"
        kn.registrar_dato("wigen.destino", "Congreso", "validado", "foto", ruta=f)
        r = kn.registrar_dato("wigen.destino", "Lafinur", "inferencia", "ocr", ruta=f)
        self.assertFalse(r["aceptado"])


if __name__ == "__main__":
    unittest.main()
