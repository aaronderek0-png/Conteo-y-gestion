import json
import os
import tempfile
import unittest
from pathlib import Path

import aprobaciones as ap
import entornos as en


class TestEntornos(unittest.TestCase):
    def setUp(self):
        self.d = Path(tempfile.mkdtemp())
        self.f = self.d / "entornos.json"
        self.f.write_text(json.dumps({"produccion": {"Cabello": "PROD123456"}, "prueba": {"Cabello": "TEST123456"}}))
        os.environ["MISHI_ENTORNOS_JSON"] = str(self.f)
        os.environ["MISHI_ENTORNO"] = "PRUEBA"

    def tearDown(self):
        os.environ.pop("MISHI_ENTORNOS_JSON", None)
        os.environ.pop("MISHI_ENTORNO", None)

    def test_rechaza_produccion_y_desconocidos(self):
        en.validar_destino("TEST123456")
        with self.assertRaises(en.EscrituraProhibida):
            en.validar_destino("PROD123456")
        with self.assertRaises(en.EscrituraProhibida):
            en.validar_destino("OTRO000000")

    def test_no_registra_copia_con_id_de_produccion_o_titulo_ajeno(self):
        with self.assertRaises(en.EscrituraProhibida):
            en.registrar_prueba("Cabello", "PROD123456", en.PREFIJO + "Cabello", True, self.f)
        with self.assertRaises(ValueError):
            en.registrar_prueba("Cabello", "NUEVO", "CABELLO", True, self.f)
        with self.assertRaises(ValueError):
            en.registrar_prueba("Cabello", "NUEVO", en.PREFIJO + "Cabello", False, self.f)

    def test_produccion_no_se_controla_en_modo_produccion(self):
        os.environ["MISHI_ENTORNO"] = "PRODUCCION"
        en.validar_destino("PROD123456")


class TestAprobaciones(unittest.TestCase):
    def _tabla(self, humano, tecnico="", motivo="ok", evid="foto"):
        base = ["ID1", "0001-1", "2026-07-07", "WIGEN", "NUECES", "KG", 1, 1, 0, 0, "", "NO", "PENDIENTE", "", "a.jpg", ""]
        meta = [humano, motivo, evid, "", "", tecnico, "", ""]
        return [["ID REMITO"], base + meta]

    def test_aprobado_no_recarga_si_ya_escrito(self):
        p = ap.planear(ap.leer(self._tabla("APROBADO", "ESCRITO")))
        self.assertEqual(p[0]["accion"], "NADA")

    def test_aprobado_carga_y_denegado_no_toca_gasto(self):
        self.assertEqual(ap.planear(ap.leer(self._tabla("APROBADO")))[0]["accion"], "CARGAR")
        self.assertEqual(ap.planear(ap.leer(self._tabla("DENEGADO")))[0]["accion"], "CERRAR_SIN_EFECTO")
        self.assertEqual(ap.celdas_resultado("CERRAR_SIN_EFECTO", True)["V"], "DENEGADO_CERRADO")

    def test_revision_o_sin_motivo_espera(self):
        self.assertEqual(ap.planear(ap.leer(self._tabla("REVISIÓN")))[0]["accion"], "ESPERAR")
        self.assertEqual(ap.planear(ap.leer(self._tabla("APROBADO", motivo="", evid="")))[0]["accion"], "ESPERAR")

    def test_error_se_reintenta_si_sigue_aprobado(self):
        self.assertEqual(ap.planear(ap.leer(self._tabla("APROBADO", "ERROR")))[0]["accion"], "CARGAR")


if __name__ == "__main__":
    unittest.main()
