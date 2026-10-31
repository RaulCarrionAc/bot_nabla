import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from api_wsp.identidad import identificadores_remitente


class IdentidadTests(unittest.TestCase):
    def test_lid_autorizado_se_conserva_al_resolver_telefono(self):
        ids = identificadores_remitente({"from": "12345678901234@lid"}, "56911111111")
        self.assertEqual(ids, {"12345678901234", "56911111111"})

    def test_grupo_no_autoriza_por_chat_o_destinatario(self):
        ids = identificadores_remitente({"author": "222@lid", "from": "111@g.us",
                                        "chatId": "111@g.us", "to": "333@c.us"}, "56922222222")
        self.assertEqual(ids, {"222", "56922222222"})

    def test_desconocido_no_adquiere_identidad_admin(self):
        ids = identificadores_remitente({"from": "222@lid", "chatId": "111@lid"}, "56922222222")
        self.assertFalse(ids.intersection({"111", "56911111111"}))

    def test_sender_nulo(self):
        self.assertEqual(identificadores_remitente({"sender": None, "from": "222@lid"}, ""), {"222"})
