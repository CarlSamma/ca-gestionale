# -*- coding: utf-8 -*-
"""Test della lettura delle mail di Como Acqua (dati di fantasia).
    py -m unittest discover -s test        (dalla cartella della skill)"""
import sys
import unittest
from pathlib import Path

QUI = Path(__file__).resolve().parent
sys.path.insert(0, str(QUI.parent / "scripts"))
from mail_ca import is_conferma_lettura, parse_mail  # noqa: E402


def leggi(nome):
    return (QUI / "mail" / nome).read_text(encoding="utf-8")


class ConfermeLettura(unittest.TestCase):
    def test_oggetti(self):
        si = ["Letto: Codice Cliente: C1 - X", "I: Letto: Codice Cliente: C1 - X", "R: I: Letto: pratica",
              "Read: New case", "Fwd: Read: x", "Notifica di lettura: pratica", "Non letto: pratica",
              "Ricevuta di lettura - pratica", "Conferma di lettura: x"]
        no = ["I: Codice Cliente: C1 - X – Nuova pratica", "Lettera di messa in mora", "Letture contatori",
              "I: NUOVA PRATICA - DEBITORE 1 - LETTO SRL"]
        for o in si:
            self.assertTrue(is_conferma_lettura(o), o)
        for o in no:
            self.assertFalse(is_conferma_lettura(o), o)

    def test_corpo(self):
        # oggetto anonimo ma corpo che riporta «Oggetto: Letto: ...»
        self.assertTrue(is_conferma_lettura("I: pratica", leggi("conferma_lettura.txt")))
        self.assertFalse(is_conferma_lettura("I: pratica", leggi("multi_liquidazione.txt")))
        d = parse_mail("I: Letto: x", leggi("conferma_lettura.txt"))
        self.assertTrue(d["conferma_lettura"])


class LetturaCompleta(unittest.TestCase):
    def controlla_multi(self, d, sdi=True):
        self.assertEqual(d["debitore"], "ALFA LAVAGGI S.R.L. IN LIQUIDAZIONE")
        self.assertEqual(d["forma"], "S.r.l.")
        self.assertTrue(d["in_liquidazione"])
        self.assertEqual(d["indirizzo"], {"via": "VIA DEI MILLE, 14", "cap": "20017", "citta": "RHO"})
        self.assertEqual(d["pec"], "alfalavaggi@pec.example.it")
        self.assertEqual(d["piva"], "01234567890")
        self.assertTrue(d["stato_piva"].startswith("CESSATA in data 21/12/2023"))
        self.assertEqual([p["codice_cliente"] for p in d["posizioni"]], ["C999", "12345"])
        self.assertEqual([p["importo"] for p in d["posizioni"]], ["4.345,23", "9.127,90"])
        self.assertEqual(d["debito"], "13.473,13")          # NOTA: il totale va sempre verificato
        self.assertEqual(d["causa"], "fatture")
        v = " ".join(d["verifiche"])
        for parola in ("LIQUIDAZIONE", "P.IVA cessata", "CESSATO", "fornitura") + (("SDI",) if sdi else ()):
            self.assertIn(parola, v)
        self.assertNotIn("IMPORTO TOTALE DA VERIFICARE", v)

    def test_gmail_due_posizioni(self):
        d = parse_mail("I: Codici Cliente: C999 / 12345 - ALFA", leggi("multi_liquidazione.txt"))
        self.controlla_multi(d)
        self.assertEqual(d["data_mail"], "25/11/2026")
        self.assertIn("Nuova pratica", d["oggetto_mail"])

    def test_elenco_puntato(self):
        self.controlla_multi(parse_mail("", leggi("elenco_puntato.md")), sdi=False)

    def test_totale_non_coincide(self):
        t = leggi("multi_liquidazione.txt").replace("13.473,13", "13.000,00")
        d = parse_mail("", t)
        self.assertEqual(d["debito"], "13.000,00")
        self.assertTrue(any("IMPORTO TOTALE DA VERIFICARE" in v for v in d["verifiche"]))

    def test_singola_con_pdr(self):
        d = parse_mail("", leggi("singola_pdr.txt"))
        self.assertEqual(d["debitore"], "BETA MECCANICA S.R.L.S.")
        self.assertEqual(d["forma"], "S.r.l.s.")
        self.assertEqual(d["codice_cliente"], "C111")
        self.assertEqual(d["indirizzo"]["citta"], "VILLA GUARDIA")
        self.assertEqual(d["debito"], "24.769,85")
        self.assertEqual(d["pdr"], {"numero": "PDR24-001600", "data": "31/12/2024"})
        self.assertEqual(d["causa"], "pdr")
        self.assertNotIn("frase_causa", d)

    def test_mail_libera(self):
        d = parse_mail("I: NUOVA PRATICA - DEBITORE 130000 - GAMMA FITNESS S.R.L.", leggi("libera.txt"))
        self.assertEqual(d["debitore"], "GAMMA FITNESS S.R.L.")
        self.assertEqual(d["codice_cliente"], "130000")
        self.assertEqual(d["debito"], "3.997,34")
        self.assertEqual(d["data_mail"], "03/11/2026")
        self.assertIn("PEC del debitore assente nella mail", d["verifiche"])


if __name__ == "__main__":
    unittest.main()
