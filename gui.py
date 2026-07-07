# -*- coding: utf-8 -*-
"""
URA — automatizacija ulaznih računa: GUI (CustomTkinter).
Dvoklik na 'pokreni_app.bat'. Prilagođeno krajnjem korisniku.
Sva logika je u main.py; ovdje je samo izgled + povezivanje gumba.
"""
import os
import queue
import logging
import threading
from datetime import datetime

import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image, ImageOps

import config
import main as app
from src.log_setup import postavi_logging

# Akcentne boje rade u oba načina (svijetli/tamni) pa ostaju obične.
ZELENA = "#2e8b40"
ZELENA_TAMNA = "#256e34"
NARANCASTA = "#c8881f"
PLAVA = "#2f6fb0"
CRVENA = "#b03a3a"
# Neutralne plohe/tekst kao (svijetla, tamna) — same se prebacuju s ctk.set_appearance_mode().
SIVA_TXT = ("#6b6b6b", "#a6a6a6")
KARTICA = ("#ffffff", "#2b2b2e")
KARTICA_HOVER = ("#eef0ea", "#3a3a40")
POZADINA = ("#f4f3ee", "#1e1e21")
GUMB = ("#ebeae4", "#343438")          # sekundarni gumbi / padajući izbornici
GUMB_HOVER = ("#dddcd5", "#42424a")
TXT = ("#1f1f1f", "#e6e6e6")           # tamni tekst -> svijetli u tamnom načinu
PLOHA = ("#fbfbf8", "#26262a")         # svijetle plohe (liste, polja)


class TkLogHandler(logging.Handler):
    def __init__(self, red):
        super().__init__()
        self.red = red

    def emit(self, record):
        try:
            self.red.put(self.format(record))
        except Exception:
            pass


def _dd(v):
    return v.strftime("%d.%m.%Y") if hasattr(v, "strftime") else (v if v is not None else "")


def _pf(s):
    """Parsiraj iznos iz teksta ('85,60' / '85.60' / '1.234,56' / '107,00 €') -> float ili None."""
    if s is None:
        return None
    s = str(s).replace("€", "").replace(" ", "").strip()
    if not s:
        return None
    if "," in s:                      # zarez = decimala, točka = tisućice
        s = s.replace(".", "").replace(",", ".")
    try:
        return round(float(s), 2)     # bez zareza: točka je decimala (ostavi)
    except ValueError:
        return None


def _vezi_klik(widget, cmd):
    """Učini widget (i svu djecu) klikabilnim."""
    widget.bind("<Button-1>", lambda e: cmd())
    for c in widget.winfo_children():
        _vezi_klik(c, cmd)


class App:
    def __init__(self, root):
        self.root = root
        self.red = queue.Queue()
        self.busy = False
        self.logger = postavi_logging(config.LOG_DIR)
        h = TkLogHandler(self.red)
        h.setFormatter(logging.Formatter("%(message)s"))
        self.logger.addHandler(h)

        root.title("URA-Automatizacija")
        try:
            import os as _os
            _ico = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "assets", "ura_icon.ico")
            if _os.path.exists(_ico):
                root.iconbitmap(_ico)            # ikona u naslovnoj traci + Alt-Tab
        except Exception:
            pass
        root.geometry("900x770")
        root.minsize(820, 690)
        root.configure(fg_color=POZADINA)

        # Windows/tkinter: znakovi č/ć/ž/š/đ tipkani preko AltGr (= Ctrl+Alt) se ne upišu jer
        # Tk ima ugrađeno pravilo "<Control-KeyPress> = ništa". Ovo ih ručno upiše u SVA polja.
        root.bind_class("Entry", "<Control-KeyPress>", self._altgr_upis)
        # Neke HR tipkovnice šalju č/ć/đ kao Latin-1 dvojnike (è/æ/ð) — ispravljamo pri upisu.
        root.bind_class("Entry", "<KeyPress>", self._key_fix, add="+")

        self._gradi()
        self._poll()
        self.osvjezi_plocice()

    # Neke HR tipkovnice šalju č/ć/đ kao Latin-1 dvojnike; preslikaj na prava slova.
    _HR_FIX = {"è": "č", "È": "Č", "æ": "ć", "Æ": "Ć", "ð": "đ", "Ð": "Đ"}

    @staticmethod
    def _key_fix(event):
        """Ako tipka pošalje krivi Latin-1 dvojnik (è/æ/ð...), zamijeni ga ispravnim
        hrvatskim slovom (č/ć/đ...). Radi bez obzira je li Tk već umetnuo krivi znak."""
        ch = App._HR_FIX.get(event.char)
        if not ch:
            return None
        w = event.widget
        try:
            pos = w.index("insert")
            if pos > 0 and w.get()[pos - 1] == event.char:
                w.delete(pos - 1, pos)
            w.insert("insert", ch)
        except Exception:
            return None
        return "break"

    @staticmethod
    def _altgr_upis(event):
        """AltGr (Ctrl+Alt) + znak -> upiši ga ručno (inače ga Tk 'proguta' preko ugrađenog
        no-op <Control-KeyPress>). Znak uzmi iz event.char; ako je prazan (AltGr zna dati
        prazan char) dohvatimo hrvatsko slovo iz keysym-a. Prave Ctrl-kratice (Ctrl+C/V/A…)
        daju neispisiv znak ili nemaju keysym u mapi pa ostaju netaknute."""
        hr = {"ccaron": "č", "Ccaron": "Č", "cacute": "ć", "Cacute": "Ć",
              "zcaron": "ž", "Zcaron": "Ž", "scaron": "š", "Scaron": "Š",
              "dstroke": "đ", "Dstroke": "Đ"}
        ch = event.char if (event.char and event.char.isprintable()) else hr.get(event.keysym, "")
        if not ch:
            return None
        alt = event.state & (0x0008 | 0x20000)        # Alt / AltGr bit (Windows/X11)
        if alt or ord(ch) > 127:
            event.widget.insert("insert", ch)
            return "break"
        return None

    # ---------- klikabilna kartica ----------
    def _kartica(self, parent, boja, hover, height, sadrzaj_fn, cmd):
        fr = ctk.CTkFrame(parent, fg_color=boja, corner_radius=14, height=height)
        fr.pack_propagate(False)
        sadrzaj_fn(fr)
        _vezi_klik(fr, cmd)

        def enter(_):
            if not self.busy:
                fr.configure(fg_color=hover)

        def leave(_):
            fr.configure(fg_color=boja)
        for w in [fr] + _svi_potomci(fr):
            w.bind("<Enter>", enter)
            w.bind("<Leave>", leave)
        return fr

    def _toggle_tema(self):
        """Prebaci svijetli/tamni način rada. Boje zadane kao (svijetla, tamna) par
        same se prilagode; gumbić mijenja ikonu (🌙 = uključi tamni, ☀️ = vrati svijetli)."""
        tamni = ctk.get_appearance_mode() == "Dark"
        ctk.set_appearance_mode("Light" if tamni else "Dark")
        self._tema_btn.configure(text="🌙" if tamni else "☀️")

    # ---------- izgradnja ----------
    def _gradi(self):
        glavni = ctk.CTkFrame(self.root, fg_color="transparent")
        glavni.pack(fill="both", expand=True, padx=18, pady=14)

        # zaglavlje
        zag = ctk.CTkFrame(glavni, fg_color="transparent")
        zag.pack(fill="x")
        lijevo = ctk.CTkFrame(zag, fg_color="transparent")
        lijevo.pack(side="left", anchor="w")
        ctk.CTkLabel(lijevo, text="Obrada ulaznih računa",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(lijevo, text=f"Knjiga: {os.path.basename(config.EXCEL_PATH)}",
                     text_color=SIVA_TXT, font=ctk.CTkFont(size=13)).pack(anchor="w")

        desno = ctk.CTkFrame(zag, fg_color="transparent")
        desno.pack(side="right", anchor="e")
        test = config.KORISTI_TEST_EXCEL
        ctk.CTkLabel(desno, text=("●  TEST kopija" if test else "●  PRAVI RAD"),
                     fg_color=("#e3f0e3" if test else "#f3e3d6"),
                     text_color=(ZELENA if test else NARANCASTA),
                     corner_radius=14, font=ctk.CTkFont(size=13, weight="bold")
                     ).pack(side="right", ipadx=12, ipady=6)
        ctk.CTkButton(desno, text="🧹 Obriši pamćenje", width=150, height=30,
                      fg_color=GUMB, hover_color=GUMB_HOVER, text_color=TXT,
                      corner_radius=10, command=self.obrisi_pamcenje).pack(side="right", padx=(0, 10))
        self._tema_btn = ctk.CTkButton(desno, text="🌙", width=36, height=30,
                      fg_color=GUMB, hover_color=GUMB_HOVER, text_color=TXT,
                      corner_radius=10, command=self._toggle_tema)
        self._tema_btn.pack(side="right", padx=(0, 10))

        # pločice
        ploce = ctk.CTkFrame(glavni, fg_color="transparent")
        ploce.pack(fill="x", pady=(14, 4))
        self.plo = {}
        # klik na "Fale podaci" -> ekran za upis vrste troška (s prikazom računa)
        klik = {"treba_potvrdu": self.fale_podaci}
        for kljuc, naslov, boja in (("racuni", "Računi", TXT),
                                    ("treba_potvrdu", "Fale podaci", PLAVA)):
            k = ctk.CTkFrame(ploce, fg_color=KARTICA, corner_radius=14)
            k.pack(side="left", expand=True, fill="x", padx=5)
            nl = ctk.CTkLabel(k, text=naslov, text_color=SIVA_TXT, font=ctk.CTkFont(size=13))
            nl.pack(anchor="w", padx=16, pady=(14, 0))
            lbl = ctk.CTkLabel(k, text="–", text_color=boja,
                               font=ctk.CTkFont(size=30, weight="bold"))
            lbl.pack(anchor="w", padx=16, pady=(0, 12))
            self.plo[kljuc] = lbl
            if kljuc in klik:                       # klikabilna pločica (pokaže popis)
                for w in (k, nl, lbl):
                    w.configure(cursor="hand2")
                    w.bind("<Button-1>", lambda e, f=klik[kljuc]: f())

        # velika tipka
        def big_sadrzaj(fr):
            li = ctk.CTkFrame(fr, fg_color="transparent")
            li.place(relx=0.03, rely=0.5, anchor="w")
            ctk.CTkLabel(li, text="Odradi sve", text_color="white",
                         font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
            ctk.CTkLabel(li, text="Računi s Parre + bankovni izvodi",
                         text_color="#dceadc", font=ctk.CTkFont(size=13)).pack(anchor="w")
            ctk.CTkLabel(fr, text="Pokreni  ▶", text_color="white",
                         font=ctk.CTkFont(size=16, weight="bold")).place(relx=0.97, rely=0.5, anchor="e")
        self._kartica(glavni, ZELENA, ZELENA_TAMNA, 86, big_sadrzaj, self.odradi).pack(fill="x", pady=(10, 6))

        # sekundarne kartice (2x2)
        mreza = ctk.CTkFrame(glavni, fg_color="transparent")
        mreza.pack(fill="x", pady=4)
        mreza.columnconfigure((0, 1), weight=1, uniform="x")
        kartice = [
            ("Pregledaj", "Što je novo (ništa ne dira)", self.pregledaj),
            ("Računi bez izvoda", "Neplaćeni računi", self.racuni_bez_izvoda),
            ("Izvodi bez računa", "Uplate koje čekaju račun", self.izvodi_bez_racuna),
            ("🔗 Spoji uplate", "Račun ↔ izvod (dobavljač + iznos)", self.spoji_uplate),
            ("📷 Fotkani računi", "Slikani papirnati računi", self.fotke),
            ("🧾 Generiraj naloge", "Svi nalozi iz 'tereni' (mjesec)", self.generiraj),
            ("➕ Novi putni nalog", "Ručno, jedan nalog", self.novi_putni),
            ("Otvori knjigu", os.path.basename(config.EXCEL_PATH), self.otvori_knjigu),
        ]
        for i, (nas, pod, cmd) in enumerate(kartice):
            def mk(nas, pod):
                def s(fr):
                    w = ctk.CTkFrame(fr, fg_color="transparent")
                    w.place(relx=0.06, rely=0.5, anchor="w")
                    ctk.CTkLabel(w, text=nas, text_color=TXT,
                                 font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w")
                    ctk.CTkLabel(w, text=pod, text_color=SIVA_TXT,
                                 font=ctk.CTkFont(size=12)).pack(anchor="w")
                return s
            kart = self._kartica(mreza, KARTICA, KARTICA_HOVER, 64, mk(nas, pod), cmd)
            kart.grid(row=i // 2, column=i % 2, sticky="ew", padx=5, pady=5)

        # putni nalozi
        pn = ctk.CTkFrame(glavni, fg_color=KARTICA, corner_radius=12)
        pn.pack(fill="x", pady=(6, 4))
        pl = ctk.CTkFrame(pn, fg_color="transparent")
        pl.pack(side="left", padx=16, pady=12)
        ctk.CTkLabel(pl, text="Putni nalozi", text_color=TXT,
                     font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(pl, text="Pročitaj mjesečni PN Excel i upiši naloge",
                     text_color=SIVA_TXT, font=ctk.CTkFont(size=12)).pack(anchor="w")
        d = datetime.now()
        self.mjesec = ctk.CTkOptionMenu(pn, width=72, values=[f"{m:02d}" for m in range(1, 13)],
                                        fg_color=GUMB, text_color=TXT, button_color=GUMB_HOVER,
                                        button_hover_color=GUMB_HOVER)
        self.mjesec.set(f"{d.month:02d}")
        self.godina = ctk.CTkOptionMenu(pn, width=86, values=[str(g) for g in range(2024, 2031)],
                                        fg_color=GUMB, text_color=TXT, button_color=GUMB_HOVER,
                                        button_hover_color=GUMB_HOVER)
        self.godina.set(str(d.year))
        ctk.CTkButton(pn, text="Obradi", width=110, height=36, corner_radius=10,
                      fg_color=ZELENA, hover_color=ZELENA_TAMNA, text_color="white",
                      command=self.putni).pack(side="right", padx=(8, 16))
        self.godina.pack(side="right", padx=4)
        self.mjesec.pack(side="right", padx=4)

        # aktivnosti
        zagA = ctk.CTkFrame(glavni, fg_color="transparent")
        zagA.pack(fill="x", pady=(10, 2))
        ctk.CTkLabel(zagA, text="Posljednje aktivnosti",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(side="left")
        self.status_lbl = ctk.CTkLabel(zagA, text="Spremno za rad", text_color=SIVA_TXT,
                                       font=ctk.CTkFont(size=13))
        self.status_lbl.pack(side="right")
        self.lista = ctk.CTkScrollableFrame(glavni, fg_color=KARTICA, corner_radius=12)
        self.lista.pack(fill="both", expand=True, pady=(2, 0))

        self.napredak = ctk.CTkProgressBar(glavni, mode="indeterminate", height=6)
        self.napredak.pack(fill="x", pady=(6, 0))
        self.napredak.set(0)

        self.aktivnost("Aplikacija spremna.", "OK", ZELENA)

    # ---------- pomoćno ----------
    def aktivnost(self, tekst, badge="OK", boja=ZELENA):
        red = ctk.CTkFrame(self.lista, fg_color=PLOHA, corner_radius=10)
        red.pack(fill="x", padx=8, pady=4)
        ctk.CTkLabel(red, text="●", text_color=boja,
                     font=ctk.CTkFont(size=14)).pack(side="left", padx=(12, 6))
        srednji = ctk.CTkFrame(red, fg_color="transparent")
        srednji.pack(side="left", fill="x", expand=True, pady=8)
        ctk.CTkLabel(srednji, text=tekst, text_color=TXT, anchor="w",
                     font=ctk.CTkFont(size=14), justify="left").pack(anchor="w")
        ctk.CTkLabel(srednji, text=datetime.now().strftime("%H:%M"), text_color=SIVA_TXT,
                     font=ctk.CTkFont(size=11)).pack(anchor="w")
        if badge:
            ctk.CTkLabel(red, text=badge, fg_color="#eaf1ea", text_color=boja, corner_radius=10,
                         font=ctk.CTkFont(size=12, weight="bold")).pack(side="right", padx=12, ipadx=10, ipady=3)

    def _poll(self):
        try:
            while True:
                linija = self.red.get_nowait()
                self.status_lbl.configure(text=linija[:70])
        except queue.Empty:
            pass
        self.root.after(150, self._poll)

    def _busy(self, on):
        self.busy = on
        if on:
            self.napredak.start()
        else:
            self.napredak.stop()
            self.napredak.set(0)

    def _zapocni(self, posao):
        if self.busy:
            return
        self._busy(True)

        def zadatak():
            try:
                posao()
            except Exception as e:
                self.root.after(0, lambda e=e: self.aktivnost(f"Greška: {e}", "Greška", CRVENA))
            finally:
                self.root.after(0, lambda: self._busy(False))
                self.root.after(0, self.osvjezi_plocice)
                self.root.after(0, lambda: self.status_lbl.configure(text="Spremno za rad"))

        threading.Thread(target=zadatak, daemon=True).start()

    def osvjezi_plocice(self):
        def radi():
            b = app.brojke_za_plocice(config.OBRADJUJ_OD)
            self.root.after(0, lambda: self._postavi_plocice(b))
        threading.Thread(target=radi, daemon=True).start()

    def _postavi_plocice(self, b):
        for k in self.plo:
            if k in b:
                self.plo[k].configure(text=str(b[k]))

    # ---------- akcije ----------
    def pregledaj(self):
        def posao():
            sesija = app.skeniraj(self.logger)
            self.root.after(0, lambda: self._otvori_pregled(sesija))
        self._zapocni(posao)

    def _otvori_pregled(self, sesija):
        if not sesija:
            self.aktivnost("Pregled nije uspio (vidi log).", "Greška", CRVENA)
            return
        if not sesija["kandidati"]:
            self.aktivnost("Nema novih računa za upis.", "OK", ZELENA)
            return
        self._pregled_prozor(sesija)

    def _pregled_prozor(self, sesija):
        kand = sesija["kandidati"]
        win = ctk.CTkToplevel(self.root)
        win.title("Pregled novih računa")
        win.geometry("840x640")
        win.configure(fg_color=POZADINA)
        ctk.CTkLabel(win, text=f"Označi koje račune upisati ({len(kand)})",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=16, pady=(12, 2))
        ctk.CTkLabel(win, text="✓ = upiši,  prazno = preskoči.   'Mogući duplikat' su prazni dok ih ti ne potvrdiš.",
                     text_color=SIVA_TXT, font=ctk.CTkFont(size=12)).pack(anchor="w", padx=16)
        lista = ctk.CTkScrollableFrame(win, fg_color=KARTICA)
        lista.pack(fill="both", expand=True, padx=16, pady=10)
        self._pregled_vars = []
        for k in kand:
            row = ctk.CTkFrame(lista, fg_color=PLOHA, corner_radius=8)
            row.pack(fill="x", padx=6, pady=3)
            dup = (k["tip"] != "novo")
            var = ctk.BooleanVar(value=not dup)
            ctk.CTkCheckBox(row, text="", width=24, variable=var,
                            command=self._azur_btn).pack(side="left", padx=(10, 4))
            txt = f"{k['datum']:<10}  {str(k['dob'])[:26]:26}  {k['iznos']:>9.2f} €   {k['broj']}"
            ctk.CTkLabel(row, text=txt, anchor="w",
                         font=ctk.CTkFont(family="Consolas", size=12)).pack(side="left", padx=4)
            if dup:
                ctk.CTkLabel(row, text="mogući duplikat", fg_color="#f3e7cf", text_color=NARANCASTA,
                             corner_radius=8, font=ctk.CTkFont(size=11, weight="bold")
                             ).pack(side="right", padx=10, ipadx=8, ipady=2)
            self._pregled_vars.append((k["idx"], var))

        donji = ctk.CTkFrame(win, fg_color="transparent")
        donji.pack(fill="x", padx=16, pady=(0, 14))
        self._btn_upisi = ctk.CTkButton(donji, text="Upiši označene", height=40, corner_radius=10,
                                        fg_color=ZELENA, hover_color=ZELENA_TAMNA,
                                        command=lambda: self._potvrdi_pregled(sesija, win))
        self._btn_upisi.pack(side="right")
        ctk.CTkButton(donji, text="Odustani", height=40, corner_radius=10, fg_color=GUMB,
                      hover_color=GUMB_HOVER, text_color=TXT, command=win.destroy).pack(side="right", padx=8)
        self._azur_btn()
        win.after(80, win.lift)

    def _azur_btn(self):
        n = sum(1 for _, v in self._pregled_vars if v.get())
        try:
            self._btn_upisi.configure(text=f"Upiši označene ({n})")
        except Exception:
            pass

    def _potvrdi_pregled(self, sesija, win):
        odabrani = [idx for idx, v in self._pregled_vars if v.get()]
        win.destroy()
        if not odabrani:
            self.aktivnost("Ništa nije označeno — ništa nije upisano.", "—", NARANCASTA)
            return

        def posao():
            stat = app.upisi_odabrane(sesija, odabrani, self.logger)
            self.root.after(0, lambda: self._sazetak(stat, "Upisano (pregled)"))
        self._zapocni(posao)

    def odradi(self):
        if not messagebox.askyesno("Odradi sve",
                                   "Upisat ću nove račune i obraditi izvode (uz backup). Nastaviti?"):
            return
        self._zapocni(lambda: self.root.after(
            0, lambda s=app.pokreni_obradu("odradi", self.logger): self._sazetak(s, "Odrađeno")))

    def putni(self):
        mj, god = int(self.mjesec.get()), int(self.godina.get())
        if not messagebox.askyesno("Putni nalozi", f"Obradi putne naloge za {mj:02d}-{god}? (uz backup)"):
            return

        def posao():
            n = app.pokreni_putne(god, mj, self.logger)
            self.root.after(0, lambda: self.aktivnost(
                f"Upisano {n} putnih naloga ({mj:02d}-{god})" if n
                else f"Nema novih putnih naloga za {mj:02d}-{god}",
                "OK" if n else "—", ZELENA if n else NARANCASTA))
        self._zapocni(posao)

    def racuni_bez_izvoda(self):
        def posao():
            r = app.izvjesce_racuni_bez_izvoda(None)   # cijela knjiga (i prije 1.5.)
            redovi = [f"UR{x['ur']:04d}   {str(x['racun'])[:24]:24}   "
                      f"{str(x['dobavljac'] or '')[:20]:20}   {_dd(x['datum']):10}   {x['ukupno']} €"
                      for x in r]
            self.root.after(0, lambda: self._popup(f"Računi kojima fali izvod ({len(r)})", redovi))
            self.root.after(0, lambda: self.aktivnost(f"Računi bez izvoda: {len(r)}", "Pregledaj", NARANCASTA))
        self._zapocni(posao)

    def izvodi_bez_racuna(self):
        def posao():
            r = app.izvjesce_izvodi_bez_racuna(None)   # cijela knjiga (i prije 1.5.)
            redovi = [f"izvod {str(x['izvadak']):>4}   {str(x['dobavljac'] or '')[:22]:22}   "
                      f"{_dd(x['datum']):10}   {x['placeno']} €   ({x['vrsta'] or '—'})"
                      for x in r]
            self.root.after(0, lambda: self._popup(f"Izvodi kojima fale računi ({len(r)})", redovi))
            self.root.after(0, lambda: self.aktivnost(f"Izvodi bez računa: {len(r)}", "Pregledaj", NARANCASTA))
        self._zapocni(posao)

    # ---------- Fale podaci (upis vrste troška uz prikaz računa) ----------
    VRSTE_TROSKA = [
        "gorivo", "cestarina", "leasing", "električna energija", "holding",
        "knjigovodstvo", "uredske potrepštine", "fiksna usluga", "pretplate",
        "osiguranje", "najam", "najam prostora", "čišćenje", "noćenje-pn",
        "preporučena pošiljka", "e paket", "napitci", "prijevoz", "dnevnice",
        "plaća", "porez", "zdravstveno osiguranje", "uplata po kreditu",
        "registracija", "tehnički pregled",
    ]

    def fale_podaci(self):
        def posao():
            sesija = app.fale_podaci_sesija(self.logger)
            self.root.after(0, lambda: self._fp_otvori(sesija))
        self._zapocni(posao)

    def _fp_otvori(self, sesija):
        if not sesija or not sesija["redovi"]:
            if sesija:
                app.zatvori_sesiju(sesija)
            self.aktivnost("Nema računa kojima fali vrsta troška. 👍", "—", ZELENA)
            return
        self._fp_sesija = sesija
        self._fp_idx = 0
        self._fp_rot = 0
        self._fp_zoom = 1.0
        self._fp_prozor()

    def _fp_prozor(self):
        win = ctk.CTkToplevel(self.root)
        self._fp_win = win
        win.title("Fale podaci — vrsta troška")
        win.geometry("940x800")
        win.configure(fg_color=POZADINA)
        win.protocol("WM_DELETE_WINDOW", self._fp_zatvori)
        self._fp_naslov = ctk.CTkLabel(win, text="", font=ctk.CTkFont(size=15, weight="bold"))
        self._fp_naslov.pack(anchor="w", padx=16, pady=(12, 6))

        donji = ctk.CTkFrame(win, fg_color="transparent")
        donji.pack(side="bottom", fill="x", padx=16, pady=(6, 14))
        ctk.CTkButton(donji, text="Upiši ✓  (pa sljedeći)", height=46, corner_radius=10,
                      fg_color=ZELENA, hover_color=ZELENA_TAMNA, font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._fp_upisi).pack(side="right")
        ctk.CTkButton(donji, text="Preskoči →", height=46, corner_radius=10, fg_color=GUMB,
                      hover_color=GUMB_HOVER, text_color=TXT, command=self._fp_preskoci).pack(side="right", padx=8)

        tijelo = ctk.CTkFrame(win, fg_color="transparent")
        tijelo.pack(fill="both", expand=True, padx=16, pady=(0, 6))

        # lijevo: prikaz računa (PDF) + zoom/rotacija/otvori
        lijevo = ctk.CTkFrame(tijelo, fg_color=KARTICA, corner_radius=12)
        lijevo.pack(side="left", fill="both", expand=True, padx=(0, 10))
        kontrole = ctk.CTkFrame(lijevo, fg_color="transparent")
        kontrole.pack(fill="x", pady=(8, 4))
        sg = dict(width=42, height=30, fg_color=GUMB, hover_color=GUMB_HOVER, text_color=TXT)
        ctk.CTkButton(kontrole, text="−", command=self._fp_zoom_out, **sg).pack(side="left", padx=(10, 2))
        ctk.CTkButton(kontrole, text="+", command=self._fp_zoom_in, **sg).pack(side="left", padx=2)
        ctk.CTkButton(kontrole, text="⟳", command=self._fp_rotiraj, **sg).pack(side="left", padx=2)
        ctk.CTkButton(kontrole, text="🔍 Otvori PDF", height=30, fg_color=GUMB,
                      hover_color=GUMB_HOVER, text_color=TXT, command=self._fp_otvori_pdf).pack(side="left", padx=8)
        skrol = ctk.CTkScrollableFrame(lijevo, fg_color="transparent")
        skrol.pack(fill="both", expand=True, padx=6, pady=(0, 8))
        self._fp_img = ctk.CTkLabel(skrol, text="")
        self._fp_img.pack(expand=True)

        # desno: podaci računa + upis vrste troška
        d = ctk.CTkFrame(tijelo, fg_color=KARTICA, corner_radius=12, width=390)
        d.pack(side="right", fill="y")
        d.pack_propagate(False)
        ctk.CTkLabel(d, text="Račun:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=16, pady=(14, 2))
        self._fp_info = ctk.CTkLabel(d, text="", justify="left", anchor="w", wraplength=350)
        self._fp_info.pack(anchor="w", padx=16, pady=(0, 10))
        ctk.CTkLabel(d, text="Vrsta troška:", text_color=SIVA_TXT).pack(anchor="w", padx=16, pady=(8, 2))
        self._fp_vrsta = ctk.CTkComboBox(d, values=self.VRSTE_TROSKA)
        self._fp_vrsta.pack(fill="x", padx=16)
        self._fp_vrsta.set("")
        ctk.CTkLabel(d, text="(odaberi iz popisa ili upiši svoju)", text_color=SIVA_TXT,
                     font=ctk.CTkFont(size=11)).pack(anchor="w", padx=16, pady=(4, 0))

        self._fp_prikazi()
        win.after(80, win.lift)

    def _fp_obnovi_sliku(self):
        r = self._fp_sesija["redovi"][self._fp_idx]
        p = r.get("pdf")
        if not p or not os.path.exists(p):
            self._fp_img.configure(image=None, text="(nema PDF-a za ovaj račun)")
            self._fp_img._image = None
            return
        img = self._f_slika(p, self._fp_rot, self._fp_zoom)   # isti učitavač kao za fotke
        self._fp_img.configure(image=img, text="" if img else "(nema pregleda)")
        self._fp_img._image = img

    def _fp_zoom_in(self):
        self._fp_zoom = min(5.0, self._fp_zoom * 1.3); self._fp_obnovi_sliku()

    def _fp_zoom_out(self):
        self._fp_zoom = max(0.3, self._fp_zoom / 1.3); self._fp_obnovi_sliku()

    def _fp_rotiraj(self):
        self._fp_rot = (self._fp_rot + 90) % 360; self._fp_obnovi_sliku()

    def _fp_otvori_pdf(self):
        try:
            os.startfile(self._fp_sesija["redovi"][self._fp_idx]["pdf"])
        except Exception as e:
            messagebox.showerror("Otvori PDF", str(e))

    def _fp_prikazi(self):
        s, i = self._fp_sesija, self._fp_idx
        if i >= len(s["redovi"]):
            self._fp_zatvori()
            return
        r = s["redovi"][i]
        self._fp_rot = 0; self._fp_zoom = 1.0
        self._fp_naslov.configure(text=f"Račun {i + 1} / {len(s['redovi'])}  —  UR{r['ur']:04d}")
        self._fp_info.configure(text=(f"UR: {r['ur']:04d}\nBroj: {r['racun']}\n"
                                      f"Dobavljač: {r['dobavljac'] or '—'}\n"
                                      f"Datum: {_dd(r['datum'])}\nIznos: {r['ukupno']} €"))
        self._fp_vrsta.set("")
        self._fp_obnovi_sliku()

    def _fp_preskoci(self):
        self._fp_idx += 1; self._fp_prikazi()

    def _fp_upisi(self):
        vrsta = self._fp_vrsta.get().strip()
        if not vrsta:
            messagebox.showwarning("Vrsta troška", "Odaberi ili upiši vrstu troška (ili klikni Preskoči).")
            return
        idx = self._fp_idx
        def posao():
            app.upisi_vrstu(self._fp_sesija, idx, vrsta, self.logger)
            ur = self._fp_sesija["redovi"][idx]["ur"]
            self.root.after(0, lambda: self.aktivnost(f"UR{ur:04d}: vrsta '{vrsta}'", "OK", ZELENA))
            self.root.after(0, self._fp_dalje)
        self._zapocni(posao)

    def _fp_dalje(self):
        self._fp_idx += 1; self._fp_prikazi()

    def _fp_zatvori(self):
        try:
            app.zatvori_sesiju(self._fp_sesija)
        except Exception:
            pass
        try:
            self._fp_win.destroy()
        except Exception:
            pass
        self.aktivnost("Gotovo — vrste troška upisane.", "OK", ZELENA)
        self.osvjezi_plocice()

    # ---------- Spoji uplate (račun ↔ izvod, ručna potvrda) ----------
    def spoji_uplate(self):
        def posao():
            sesija = app.kandidati_spajanja(self.logger)
            self.root.after(0, lambda: self._su_otvori(sesija))
        self._zapocni(posao)

    def _su_otvori(self, sesija):
        if not sesija or not sesija["parovi"]:
            if sesija:
                app.zatvori_sesiju(sesija)
            self.aktivnost("Nema mogućih spajanja (sve spojeno ili nema kandidata). 👍", "—", ZELENA)
            return
        self._su_sesija = sesija
        self._su_idx = 0
        self._su_prozor()

    def _su_prozor(self):
        win = ctk.CTkToplevel(self.root)
        self._su_win = win
        win.title("Spoji uplate s računima")
        win.geometry("860x520")
        win.configure(fg_color=POZADINA)
        win.protocol("WM_DELETE_WINDOW", self._su_zatvori)
        self._su_naslov = ctk.CTkLabel(win, text="", font=ctk.CTkFont(size=15, weight="bold"))
        self._su_naslov.pack(anchor="w", padx=16, pady=(12, 2))
        ctk.CTkLabel(win, text="Je li ova uplata za ovaj račun? Potvrdi ako jest.",
                     text_color=SIVA_TXT, font=ctk.CTkFont(size=12)).pack(anchor="w", padx=16)

        donji = ctk.CTkFrame(win, fg_color="transparent")
        donji.pack(side="bottom", fill="x", padx=16, pady=(6, 14))
        ctk.CTkButton(donji, text="✓ Da, spoji  (pa sljedeći)", height=46, corner_radius=10,
                      fg_color=ZELENA, hover_color=ZELENA_TAMNA, font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._su_spoji).pack(side="right")
        ctk.CTkButton(donji, text="Preskoči →", height=46, corner_radius=10, fg_color=GUMB,
                      hover_color=GUMB_HOVER, text_color=TXT, command=self._su_preskoci).pack(side="right", padx=8)

        tijelo = ctk.CTkFrame(win, fg_color="transparent")
        tijelo.pack(fill="both", expand=True, padx=16, pady=10)
        tijelo.columnconfigure((0, 1), weight=1, uniform="x")
        lev = ctk.CTkFrame(tijelo, fg_color=KARTICA, corner_radius=12)
        lev.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ctk.CTkLabel(lev, text="RAČUN (nije plaćen)", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=NARANCASTA).pack(anchor="w", padx=16, pady=(14, 6))
        self._su_inv = ctk.CTkLabel(lev, text="", justify="left", anchor="w", wraplength=360)
        self._su_inv.pack(anchor="w", padx=16, pady=(0, 12))
        des = ctk.CTkFrame(tijelo, fg_color=KARTICA, corner_radius=12)
        des.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ctk.CTkLabel(des, text="UPLATA s izvoda", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=ZELENA).pack(anchor="w", padx=16, pady=(14, 6))
        self._su_pay = ctk.CTkLabel(des, text="", justify="left", anchor="w", wraplength=360)
        self._su_pay.pack(anchor="w", padx=16, pady=(0, 12))

        self._su_prikazi()
        win.after(80, win.lift)

    def _su_prikazi(self):
        s, i = self._su_sesija, self._su_idx
        if i >= len(s["parovi"]):
            self._su_zatvori()
            return
        p = s["parovi"][i]
        self._su_naslov.configure(text=f"Mogući par {i + 1} / {len(s['parovi'])}")
        self._su_inv.configure(text=(f"UR: {p['ur']}\nBroj: {p['broj']}\n"
                                     f"Dobavljač: {p['dob_inv'] or '—'}\n"
                                     f"Datum: {_dd(p['dat_inv'])}\nIznos: {p['iznos']} €"))
        self._su_pay.configure(text=(f"Izvod: {p['izvadak']}\nDatum: {_dd(p['dat_pay'])}\n"
                                     f"Iznos: {p['placeno']} €\nSredstvo: {p['sred'] or '—'}\n"
                                     f"Naziv s izvoda: {p['dob_pay'] or '—'}"))

    def _su_preskoci(self):
        self._su_idx += 1
        self._su_prikazi()

    def _su_spoji(self):
        idx = self._su_idx
        def posao():
            rez = app.potvrdi_spajanje(self._su_sesija, idx, self.logger)
            p = self._su_sesija["parovi"][idx]
            if rez.get("ok"):
                self.root.after(0, lambda: self.aktivnost(
                    f"Spojeno: UR{p['ur']} ← izvod {p['izvadak']}", "OK", ZELENA))
            else:
                self.root.after(0, lambda: self.aktivnost(
                    f"Nije spojeno: {rez.get('greska', '')}", "—", NARANCASTA))
            self.root.after(0, self._su_dalje)
        self._zapocni(posao)

    def _su_dalje(self):
        self._su_idx += 1
        self._su_prikazi()

    def _su_zatvori(self):
        try:
            app.zatvori_sesiju(self._su_sesija)
        except Exception:
            pass
        try:
            self._su_win.destroy()
        except Exception:
            pass
        self.aktivnost("Gotovo sa spajanjem uplata.", "OK", ZELENA)
        self.osvjezi_plocice()

    # ---------- fotkani računi ----------
    def fotke(self):
        def posao():
            sesija = app.skeniraj_fotke(self.logger)
            self.root.after(0, lambda: self._fotke_otvori(sesija))
        self._zapocni(posao)

    def _fotke_otvori(self, sesija):
        if not sesija or not sesija["kandidati"]:
            self.aktivnost("Nema fotki za obradu (mapa prazna).", "—", NARANCASTA)
            return
        self._f_sesija = sesija
        self._f_idx = 0
        self._f_rot = 0
        self._fotke_prozor()
        self._f_prefetch_radi = True                  # pozadinsko predčitanje OCR-a (sve redom)
        threading.Thread(target=self._f_prefetch, daemon=True).start()

    def _fotke_prozor(self):
        win = ctk.CTkToplevel(self.root)
        self._f_win = win
        win.title("Fotkani računi")
        win.geometry("940x800")
        win.configure(fg_color=POZADINA)
        win.protocol("WM_DELETE_WINDOW", self._f_zatvori_fotke)
        self._f_naslov = ctk.CTkLabel(win, text="", font=ctk.CTkFont(size=15, weight="bold"))
        self._f_naslov.pack(anchor="w", padx=16, pady=(12, 6))

        # GUMBI — usidreni na DNO (uvijek vidljivi, bez obzira na visinu slike)
        donji = ctk.CTkFrame(win, fg_color="transparent")
        donji.pack(side="bottom", fill="x", padx=16, pady=(6, 14))
        ctk.CTkButton(donji, text="Upiši ✓  (pa sljedeći)", height=46, corner_radius=10,
                      fg_color=ZELENA, hover_color=ZELENA_TAMNA, font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._f_upisi).pack(side="right")
        ctk.CTkButton(donji, text="Preskoči →", height=46, corner_radius=10, fg_color=GUMB,
                      hover_color=GUMB_HOVER, text_color=TXT, command=self._f_preskoci).pack(side="right", padx=8)

        tijelo = ctk.CTkFrame(win, fg_color="transparent")
        tijelo.pack(fill="both", expand=True, padx=16, pady=(0, 6))

        # lijevo: slika + kontrole (zoom / rotacija / otvori)
        lijevo = ctk.CTkFrame(tijelo, fg_color=KARTICA, corner_radius=12)
        lijevo.pack(side="left", fill="both", expand=True, padx=(0, 10))
        kontrole = ctk.CTkFrame(lijevo, fg_color="transparent")
        kontrole.pack(fill="x", pady=(8, 4))
        sg = dict(width=42, height=30, fg_color=GUMB, hover_color=GUMB_HOVER, text_color=TXT)
        ctk.CTkButton(kontrole, text="−", command=self._f_zoom_out, **sg).pack(side="left", padx=(10, 2))
        ctk.CTkButton(kontrole, text="+", command=self._f_zoom_in, **sg).pack(side="left", padx=2)
        ctk.CTkButton(kontrole, text="⟳", command=self._f_rotiraj, **sg).pack(side="left", padx=2)
        ctk.CTkButton(kontrole, text="🔍 Otvori sliku", height=30, fg_color=GUMB,
                      hover_color=GUMB_HOVER, text_color=TXT,
                      command=self._f_otvori_sliku).pack(side="left", padx=8)
        # scroll područje (da se zoomirana slika može pomicati)
        skrol = ctk.CTkScrollableFrame(lijevo, fg_color="transparent")
        skrol.pack(fill="both", expand=True, padx=6, pady=(0, 8))
        self._f_img = ctk.CTkLabel(skrol, text="")
        self._f_img.pack(expand=True)

        # desno: polja
        d = ctk.CTkFrame(tijelo, fg_color=KARTICA, corner_radius=12, width=390)
        d.pack(side="right", fill="y")
        d.pack_propagate(False)
        ctk.CTkLabel(d, text="Provjeri i potvrdi:", font=ctk.CTkFont(size=14, weight="bold")
                     ).pack(anchor="w", padx=16, pady=(14, 2))

        def red(naziv):
            r = ctk.CTkFrame(d, fg_color="transparent")
            r.pack(fill="x", padx=16, pady=6)
            ctk.CTkLabel(r, text=naziv, width=104, anchor="w", text_color=SIVA_TXT).pack(side="left")
            return r

        r = red("Dobavljač"); self._f_dob = ctk.CTkEntry(r); self._f_dob.pack(side="right", fill="x", expand=True)
        r = red("Datum"); self._f_datum = ctk.CTkEntry(r); self._f_datum.pack(side="right", fill="x", expand=True)
        r = red("Bruto (za platiti)")
        self._f_bruto = ctk.CTkEntry(r); self._f_bruto.pack(side="right", fill="x", expand=True)
        r = red("Osnovica"); self._f_osn = ctk.CTkEntry(r); self._f_osn.pack(side="right", fill="x", expand=True)
        r = red("PDV"); self._f_pdv = ctk.CTkEntry(r); self._f_pdv.pack(side="right", fill="x", expand=True)
        r = red("Broj računa"); self._f_broj = ctk.CTkEntry(r); self._f_broj.pack(side="right", fill="x", expand=True)
        r = red("Auto")
        self._f_auto = ctk.CTkOptionMenu(r, values=[v[0] for v in config.VOZILA],
                                         command=lambda _=None: self._f_osobni_pdv(),
                                         fg_color=GUMB, text_color=TXT,
                                         button_color=GUMB_HOVER, button_hover_color=GUMB_HOVER)
        self._f_auto.pack(side="right", fill="x", expand=True)
        r = red("Vrsta troška"); self._f_vrsta = ctk.CTkEntry(r); self._f_vrsta.pack(side="right", fill="x", expand=True)
        for e in (self._f_bruto, self._f_osn, self._f_pdv):
            e.bind("<KeyRelease>", lambda _=None: self._f_provjera())
        # osobni auto: kad korisnik izađe iz PDV polja, prepolovi (50% pretporeza)
        self._f_pdv.bind("<FocusOut>", lambda _=None: self._f_osobni_pdv())
        self._f_provj = ctk.CTkLabel(d, text="", text_color=ZELENA, font=ctk.CTkFont(size=12),
                                     wraplength=350, justify="left")
        self._f_provj.pack(anchor="w", padx=16, pady=(10, 0))

        self._f_prikazi()
        win.after(80, win.lift)

    def _f_slika(self, path, rot, zoom=1.0):
        try:
            if path.lower().endswith(".pdf"):
                from src import racuni_ocr
                im = racuni_ocr._ucitaj_sliku(path)
            else:
                im = Image.open(path)
            im = ImageOps.exif_transpose(im)
            if rot:
                im = im.rotate(rot, expand=True)
            w, h = im.size
            f = min(380 / w, 560 / h, 1.0) * zoom
            return ctk.CTkImage(light_image=im, size=(max(1, int(w * f)), max(1, int(h * f))))
        except Exception:
            return None

    def _f_obnovi_sliku(self):
        k = self._f_sesija["kandidati"][self._f_idx]
        img = self._f_slika(k["slika"], self._f_rot, self._f_zoom)
        self._f_img.configure(image=img, text="" if img else "(nema pregleda)")
        self._f_img._image = img

    def _f_zoom_in(self):
        self._f_zoom = min(5.0, self._f_zoom * 1.3)
        self._f_obnovi_sliku()

    def _f_zoom_out(self):
        self._f_zoom = max(0.3, self._f_zoom / 1.3)
        self._f_obnovi_sliku()

    def _f_otvori_sliku(self):
        try:
            os.startfile(self._f_sesija["kandidati"][self._f_idx]["slika"])
        except Exception as e:
            messagebox.showerror("Otvori sliku", str(e))

    def _f_prikazi(self):
        s = self._f_sesija
        i = self._f_idx
        if i >= len(s["kandidati"]):
            self._f_zatvori_fotke()
            return
        k = s["kandidati"][i]
        self._f_rot = 0
        self._f_zoom = 1.0
        self._f_pdv_pola = None      # zadnji prepolovljeni PDV (osobni auto)
        self._f_pdv_puni = None      # puni PDV prije polovljenja (za vraćanje)
        import os as _os
        self._f_naslov.configure(text=f"Račun {i + 1} / {len(s['kandidati'])}  —  {_os.path.basename(k['slika'])}")
        self._f_obnovi_sliku()                       # slika odmah (ne treba OCR)
        if k["ocr"] is None:
            # još nije pročitana — pozadinska nit (predčitanje) prioritetno OCR-a baš ovu pa popuni
            self._f_ocisti_polja()
            self._f_provj.configure(text="Čitam račun (OCR)…", text_color=SIVA_TXT)
        else:
            self._f_popuni()

    def _f_prefetch(self):
        """Pozadinski OCR svih fotki redom (prioritet onoj koju gledaš) dok ti upisuješ."""
        s = self._f_sesija
        while getattr(self, "_f_prefetch_radi", False):
            cur = self._f_idx
            idx = None
            if 0 <= cur < len(s["kandidati"]) and s["kandidati"][cur]["ocr"] is None:
                idx = cur                                # prioritet: ono što trenutno gledaš
            else:
                for j in range(len(s["kandidati"])):
                    if s["kandidati"][j]["ocr"] is None:
                        idx = j
                        break
            if idx is None:
                break                                    # sve pročitano
            try:
                app.ocr_fotke_jedan(s, idx, self.logger)
            except Exception as e:
                self.logger.error("Predčitanje OCR greška (%s): %s", idx, e)
            if idx == self._f_idx and getattr(self, "_f_prefetch_radi", False):
                self.root.after(0, self._f_popuni)       # gledaš baš tu -> osvježi polja

    def _f_zatvori_fotke(self):
        self._f_prefetch_radi = False
        try:
            app.zatvori_sesiju(self._f_sesija)
        except Exception:
            pass
        try:
            self._f_win.destroy()
        except Exception:
            pass
        self.aktivnost("Gotovo s fotkanim računima.", "OK", ZELENA)
        self.osvjezi_plocice()

    def _f_ocisti_polja(self):
        for e in (self._f_dob, self._f_datum, self._f_bruto, self._f_osn, self._f_pdv, self._f_broj, self._f_vrsta):
            e.delete(0, "end")
        self._f_auto.set("(nije vezano uz auto)")

    def _f_popuni(self):
        k = self._f_sesija["kandidati"][self._f_idx]
        o, m = k["ocr"], k["match"]
        if not o:
            return
        def post(e, v):
            e.delete(0, "end")
            if v not in (None, ""):
                e.insert(0, str(v))
        post(self._f_dob, (m["dobavljac"] if m else None))
        post(self._f_datum, o.get("datum"))
        post(self._f_bruto, (f"{m['placeno']:.2f}".replace(".", ",") if m
                             else (f"{o['bruto']:.2f}".replace(".", ",") if o.get("bruto") else "")))
        post(self._f_osn, (f"{o['osnovica']:.2f}".replace(".", ",") if o.get("osnovica") else ""))
        post(self._f_pdv, (f"{o['pdv']:.2f}".replace(".", ",") if o.get("pdv") else ""))
        post(self._f_broj, o.get("broj"))
        self._f_auto.set("(nije vezano uz auto)")
        post(self._f_vrsta, "")
        self._f_provjera()

    def _f_osobni_pdv(self):
        """Osobni auto: priznaje se 50% pretporeza -> PDV polje prikaže POLA upisanog iznosa.
        Pamti puni iznos pa ga vrati ako se vozilo vrati na ne-osobno. Ne polovi dvaput."""
        osobni = config.vozilo_porez(self._f_auto.get()) == "osobni"
        trenutno = _pf(self._f_pdv.get())
        if osobni:
            # već prepolovljeno (polje sadrži baš naš izračun) -> ne diraj
            if (self._f_pdv_pola is not None and trenutno is not None
                    and abs(trenutno - self._f_pdv_pola) < 0.005):
                self._f_provjera(); return
            if trenutno is not None:
                self._f_pdv_puni = trenutno
                self._f_pdv_pola = round(trenutno / 2, 2)
                self._f_pdv.delete(0, "end")
                self._f_pdv.insert(0, f"{self._f_pdv_pola:.2f}".replace(".", ","))
        else:
            # vraćeno na ne-osobno: ako smo bili prepolovili, vrati puni iznos
            if (self._f_pdv_pola is not None and self._f_pdv_puni is not None
                    and trenutno is not None and abs(trenutno - self._f_pdv_pola) < 0.005):
                self._f_pdv.delete(0, "end")
                self._f_pdv.insert(0, f"{self._f_pdv_puni:.2f}".replace(".", ","))
            self._f_pdv_pola = None
            self._f_pdv_puni = None
        self._f_provjera()

    def _f_provjera(self):
        osn, pdv = _pf(self._f_osn.get()), _pf(self._f_pdv.get())
        bruto = _pf(self._f_bruto.get())
        osobni = config.vozilo_porez(self._f_auto.get()) == "osobni"
        if osobni:
            self._f_provj.configure(text="Osobni auto — PDV prepolovljen (50% pretporeza).", text_color=NARANCASTA)
        elif osn is not None and pdv is not None and bruto is not None:
            if abs((osn + pdv) - bruto) <= 0.02:
                self._f_provj.configure(text=f"✓ osnovica + PDV = {osn + pdv:.2f} = bruto", text_color=ZELENA)
            else:
                self._f_provj.configure(text=f"⚠ osn+PDV={osn + pdv:.2f} ≠ bruto {bruto:.2f} — provjeri!", text_color=CRVENA)
        else:
            self._f_provj.configure(text="", text_color=SIVA_TXT)

    def _f_rotiraj(self):
        self._f_rot = (self._f_rot + 90) % 360
        self._f_obnovi_sliku()

    def _f_preskoci(self):
        self._f_idx += 1
        self._f_prikazi()

    def _f_upisi(self):
        polja = {
            "broj": self._f_broj.get().strip() or None,
            "datum": app._parse_datum_hr(self._f_datum.get()),
            "dobavljac": self._f_dob.get().strip() or None,
            "osnovica": _pf(self._f_osn.get()),
            "pdv": _pf(self._f_pdv.get()),
            "vozilo": (self._f_auto.get() if config.vozilo_porez(self._f_auto.get()) else None),
            "vrsta": self._f_vrsta.get().strip() or None,
        }
        idx = self._f_idx

        def posao():
            try:
                rez = app.upisi_fotku(self._f_sesija, idx, polja, self.logger)
            except Exception as e:
                self.logger.error("Upiši fotku — greška: %s", e, exc_info=True)
                poruka = str(e)
                if "process" in poruka.lower() or "permission" in poruka.lower() or "WinError 32" in poruka:
                    poruka = ("Ne mogu spremiti knjigu — vjerojatno je OTVORENA u Excelu.\n"
                              "Zatvori URA knjigu u Excelu pa ponovi.")
                self.root.after(0, lambda p=poruka: messagebox.showerror("Ne mogu upisati", p))
                return
            if isinstance(rez, dict) and rez.get("duplikat"):
                pu = rez.get("ur")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Već upisan",
                    f"Račun {polja.get('broj')} već postoji u knjizi (UR {pu}).\n"
                    f"NIJE upisan (duplikat) — preskačem na sljedeći."))
                self.root.after(0, lambda: self.aktivnost(
                    f"Duplikat preskočen: {polja.get('broj')} (već UR {pu})", "—", NARANCASTA))
            else:
                self.root.after(0, lambda: self.aktivnost(
                    f"Fotka upisana: UR{rez:04d} ({polja.get('broj') or '—'})", "OK", ZELENA))
            self.root.after(0, self._f_dalje)
        self._zapocni(posao)

    def _f_dalje(self):
        self._f_idx += 1
        self._f_prikazi()

    # ---------- generiranje naloga iz 'tereni' ----------
    def generiraj(self):
        d = datetime.now()
        win = ctk.CTkToplevel(self.root)
        self._g_win = win
        win.title("Generiraj putne naloge"); win.geometry("520x300"); win.configure(fg_color=POZADINA)
        ctk.CTkLabel(win, text="Generiraj naloge za mjesec", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", padx=16, pady=(14, 2))
        ctk.CTkLabel(win, text="Iz tablice 'tereni' + ENC vremena. Excel sam računa dnevnice.",
                     text_color=SIVA_TXT, font=ctk.CTkFont(size=12)).pack(anchor="w", padx=16)
        karta = ctk.CTkFrame(win, fg_color=KARTICA, corner_radius=12); karta.pack(fill="x", padx=16, pady=12)

        r1 = ctk.CTkFrame(karta, fg_color="transparent"); r1.pack(fill="x", padx=14, pady=(12, 6))
        ctk.CTkLabel(r1, text="Mjesec / godina", width=120, anchor="w", text_color=TXT).pack(side="left")
        self._g_mj = ctk.CTkOptionMenu(r1, width=70, values=[f"{m:02d}" for m in range(1, 13)],
                                       fg_color=GUMB, text_color=TXT, button_color=GUMB_HOVER)
        self._g_mj.set(f"{d.month:02d}"); self._g_mj.pack(side="left", padx=4)
        self._g_god = ctk.CTkOptionMenu(r1, width=84, values=[str(g) for g in range(2024, 2031)],
                                        fg_color=GUMB, text_color=TXT, button_color=GUMB_HOVER)
        self._g_god.set(str(d.year)); self._g_god.pack(side="left", padx=4)

        r2 = ctk.CTkFrame(karta, fg_color="transparent"); r2.pack(fill="x", padx=14, pady=6)
        self._g_enc = None
        self._g_enc_lbl = ctk.CTkLabel(r2, text="ENC: (nije odabran — bit će standardna vremena)",
                                       text_color=SIVA_TXT, anchor="w")
        ctk.CTkButton(r2, text="Odaberi ENC CSV…", width=150, height=30, fg_color=GUMB,
                      hover_color=GUMB_HOVER, text_color=TXT, command=self._g_odaberi_enc).pack(side="left")
        self._g_enc_lbl.pack(side="left", padx=10)

        donji = ctk.CTkFrame(win, fg_color="transparent"); donji.pack(fill="x", padx=16, pady=(4, 14))
        ctk.CTkButton(donji, text="Generiraj ✓", height=42, corner_radius=10, fg_color=ZELENA,
                      hover_color=ZELENA_TAMNA, font=ctk.CTkFont(size=14, weight="bold"),
                      command=lambda: self._g_pokreni(win)).pack(side="right")
        ctk.CTkButton(donji, text="Odustani", height=42, corner_radius=10, fg_color=GUMB,
                      hover_color=GUMB_HOVER, text_color=TXT, command=win.destroy).pack(side="right", padx=8)
        win.after(80, win.lift)

    def _g_odaberi_enc(self):
        p = filedialog.askopenfilename(parent=self._g_win, title="Odaberi ENC CSV",
                                       filetypes=[("CSV", "*.csv"), ("Sve", "*.*")])
        if p:
            self._g_enc = p
            self._g_enc_lbl.configure(text=f"ENC: {os.path.basename(p)}", text_color=TXT)
        # vrati fokus na prozor (inače se sakri pa treba ponovo kliknuti)
        try:
            self._g_win.lift()
            self._g_win.focus_force()
            self._g_win.attributes("-topmost", True)
            self._g_win.after(200, lambda: self._g_win.attributes("-topmost", False))
        except Exception:
            pass

    def _g_pokreni(self, win):
        mj, god = int(self._g_mj.get()), int(self._g_god.get())
        enc = self._g_enc
        win.destroy()
        if not messagebox.askyesno("Generiraj naloge",
                                   f"Izradi sve naloge za {mj:02d}-{god} iz 'tereni'?\n(uz backup PN Excela)"):
            return
        def posao():
            try:
                n = app.generiraj_naloge(god, mj, enc, self.logger)
            except Exception as e:
                self.root.after(0, lambda e=e: (
                    self.aktivnost(f"Greška: {e}", "Greška", CRVENA),
                    messagebox.showerror("Greška", str(e))))
                return

            def gotovo():
                if n:
                    self.aktivnost(f"Generirano {n} naloga ({mj:02d}-{god})", "OK", ZELENA)
                    messagebox.showinfo("Gotovo ✓",
                                        f"Izrađeno {n} putnih naloga za {mj:02d}-{god}.\n\nOtvaram PN Excel…")
                else:
                    self.aktivnost(f"Nema novih naloga za {mj:02d}-{god}", "—", NARANCASTA)
                    messagebox.showwarning("Generiranje",
                                           f"Nema novih naloga za {mj:02d}-{god}\n(svi već postoje, ili pogledaj popis aktivnosti).")
            self.root.after(0, gotovo)
        self._zapocni(posao)

    # ---------- izrada putnog naloga ----------
    def novi_putni(self):
        d = datetime.now()
        def posao():
            pop = app.putni_popisi(d.year, d.month)
            self.root.after(0, lambda: self._putni_form(pop))
        self._zapocni(posao)

    def _putni_form(self, pop):
        if not pop:
            self.aktivnost("Ne nalazim PN Excel za popise (djelatnici/auti).", "Greška", CRVENA)
            return
        d = datetime.now()
        win = ctk.CTkToplevel(self.root)
        win.title("Novi putni nalog"); win.geometry("560x720"); win.configure(fg_color=POZADINA)
        ctk.CTkLabel(win, text="Novi putni nalog", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", padx=16, pady=(12, 2))
        ctk.CTkLabel(win, text="Dnevnice/za isplatu Excel sam izračuna.", text_color=SIVA_TXT,
                     font=ctk.CTkFont(size=12)).pack(anchor="w", padx=16)
        skrol = ctk.CTkScrollableFrame(win, fg_color=KARTICA)
        skrol.pack(fill="both", expand=True, padx=16, pady=10)
        self._pn_w = {}

        def red(naziv, w):
            r = ctk.CTkFrame(skrol, fg_color="transparent"); r.pack(fill="x", padx=10, pady=5)
            ctk.CTkLabel(r, text=naziv, width=128, anchor="w", text_color=TXT).pack(side="left")
            w.pack(side="right", fill="x", expand=True)
            return w

        def menu(values, default=None):
            m = ctk.CTkOptionMenu(skrol, values=values or ["—"], fg_color=GUMB, text_color=TXT,
                                  button_color=GUMB_HOVER, button_hover_color=GUMB_HOVER)
            if default:
                m.set(default)
            return m

        def entry(default=""):
            e = ctk.CTkEntry(skrol)
            if default:
                e.insert(0, default)
            return e

        self._pn_god = entry(str(d.year)); self._pn_mj = entry(f"{d.month:02d}")
        red("Mjesec PN (MM)", self._pn_mj); red("Godina", self._pn_god)
        self._pn_djel = menu(pop["djelatnici"]); red("Djelatnik", self._pn_djel)
        self._pn_mjesto = entry(); red("Mjesto odlaska", self._pn_mjesto)
        self._pn_svrha = menu(pop["svrhe"]); red("Svrha", self._pn_svrha)
        self._pn_d_od = entry(d.strftime("%d.%m.%Y")); red("Datum odlaska", self._pn_d_od)
        self._pn_v_od = entry("07:00"); red("Vrijeme odlaska", self._pn_v_od)
        self._pn_d_po = entry(d.strftime("%d.%m.%Y")); red("Datum povratka", self._pn_d_po)
        self._pn_v_po = entry("16:00"); red("Vrijeme povratka", self._pn_v_po)
        self._pn_vozilo = menu(pop["auti"]); red("Vozilo", self._pn_vozilo)
        self._pn_vrsta = menu(pop["vrste"], "službeno"); red("Vrsta prijevoza", self._pn_vrsta)
        self._pn_poc = entry(); red("Poč. brojilo", self._pn_poc)
        self._pn_zav = entry(); red("Zav. brojilo", self._pn_zav)
        self._pn_drz = menu([x[0] for x in pop["drzave"]], "HR"); red("Država (dnevnica)", self._pn_drz)
        self._pn_iznos = entry("30"); red("Dnevnica €", self._pn_iznos)
        self._pn_locco = entry("0,4"); red("Locco naknada", self._pn_locco)

        donji = ctk.CTkFrame(win, fg_color="transparent"); donji.pack(fill="x", padx=16, pady=(0, 14))
        ctk.CTkButton(donji, text="Izradi nalog ✓", height=42, corner_radius=10, fg_color=ZELENA,
                      hover_color=ZELENA_TAMNA, font=ctk.CTkFont(size=14, weight="bold"),
                      command=lambda: self._putni_izradi(win)).pack(side="right")
        ctk.CTkButton(donji, text="Odustani", height=42, corner_radius=10, fg_color=GUMB,
                      hover_color=GUMB_HOVER, text_color=TXT, command=win.destroy).pack(side="right", padx=8)
        win.after(80, win.lift)

    def _putni_izradi(self, win):
        try:
            god, mj = int(self._pn_god.get()), int(self._pn_mj.get())
        except ValueError:
            messagebox.showerror("Nalog", "Mjesec/godina nisu ispravni."); return
        podaci = {
            "djelatnik": self._pn_djel.get(), "mjesto": self._pn_mjesto.get().strip(),
            "svrha": self._pn_svrha.get(),
            "datum_odlaska": app._parse_datum_hr(self._pn_d_od.get()),
            "vrijeme_odlaska": self._pn_v_od.get().strip(),
            "datum_povratka": app._parse_datum_hr(self._pn_d_po.get()),
            "vrijeme_povratka": self._pn_v_po.get().strip(),
            "vozilo": self._pn_vozilo.get(), "vrsta_prijevoz": self._pn_vrsta.get(),
            "vrsta_prijevoza": self._pn_vrsta.get(),
            "poc_brojilo": _pf(self._pn_poc.get()), "zav_brojilo": _pf(self._pn_zav.get()),
            "drzava": self._pn_drz.get(), "dnevnica_iznos": _pf(self._pn_iznos.get()),
            "locco": _pf(self._pn_locco.get()),
        }
        win.destroy()
        def posao():
            broj = app.kreiraj_putni(god, mj, podaci, self.logger)
            self.root.after(0, lambda: self.aktivnost(
                f"Putni nalog {broj} izrađen ({mj:02d}-{god})" if broj else
                "Izrada naloga nije uspjela (vidi log)", "OK" if broj else "Greška",
                ZELENA if broj else CRVENA))
        self._zapocni(posao)

    def otvori_knjigu(self):
        try:
            os.startfile(config.EXCEL_PATH)
        except Exception as e:
            messagebox.showerror("Otvori knjigu", str(e))

    def obrisi_pamcenje(self):
        if not messagebox.askyesno("Obriši pamćenje",
                                   "Obrisat ću zapamćeno stanje (što je već obrađeno) — za ponovni čisti "
                                   "test.\nU pravom radu ovo NE treba dirati.\nNastaviti?"):
            return
        app.obrisi_pamcenje(self.logger)
        self.aktivnost("Pamćenje očišćeno — sljedeća obrada kreće ispočetka.", "OK", ZELENA)
        self.osvjezi_plocice()

    # ---------- prozori ----------
    def _popup(self, naslov, redovi):
        win = ctk.CTkToplevel(self.root)
        win.title(naslov)
        win.geometry("740x540")
        win.configure(fg_color=POZADINA)
        ctk.CTkLabel(win, text=naslov, font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", padx=16, pady=10)
        box = ctk.CTkTextbox(win, font=ctk.CTkFont(family="Consolas", size=12))
        box.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        box.insert("end", "\n".join(redovi) if redovi else "(nema stavki)")
        box.configure(state="disabled")
        win.after(60, win.lift)

    def _sazetak(self, stat, naslov):
        if not stat:
            return
        if stat.get("zadnji_racun"):
            self.aktivnost(f"Zadnji račun s Parre: {stat['zadnji_racun']}", "info", PLAVA)
        self.aktivnost(
            f"{naslov}: upisano {stat.get('upisano', 0)} računa, spojeno "
            f"{stat.get('spojeno', 0)} stavki izvoda, novih UR0 {stat.get('ur0_redova', 0)}",
            "OK", ZELENA)
        if stat.get("deferno"):
            self.aktivnost(f"Deferno dopunjeno (uplata je čekala račun): {stat['deferno']}", "OK", ZELENA)
        if stat.get("neizvjesno"):
            self.aktivnost(f"{stat['neizvjesno']} stavki treba tvoju potvrdu (mogući duplikat)",
                           "Pregledaj", NARANCASTA)
        if stat.get("pdf_nije_nadjen"):
            self.aktivnost(f"PDF nije nađen za {stat['pdf_nije_nadjen']} računa", "Provjeri", NARANCASTA)
        if stat.get("greske"):
            self.aktivnost(f"Greške: {stat['greske']} (pogledaj log)", "Greška", CRVENA)


def _svi_potomci(w):
    out = []
    for c in w.winfo_children():
        out.append(c)
        out += _svi_potomci(c)
    return out


def pokreni_app():
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("green")
    root = ctk.CTk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    pokreni_app()
