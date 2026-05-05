"""
GUI (Tkinter) dla scrapera Gratka.pl
Uruchom:  python gui.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import os
import subprocess
import sys
from datetime import datetime

from scraper_nieruchomosci import pobierz_oferty
from excel_generator import generuj_excel


# ─────────────────────────────────────────────
#  STAŁE
# ─────────────────────────────────────────────

TYPY_NIERUCHOMOSCI = ["Mieszkanie", "Dom", "Kawalerka", "Działka", "Lokal"]
TYPY_TRANSAKCJI    = ["Sprzedaż", "Wynajem"]

KOLOR_TLO      = "#F8FAFF"
KOLOR_AKCENT   = "#2563EB"
KOLOR_AKCENT_H = "#1D4ED8"
KOLOR_PANEL    = "#FFFFFF"
KOLOR_BORDER   = "#E5E7EB"
KOLOR_TEKST    = "#1B2A4A"
KOLOR_MUTED    = "#6B7280"
KOLOR_SUKCES   = "#059669"
KOLOR_BLAD     = "#DC2626"
KOLOR_LOG_TLO  = "#0F172A"
KOLOR_LOG_TEKST= "#94A3B8"
CZCIONKA       = "Segoe UI"


# ─────────────────────────────────────────────
#  GŁÓWNE OKNO
# ─────────────────────────────────────────────

class ScraperGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Scraper Gratka.pl → Excel")
        self.root.geometry("680x720")
        self.root.minsize(600, 620)
        self.root.configure(bg=KOLOR_TLO)
        self.root.resizable(True, True)

        self._ostatni_plik: str = ""
        self._w_trakcie = False

        self._zbuduj_ui()

    # ── BUDOWANIE UI ──────────────────────────────────────────────

    def _zbuduj_ui(self):
        # Nagłówek
        naglowek = tk.Frame(self.root, bg=KOLOR_AKCENT, pady=0)
        naglowek.pack(fill="x")

        tk.Label(
            naglowek,
            text="Scraper Gratka.pl  →  Excel",
            font=(CZCIONKA, 16, "bold"),
            bg=KOLOR_AKCENT, fg="white",
            pady=14, padx=20, anchor="w"
        ).pack(fill="x")

        tk.Label(
            naglowek,
            text="Ustaw filtry i kliknij Start. Wyniki zostaną zapisane do pliku .xlsx",
            font=(CZCIONKA, 9),
            bg=KOLOR_AKCENT, fg="#BFDBFE",
            pady=0, padx=20, anchor="w"
        ).pack(fill="x")

        tk.Frame(naglowek, bg=KOLOR_AKCENT, height=10).pack()

        # Kontener główny
        main = tk.Frame(self.root, bg=KOLOR_TLO, padx=20, pady=16)
        main.pack(fill="both", expand=True)

        # ── Panel filtrów ─────────────────────────────────────────
        panel = tk.Frame(main, bg=KOLOR_PANEL, bd=0, relief="flat",
                         highlightbackground=KOLOR_BORDER, highlightthickness=1)
        panel.pack(fill="x", pady=(0, 12))

        tk.Label(panel, text="Filtry wyszukiwania",
                 font=(CZCIONKA, 10, "bold"), bg=KOLOR_PANEL,
                 fg=KOLOR_TEKST, padx=16, pady=10, anchor="w"
                 ).pack(fill="x")

        tk.Frame(panel, bg=KOLOR_BORDER, height=1).pack(fill="x")

        grid = tk.Frame(panel, bg=KOLOR_PANEL, padx=16, pady=14)
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(3, weight=1)

        def label(row, col, tekst):
            tk.Label(
                grid, text=tekst, font=(CZCIONKA, 9),
                fg=KOLOR_MUTED, bg=KOLOR_PANEL, anchor="w"
            ).grid(row=row, column=col, sticky="w", pady=(0, 2))

        def wpis(row, col, var, placeholder="", width=22):
            e = tk.Entry(
                grid, textvariable=var,
                font=(CZCIONKA, 10), relief="flat",
                bg="#F3F4F6", fg=KOLOR_TEKST,
                insertbackground=KOLOR_AKCENT,
                width=width, bd=0,
                highlightbackground=KOLOR_BORDER,
                highlightthickness=1,
            )
            e.grid(row=row, column=col, sticky="ew", ipady=6, pady=(0, 10), padx=(0, 16))
            return e

        def combo(row, col, var, wartosci, width=20):
            c = ttk.Combobox(
                grid, textvariable=var, values=wartosci,
                state="readonly", font=(CZCIONKA, 10), width=width
            )
            c.grid(row=row, column=col, sticky="ew", ipady=4, pady=(0, 10), padx=(0, 16))
            return c

        # Styl Combobox
        styl = ttk.Style()
        styl.theme_use("clam")
        styl.configure("TCombobox",
                        fieldbackground="#F3F4F6",
                        background="#F3F4F6",
                        bordercolor=KOLOR_BORDER,
                        arrowcolor=KOLOR_MUTED,
                        relief="flat")

        # Zmienne
        self.v_miasto    = tk.StringVar(value="")
        self.v_typ       = tk.StringVar(value="Mieszkanie")
        self.v_transakcja= tk.StringVar(value="Sprzedaż")
        self.v_liczba    = tk.StringVar(value="20")
        self.v_plik      = tk.StringVar(value="raport_gratka.xlsx")

        # Wiersz 1: Miejscowość | Typ nieruchomości
        label(0, 0, "Miejscowość (puste = cała Polska)")
        label(0, 2, "Rodzaj nieruchomości")
        wpis(1, 0, self.v_miasto, "np. Warszawa, Jeżyce")
        combo(1, 2, self.v_typ, TYPY_NIERUCHOMOSCI)

        # Wiersz 2: Transakcja | Liczba ofert
        label(2, 0, "Sprzedaż / wynajem")
        label(2, 2, "Liczba ofert do pobrania")
        combo(3, 0, self.v_transakcja, TYPY_TRANSAKCJI)
        wpis(3, 2, self.v_liczba, "ex. 30", width=10)

        # Wiersz 3: Nazwa pliku
        label(4, 0, "Nazwa pliku wyjściowego (.xlsx)")
        wpis(5, 0, self.v_plik, "raport_gratka.xlsx", width=48)

        # ── Przyciski ─────────────────────────────────────────────
        btn_row = tk.Frame(main, bg=KOLOR_TLO)
        btn_row.pack(fill="x", pady=(0, 12))

        self.btn_start = tk.Button(
            btn_row,
            text="▶  Start",
            font=(CZCIONKA, 10, "bold"),
            bg=KOLOR_AKCENT, fg="white",
            activebackground=KOLOR_AKCENT_H, activeforeground="white",
            relief="flat", bd=0, padx=20, pady=9, cursor="hand2",
            command=self._start
        )
        self.btn_start.pack(side="left")

        self.btn_otworz = tk.Button(
            btn_row,
            text="📂  Otwórz plik wyjściowy Excel",
            font=(CZCIONKA, 10),
            bg=KOLOR_PANEL, fg=KOLOR_TEKST,
            activebackground="#E5E7EB",
            relief="flat", bd=0, padx=16, pady=9, cursor="hand2",
            highlightbackground=KOLOR_BORDER, highlightthickness=1,
            command=self._otworz_plik,
            state="disabled"
        )
        self.btn_otworz.pack(side="left", padx=(10, 0))

        self.btn_wyczysc = tk.Button(
            btn_row,
            text="🗑  Wyczyść log",
            font=(CZCIONKA, 10),
            bg=KOLOR_PANEL, fg=KOLOR_MUTED,
            activebackground="#E5E7EB",
            relief="flat", bd=0, padx=16, pady=9, cursor="hand2",
            highlightbackground=KOLOR_BORDER, highlightthickness=1,
            command=self._wyczysc_log
        )
        self.btn_wyczysc.pack(side="right")

        # ── Pasek postępu ─────────────────────────────────────────
        self.pasek = ttk.Progressbar(main, mode="indeterminate", length=200)
        styl.configure("TProgressbar", troughcolor=KOLOR_BORDER,
                        background=KOLOR_AKCENT, thickness=4)
        self.pasek.pack(fill="x", pady=(0, 10))

        # ── Log ───────────────────────────────────────────────────
        log_frame = tk.Frame(main, bg=KOLOR_PANEL,
                             highlightbackground=KOLOR_BORDER, highlightthickness=1)
        log_frame.pack(fill="both", expand=True)

        tk.Label(log_frame, text="Log działania",
                 font=(CZCIONKA, 9, "bold"), bg=KOLOR_PANEL,
                 fg=KOLOR_MUTED, padx=12, pady=6, anchor="w"
                 ).pack(fill="x")
        tk.Frame(log_frame, bg=KOLOR_BORDER, height=1).pack(fill="x")

        self.log_box = scrolledtext.ScrolledText(
            log_frame,
            font=("Consolas", 9),
            bg=KOLOR_LOG_TLO, fg=KOLOR_LOG_TEKST,
            insertbackground=KOLOR_LOG_TEKST,
            relief="flat", bd=0,
            wrap="word", state="disabled",
            padx=12, pady=10,
        )
        self.log_box.pack(fill="both", expand=True)

        # Tagi kolorów w logu
        self.log_box.tag_config("info",    foreground="#94A3B8")
        self.log_box.tag_config("sukces",  foreground="#34D399")
        self.log_box.tag_config("blad",    foreground="#F87171")
        self.log_box.tag_config("ostrzez", foreground="#FBBF24")
        self.log_box.tag_config("akcent",  foreground="#60A5FA")

        self._log("Gotowy do uruchomienia. Ustaw filtry i wciśnij Start.", "info")

    # ── LOGOWANIE ─────────────────────────────────────────────────

    def _log(self, msg: str, tag: str = "info"):
        def _wstaw():
            self.log_box.configure(state="normal")
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_box.insert("end", f"[{ts}] {msg}\n", tag)
            self.log_box.configure(state="disabled")
            self.log_box.see("end")
        self.root.after(0, _wstaw)

    def _wyczysc_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ── WALIDACJA ─────────────────────────────────────────────────

    def _waliduj(self) -> bool:
        try:
            n = int(self.v_liczba.get())
            if n < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Błąd", "Liczba ofert musi być liczbą całkowitą powyżej zera.")
            return False

        plik = self.v_plik.get().strip()
        if not plik:
            messagebox.showerror("Błąd", "Ustaw nazwę pliku wyjściowego.")
            return False
        if not plik.endswith(".xlsx"):
            self.v_plik.set(plik + ".xlsx")

        return True

    # ── START ─────────────────────────────────────────────────────

    def _start(self):
        if self._w_trakcie:
            return
        if not self._waliduj():
            return

        self._w_trakcie = True
        self.btn_start.configure(state="disabled", text="⏳ Przetwarzanie...")
        self.btn_otworz.configure(state="disabled")
        self.pasek.start(12)

        miasto     = self.v_miasto.get().strip()
        typ        = self.v_typ.get()
        transakcja = self.v_transakcja.get()
        liczba     = int(self.v_liczba.get())
        plik       = self.v_plik.get().strip()

        self._log(f"Start: {typ} | {transakcja} | {miasto or 'cała Polska'} | {liczba} ofert", "akcent")

        threading.Thread(
            target=self._scrapuj,
            args=(miasto, typ, transakcja, liczba, plik),
            daemon=True
        ).start()

    def _scrapuj(self, miasto, typ, transakcja, liczba, plik):
        """Działa w osobnym wątku żeby nie blokować UI."""
        try:
            wyniki, osiagnieto_limit = pobierz_oferty(
                liczba=liczba,
                max_stron=20,
                miasto=miasto,
                typ_nieruchomosci=typ,
                transakcja=transakcja,
                callback_postep=lambda msg: self._log(msg, "info"),
            )

            if not wyniki:
                self._log("Nie znaleziono ofert. Sprawdź filtry.", "blad")
                self._zakoncz(sukces=False)
                return

            # Komunikat gdy Gratka ma mniej ofert niż żądano
            if not osiagnieto_limit:
                ile = len(wyniki)
                self._log(
                    f"⚠ Serwis Gratka.pl miał mniej ofert niż oczekiwana liczba ({liczba}). "
                    f"Pobrano wszystkie dostępne oferty: {ile}.",
                    "ostrzez"
                )
                self.root.after(0, lambda: messagebox.showwarning(
                    "Mniej ofert dostępnych niż oczekiwano"
                    f"Serwis gratka.pl ma mniej ofert niż oczekiwana liczba ({liczba}).\n\n"
                    f"Pobrano wszystkie dostępne oferty: {ile}.\n"
                    f"Raport zostanie wygenerowany w oparciu o wszystkie dostępne oferty."
                ))

            self._log(f"Generowanie pliku Excel: {plik}", "akcent")
            sciezka = generuj_excel(wyniki, plik)
            self._ostatni_plik = os.path.abspath(sciezka)

            self._log(f"✓ Zapisano: {self._ostatni_plik}", "sukces")
            self._log(f"✓ Raport składa się z {len(wyniki)} ofert", "sukces")
            self._zakoncz(sukces=True)

        except Exception as e:
            self._log(f"Błąd: {e}", "blad")
            self._zakoncz(sukces=False)

    def _zakoncz(self, sukces: bool):
        def _ui():
            self.pasek.stop()
            self._w_trakcie = False
            self.btn_start.configure(state="normal", text="▶  Start")
            if sukces:
                self.btn_otworz.configure(state="normal")
        self.root.after(0, _ui)

    # ── OTWIERANIE PLIKU ──────────────────────────────────────────

    def _otworz_plik(self):
        if not self._ostatni_plik or not os.path.exists(self._ostatni_plik):
            messagebox.showerror("Błąd", "Plik nie istnieje.")
            return
        try:
            if sys.platform == "win32":
                os.startfile(self._ostatni_plik)
            elif sys.platform == "darwin":
                subprocess.run(["open", self._ostatni_plik])
            else:
                subprocess.run(["xdg-open", self._ostatni_plik])
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie można otworzyć pliku:\n{e}")


# ─────────────────────────────────────────────
#  URUCHOMIENIE
# ─────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app = ScraperGUI(root)
    root.mainloop()
