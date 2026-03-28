# Adeverință Notar — Tab Omise & Tab Notariat Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Adaugă două tab-uri noi în `adeverinta_notar_app.py`: "Omise" (introducere manuală dată deces pentru CNP-uri nedetectate) și "Notariat" (generare adeverință pentru birouri notariale cu OCR via Claude Vision API).

**Architecture:** UI-ul existent (un singur `CTkScrollableFrame`) se convertește la `CTkTabview` cu trei tab-uri: "Principal" (UI existent), "Omise" (nou), "Notariat" (nou). Metoda `generate_notariat()` se adaugă în `AdeverintaEngine`. OCR folosește PyMuPDF + Claude Haiku 4.5 API.

**Tech Stack:** Python, CustomTkinter, openpyxl, python-docx, PyMuPDF (fitz), anthropic SDK

---

## Chunk 1: CTkTabview + Tab Omise

### Task 1: Convertire UI la CTkTabview

**Files:**
- Modify: `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py` (metoda `__init__` și `_build_ui`)

Context: UI-ul actual constă dintr-un singur `CTkScrollableFrame` în `_build_ui`. Trebuie mutat conținutul în tab-ul "Principal" al unui `CTkTabview`.

- [ ] **Step 1: Înlocuiește frame-ul principal cu CTkTabview în `__init__`**

  Găsește în `__init__` (linia ~943):
  ```python
  self._build_ui()
  self.update_idletasks()
  self.geometry("610x780")
  self.minsize(540, 650)
  ```
  Modifică la:
  ```python
  self._cnps_omise = []
  self._omise_entries = {}
  self._build_ui()
  self.update_idletasks()
  self.geometry("650x820")
  self.minsize(580, 700)
  ```

- [ ] **Step 2: Adaugă CTkTabview în `_build_ui`, înaintea conținutului existent**

  La începutul `_build_ui()`, înlocuiește:
  ```python
  def _build_ui(self):
      A = self.ACCENT
      _scroll = ctk.CTkScrollableFrame(self, fg_color="#EDE8E0", corner_radius=0)
      _scroll.pack(fill="both", expand=True)
      wrap = ctk.CTkFrame(_scroll, fg_color="transparent")
      wrap.pack(fill="x", padx=12, pady=6)
  ```
  Cu:
  ```python
  def _build_ui(self):
      A = self.ACCENT
      self.tabview = ctk.CTkTabview(self, fg_color="#EDE8E0",
                                    segmented_button_fg_color="#D8D0C4",
                                    segmented_button_selected_color="#2B6CB0",
                                    segmented_button_selected_hover_color="#1A4E8A",
                                    text_color="#1A3F5C")
      self.tabview.pack(fill="both", expand=True, padx=4, pady=4)
      self.tabview.add("Principal")
      self.tabview.add("Omise")
      self.tabview.add("Notariat")

      # ── Tab Principal ────────────────────────────────────────────────────
      _scroll = ctk.CTkScrollableFrame(self.tabview.tab("Principal"),
                                       fg_color="#EDE8E0", corner_radius=0)
      _scroll.pack(fill="both", expand=True)
      wrap = ctk.CTkFrame(_scroll, fg_color="transparent")
      wrap.pack(fill="x", padx=12, pady=6)
  ```

- [ ] **Step 3: Lansează aplicația și verifică vizual**

  Rulează: `python "C:\EXCELURI\Adeverinte\adeverinta_notar_app.py"`

  Așteptat: fereastra se deschide cu trei tab-uri (Principal, Omise, Notariat). Tab-ul Principal conține UI-ul existent intact. Tab-urile Omise și Notariat sunt goale deocamdată.

- [ ] **Step 4: Commit**

  ```bash
  cd "C:\EXCELURI\Adeverinte"
  git add adeverinta_notar_app.py
  git commit -m "refactor: converteste UI la CTkTabview (Principal/Omise/Notariat)"
  ```

---

### Task 2: Tab Omise — UI de bază

**Files:**
- Modify: `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py` — adaugă metoda `_build_omise_tab()` și cheam-o din `_build_ui()`

- [ ] **Step 1: Adaugă metoda `_build_omise_tab()`**

  Adaugă după `_build_ui()` (înainte de `_card()`):
  ```python
  def _build_omise_tab(self):
      """Construieste continutul static al tab-ului Omise."""
      parent = self.tabview.tab("Omise")

      # Header
      self.lbl_omise_header = ctk.CTkLabel(
          parent,
          text="Tab-ul se populează după o rulare din tab-ul Principal.",
          font=ctk.CTkFont(size=12),
          text_color="#7A8FA6",
          wraplength=500,
          justify="center"
      )
      self.lbl_omise_header.pack(pady=(20, 8), padx=16)

      # Nota informativa (initial ascunsa)
      self.lbl_omise_nota = ctk.CTkLabel(
          parent,
          text="Aceste CNP-uri nu au marcajul «Beneficiar decedat» în Istoric Plăți.\n"
               "Introduceți luna și anul decesului (format MM.YYYY, ex: 06.2024).",
          font=ctk.CTkFont(size=10),
          fg_color="#E8F0F8",
          corner_radius=6,
          text_color="#2B5282",
          wraplength=500,
          justify="left"
      )
      # nu .pack() inca — se afiseaza dupa populare

      # ScrollableFrame pentru randuri CNP
      self.omise_scroll = ctk.CTkScrollableFrame(
          parent, fg_color="#FAF7F2",
          border_width=1, border_color="#D8D0C4",
          corner_radius=6, height=260
      )
      # nu .pack() inca

      # Mesaj validare (initial ascuns)
      self.lbl_omise_err = ctk.CTkLabel(
          parent, text="",
          font=ctk.CTkFont(size=10),
          text_color="#C53030"
      )

      # Buton generare
      self.btn_omise_gen = ctk.CTkButton(
          parent,
          text="Generează pentru completați",
          font=ctk.CTkFont(size=12, weight="bold"),
          fg_color="#276749", hover_color="#1C4D36",
          height=32, corner_radius=7,
          state="disabled",
          command=self._on_omise_generate
      )

      # Mini-log
      self.omise_log = ctk.CTkTextbox(
          parent, height=70,
          font=ctk.CTkFont(family="Consolas", size=10),
          fg_color="#F0F5FA", text_color="#1E3A52",
          state="disabled"
      )
  ```

- [ ] **Step 2: Cheamă `_build_omise_tab()` din `_build_ui()`**

  La finalul `_build_ui()`, înainte de ultima linie, adaugă:
  ```python
  # ── Tab Omise ────────────────────────────────────────────────────────
  self._build_omise_tab()
  ```

- [ ] **Step 3: Verifică vizual**

  Rulează aplicația. Tab-ul Omise trebuie să afișeze textul placeholder. Celelalte widget-uri sunt create dar nu afișate (vor apărea după populare).

- [ ] **Step 4: Commit**

  ```bash
  git add adeverinta_notar_app.py
  git commit -m "feat: tab Omise - UI de baza (header, scroll, buton, minilog)"
  ```

---

### Task 3: Popularea tab-ului Omise după generare

**Files:**
- Modify: `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py` — adaugă `_populate_omise_tab()` și `_omise_minilog()`; modifică `_run_generate()` pentru colectare CNP-uri omise

- [ ] **Step 1: Adaugă `_omise_minilog()`**

  ```python
  def _omise_minilog(self, msg):
      """Scrie in mini-log-ul din tab Omise (thread-safe)."""
      def _write():
          self.omise_log.configure(state="normal")
          ts = datetime.now().strftime("%H:%M:%S")
          self.omise_log.insert("end", f"[{ts}] {msg}\n")
          self.omise_log.see("end")
          self.omise_log.configure(state="disabled")
      self.after(0, _write)
  ```

- [ ] **Step 2: Adaugă `_populate_omise_tab()`**

  ```python
  def _populate_omise_tab(self):
      """Repopuleaza tab-ul Omise cu lista CNP-urilor fara deces detectat."""
      # Curata randurile vechi din scroll frame
      for w in self.omise_scroll.winfo_children():
          w.destroy()
      self._omise_entries.clear()

      n = len(self._cnps_omise)

      if n == 0:
          self.lbl_omise_header.configure(
              text="Nu există CNP-uri fără deces detectat. Toate au fost procesate."
          )
          self.lbl_omise_nota.pack_forget()
          self.omise_scroll.pack_forget()
          self.lbl_omise_err.pack_forget()
          self.btn_omise_gen.pack_forget()
          self.omise_log.pack_forget()
          return

      # Header cu numarul de CNP-uri
      self.lbl_omise_header.configure(
          text=f"● {n} CNP-uri fără deces detectat automat",
          font=ctk.CTkFont(size=12, weight="bold"),
          text_color="#744210"
      )

      # Afiseaza nota informativa
      self.lbl_omise_nota.pack(fill="x", padx=12, pady=(0, 8))

      # Titlu coloane
      hdr = ctk.CTkFrame(self.omise_scroll, fg_color="transparent")
      hdr.pack(fill="x", pady=(0, 2))
      ctk.CTkLabel(hdr, text="CNP", width=130,
                   font=ctk.CTkFont(size=10, weight="bold"),
                   text_color="#5A7A96", anchor="w").pack(side="left", padx=(4, 0))
      ctk.CTkLabel(hdr, text="Beneficiar", width=160,
                   font=ctk.CTkFont(size=10, weight="bold"),
                   text_color="#5A7A96", anchor="w").pack(side="left")
      ctk.CTkLabel(hdr, text="Data deces (MM.YYYY)",
                   font=ctk.CTkFont(size=10, weight="bold"),
                   text_color="#5A7A96", anchor="w").pack(side="left")

      # Rand per CNP
      for cnp in self._cnps_omise:
          p = self.engine.raw_data.get(cnp, {})
          nume = f"{p.get('prenume', '')} {p.get('nume', '')}".strip() or cnp

          row = ctk.CTkFrame(self.omise_scroll,
                             fg_color="#FFFFFF", corner_radius=4,
                             border_width=1, border_color="#E8E0D8")
          row.pack(fill="x", pady=2, padx=2)

          ctk.CTkLabel(row, text=cnp, width=130,
                       font=ctk.CTkFont(family="Courier", size=10),
                       text_color="#4A5568", anchor="w").pack(side="left", padx=(6, 0))
          ctk.CTkLabel(row, text=nume, width=160,
                       font=ctk.CTkFont(size=10),
                       text_color="#2D3748", anchor="w").pack(side="left")
          entry = ctk.CTkEntry(row, width=110, height=24,
                               placeholder_text="MM.YYYY",
                               font=ctk.CTkFont(family="Courier", size=10))
          entry.pack(side="left", padx=6)
          entry.bind("<KeyRelease>", self._on_omise_entry_change)
          self._omise_entries[cnp] = entry

      # Afiseaza scroll, err label, buton, minilog
      self.omise_scroll.pack(fill="both", expand=True, padx=12, pady=(0, 4))
      self.lbl_omise_err.pack(anchor="w", padx=12)
      self.btn_omise_gen.pack(fill="x", padx=12, pady=(4, 4))
      self.omise_log.pack(fill="x", padx=12, pady=(0, 8))
      self.btn_omise_gen.configure(state="disabled")
  ```

- [ ] **Step 3: Modifică `_run_generate()` pentru colectarea CNP-urilor omise**

  Înlocuiește blocul `worker()` din `_run_generate()` cu versiunea de mai jos. Modificarea constă în adăugarea colectării CNP-urilor în blocul `finally`, astfel încât să ruleze indiferent de succes sau eroare:

  ```python
  def worker():
      try:
          n_gen, n_fara, erori = self.engine.generate_all(
              tpl, out, indicativ=indicativ, nr_inreg=nr_inreg, cod_mt=cod_mt,
              filter_cnp=filter_cnp, callback=self._callback
          )
          suffix = f" (filtrat: CNP {filter_cnp})" if filter_cnp else ""
          self._callback(100,
              f"FINALIZAT: {n_gen} adeverințe generate, "
              f"{n_fara} fără restanțe, {len(erori)} erori{suffix}"
          )
          if erori:
              self._log("Erori:")
              for e in erori:
                  self._log(f"  {e}")
          self.after(0, lambda: self.btn_open.configure(state="normal"))
      except Exception as e:
          self._callback(0, f"EROARE generare: {e}")
          import traceback
          self._log(traceback.format_exc())
      finally:
          # Colecteaza CNP-urile fara deces detectat (ruleaza intotdeauna)
          self._cnps_omise = [
              cnp for cnp, p in self.engine.raw_data.items()
              if p['data_deces_raw'] is None
          ]
          if self._cnps_omise:
              self._log(f"⚠ {len(self._cnps_omise)} CNP-uri fără deces → tab «Omise»")
          self.after(0, self._populate_omise_tab)
          self.after(0, lambda: self._set_buttons(True))
  ```

- [ ] **Step 4: Testare manuală**

  1. Pornește aplicația
  2. Încarcă un raport Excel din tab-ul Principal
  3. Apasă "Generează Adeverințe"
  4. Treci la tab-ul Omise
  5. Așteptat: dacă există CNP-uri fără deces — lista se afișează cu intrări; dacă nu există — mesaj "Nu există CNP-uri"

- [ ] **Step 5: Commit**

  ```bash
  git add adeverinta_notar_app.py
  git commit -m "feat: tab Omise - populare automata dupa generare"
  ```

---

### Task 4: Validare și activare buton în Tab Omise

**Files:**
- Modify: `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py` — adaugă `_on_omise_entry_change()` și `_validate_omise_entries()`

- [ ] **Step 1: Adaugă `_on_omise_entry_change()` și `_validate_omise_entries()`**

  ```python
  def _on_omise_entry_change(self, event=None):
      """Activare/dezactivare buton dupa keystroke — fara validare regex."""
      any_filled = any(
          e.get().strip()
          for e in self._omise_entries.values()
          if e.cget("state") != "disabled"
      )
      self.btn_omise_gen.configure(
          state="normal" if any_filled else "disabled"
      )
      # Reseteaza bordurile rosii la editare
      if event and event.widget in self._omise_entries.values():
          event.widget.configure(border_color=["#979DA2", "#565B5E"])
      self.lbl_omise_err.configure(text="")

  def _validate_omise_entries(self):
      """
      Valideaza toate campurile nevide la submit.
      Returneaza dict {cnp: (mm, yyyy)} pentru cele valide
      sau None daca exista erori.
      """
      import re
      valid = {}
      has_error = False
      for cnp, entry in self._omise_entries.items():
          if entry.cget("state") == "disabled":
              continue
          val = entry.get().strip()
          if not val:
              continue
          m = re.match(r'^(\d{1,2})\.(\d{4})$', val)
          if not m:
              entry.configure(border_color="red")
              has_error = True
          else:
              entry.configure(border_color=["#979DA2", "#565B5E"])
              valid[cnp] = (m.group(1).zfill(2), m.group(2))
      if has_error:
          self.lbl_omise_err.configure(
              text="⚠ Format invalid — folosiți MM.YYYY (ex: 06.2024)"
          )
          return None
      self.lbl_omise_err.configure(text="")
      return valid
  ```

- [ ] **Step 2: Verifică comportamentul**

  Pornește aplicația. Populează tab-ul Omise (rulează o generare). Testează:
  - Câmp gol → buton dezactivat
  - Câmp cu text oarecare → buton activ
  - Apasă "Generează pentru completați" cu format invalid (ex: `6-2024`) → bordură roșie + mesaj eroare
  - Corectează câmpul (`06.2024`) → bordura revine normală

- [ ] **Step 3: Commit**

  ```bash
  git add adeverinta_notar_app.py
  git commit -m "feat: tab Omise - validare MM.YYYY si activare buton"
  ```

---

### Task 5: Generare din Tab Omise

**Files:**
- Modify: `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py` — adaugă `_on_omise_generate()` și `_run_omise_generation()`; modifică `_set_buttons()`

- [ ] **Step 1: Extinde `_set_buttons()` pentru butonul Omise**

  Înlocuiește metoda existentă:
  ```python
  def _set_buttons(self, enabled):
      state = "normal" if enabled else "disabled"
      self.btn_load.configure(state=state)
      if enabled and self.engine.raw_data:
          self.btn_gen.configure(state=state)
  ```
  Cu:
  ```python
  def _set_buttons(self, enabled):
      state = "normal" if enabled else "disabled"
      self.btn_load.configure(state=state)
      if enabled and self.engine.raw_data:
          self.btn_gen.configure(state=state)
      # Butonul Omise — reactivat doar daca mai sunt campuri nevide
      if enabled:
          any_filled = any(
              e.get().strip()
              for e in self._omise_entries.values()
              if e.cget("state") != "disabled"
          )
          if any_filled:
              self.btn_omise_gen.configure(state="normal")
      else:
          self.btn_omise_gen.configure(state="disabled")
      # Butonul Notariat — reactivat de _check_notariat_ready()
      if enabled:
          self.after(0, self._check_notariat_ready)
  ```

  **Notă:** `_check_notariat_ready()` se implementează în Task 9 (Tab Notariat). Adaugă deocamdată un stub gol:
  ```python
  def _check_notariat_ready(self):
      pass  # implementat in Task 9
  ```

- [ ] **Step 2: Adaugă `_on_omise_generate()` și `_run_omise_generation()`**

  ```python
  def _on_omise_generate(self):
      valid = self._validate_omise_entries()
      if valid is None:
          return  # eroare validare
      if not valid:
          self.lbl_omise_err.configure(text="⚠ Completați cel puțin un câmp valid.")
          return

      tpl = self.tpl_entry.get().strip()
      out = self.out_entry.get().strip()
      if not tpl or not os.path.isfile(tpl):
          self._omise_minilog("✘ Template nedefinit — configurați din tab-ul Principal.")
          return
      if not out:
          self._omise_minilog("✘ Folder output nedefinit — configurați din tab-ul Principal.")
          return

      cod_mt    = self.cod_mt_entry.get().strip() or 'mt'
      indicativ = self.indicativ_var.get().strip()
      nr_inreg  = self.nr_inreg_entry.get().strip()

      self._set_buttons(False)
      threading.Thread(
          target=self._run_omise_generation,
          args=(valid, tpl, out, indicativ, nr_inreg, cod_mt),
          daemon=True
      ).start()

  def _run_omise_generation(self, valid, tpl, out, indicativ, nr_inreg, cod_mt):
      """Genereaza adeverinte pentru CNP-urile cu data deces introdusa manual."""
      for cnp, (mm, yyyy) in valid.items():
          self.engine.raw_data[cnp]['data_deces_raw'] = f"01.{mm}.{yyyy}"
          try:
              n_gen, n_fara, erori = self.engine.generate_all(
                  tpl, out,
                  indicativ=indicativ, nr_inreg=nr_inreg, cod_mt=cod_mt,
                  filter_cnp=cnp,
                  callback=lambda pct, msg: self._omise_minilog(msg)
              )
              if erori:
                  self._omise_minilog(f"✘ {cnp}: {erori[0]}")
              elif n_gen == 0:
                  self._omise_minilog(f"⚠ {cnp}: fără restanțe — adeverință negenerată")
              else:
                  self._omise_minilog(f"✔ {cnp}: adeverință generată")
                  entry = self._omise_entries.get(cnp)
                  if entry:
                      self.after(0, lambda e=entry: e.configure(
                          state="disabled", text_color="gray"
                      ))
          except Exception as ex:
              self._omise_minilog(f"✘ {cnp}: eroare — {ex}")
      self.after(0, lambda: self._set_buttons(True))
  ```

- [ ] **Step 3: Test end-to-end**

  1. Încarcă Excel → generează din Principal
  2. În tab Omise, dacă există CNP-uri, introdu `06.2024` într-un câmp
  3. Apasă "Generează pentru completați"
  4. Așteptat: mini-log afișează `✔ CNP: adeverință generată` sau `⚠ fără restanțe`
  5. Câmpul devine gri/dezactivat după succes
  6. Butonul "Generează" din Principal e dezactivat pe durata generării

- [ ] **Step 4: Commit**

  ```bash
  git add adeverinta_notar_app.py
  git commit -m "feat: tab Omise - generare adeverinte cu data deces manuala"
  ```

---

## Chunk 2: Tab Notariat

### Task 6: Tab Notariat — UI

**Files:**
- Modify: `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py` — adaugă `_build_notariat_tab()`

- [ ] **Step 1: Adaugă `_build_notariat_tab()`**

  ```python
  def _build_notariat_tab(self):
      """Construieste UI-ul tab-ului Notariat."""
      A = self.ACCENT
      parent = self.tabview.tab("Notariat")
      scroll = ctk.CTkScrollableFrame(parent, fg_color="#EDE8E0", corner_radius=0)
      scroll.pack(fill="both", expand=True)
      wrap = ctk.CTkFrame(scroll, fg_color="transparent")
      wrap.pack(fill="x", padx=12, pady=6)

      def card(title, accent):
          outer = ctk.CTkFrame(wrap, fg_color="#FAF7F2", corner_radius=7,
                               border_width=1, border_color="#D8D0C4")
          outer.pack(fill="x", pady=(4, 0))
          stripe = ctk.CTkFrame(outer, width=4, fg_color=accent, corner_radius=2)
          stripe.pack(side="left", fill="y"); stripe.pack_propagate(False)
          inner = ctk.CTkFrame(outer, fg_color="transparent")
          inner.pack(side="left", fill="x", expand=True, padx=8, pady=4)
          ctk.CTkLabel(inner, text=title,
                       font=ctk.CTkFont(size=11, weight="bold"),
                       text_color=accent).pack(anchor="w")
          return inner

      def row_entry(parent, label, placeholder, width=None):
          r = ctk.CTkFrame(parent, fg_color="transparent")
          r.pack(fill="x", pady=2)
          ctk.CTkLabel(r, text=label, width=130, anchor="w",
                       font=ctk.CTkFont(size=10), text_color="#4A6070").pack(side="left")
          kw = {"height": 26, "font": ctk.CTkFont(size=10),
                "placeholder_text": placeholder}
          if width:
              kw["width"] = width
          e = ctk.CTkEntry(r, **kw)
          e.pack(side="left", fill="x", expand=True)
          return e

      # ── Template Notariat ───────────────────────────────────────────────
      c0 = card("Template Notariat (.docx)", A["template"])
      r0 = ctk.CTkFrame(c0, fg_color="transparent")
      r0.pack(fill="x", pady=(2, 0))
      self.not_tpl_entry = ctk.CTkEntry(r0, height=26, font=ctk.CTkFont(size=10),
                                         placeholder_text="Adeverinta Notariat master.docx")
      # Auto-detect template langa script
      default_not_tpl = os.path.join(
          os.path.dirname(os.path.abspath(__file__)),
          "Adeverinta Notariat master.docx"
      )
      if os.path.exists(default_not_tpl):
          self.not_tpl_entry.insert(0, default_not_tpl)
      self.not_tpl_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
      ctk.CTkButton(r0, text="Browse", width=72, height=26,
                    fg_color=A["template"], hover_color="#5E4A7A",
                    command=self._browse_notariat_template).pack(side="right")

      # ── Solicitare PDF ──────────────────────────────────────────────────
      c1 = card("Solicitare PDF (opțional — OCR auto)", "#1A7A8A")
      rp = ctk.CTkFrame(c1, fg_color="transparent")
      rp.pack(fill="x", pady=(2, 0))
      self.not_pdf_label = ctk.CTkLabel(rp, text="Niciun PDF selectat",
                                         font=ctk.CTkFont(size=10), text_color="#7A8FA6")
      self.not_pdf_label.pack(side="left", fill="x", expand=True)
      ctk.CTkButton(rp, text="Încarcă PDF", width=90, height=26,
                    fg_color="#1A7A8A", hover_color="#135F6D",
                    command=self._browse_notariat_pdf).pack(side="right")
      self.not_ocr_status = ctk.CTkLabel(c1, text="",
                                          font=ctk.CTkFont(size=10),
                                          text_color="#276749")
      self.not_ocr_status.pack(anchor="w")

      # ── Birou Notarial ──────────────────────────────────────────────────
      c2 = card("Birou Notarial", "#744210")
      self.not_birou_entry    = row_entry(c2, "Denumire birou:",  "Societate Profesională Notarială...")
      self.not_email_entry    = row_entry(c2, "Email birou:",     "email@notariat.ro")
      self.not_notar_entry    = row_entry(c2, "Nume notar:",      "Prenume Nume")

      rg = ctk.CTkFrame(c2, fg_color="transparent")
      rg.pack(fill="x", pady=2)
      ctk.CTkLabel(rg, text="Gen notar:", width=130, anchor="w",
                   font=ctk.CTkFont(size=10), text_color="#4A6070").pack(side="left")
      self.not_gen_seg = ctk.CTkSegmentedButton(rg, values=["Doamnă", "Domnule"],
                                                 font=ctk.CTkFont(size=10))
      self.not_gen_seg.set("Doamnă")
      self.not_gen_seg.pack(side="left")

      rns = ctk.CTkFrame(c2, fg_color="transparent")
      rns.pack(fill="x", pady=2)
      ctk.CTkLabel(rns, text="Nr. solicitare ext.:", width=130, anchor="w",
                   font=ctk.CTkFont(size=10), text_color="#4A6070").pack(side="left")
      self.not_nr_sol_entry   = ctk.CTkEntry(rns, width=80, height=26,
                                              placeholder_text="498",
                                              font=ctk.CTkFont(size=10))
      self.not_nr_sol_entry.pack(side="left", padx=(0, 10))
      ctk.CTkLabel(rns, text="Dată:", font=ctk.CTkFont(size=10)).pack(side="left", padx=(0, 4))
      self.not_data_sol_entry = ctk.CTkEntry(rns, width=90, height=26,
                                              placeholder_text="12.03.2026",
                                              font=ctk.CTkFont(size=10))
      self.not_data_sol_entry.pack(side="left")

      rni = ctk.CTkFrame(c2, fg_color="transparent")
      rni.pack(fill="x", pady=2)
      ctk.CTkLabel(rni, text="Nr. înregistrare int.:", width=130, anchor="w",
                   font=ctk.CTkFont(size=10), text_color="#4A6070").pack(side="left")
      self.not_nr_inreg_entry  = ctk.CTkEntry(rni, width=80, height=26,
                                               placeholder_text="43/659",
                                               font=ctk.CTkFont(size=10))
      self.not_nr_inreg_entry.pack(side="left", padx=(0, 10))
      ctk.CTkLabel(rni, text="Dată:", font=ctk.CTkFont(size=10)).pack(side="left", padx=(0, 4))
      self.not_data_inreg_entry = ctk.CTkEntry(rni, width=90, height=26,
                                                placeholder_text="13.03.2026",
                                                font=ctk.CTkFont(size=10))
      self.not_data_inreg_entry.pack(side="left")

      # ── Beneficiar ──────────────────────────────────────────────────────
      c3 = card("Beneficiar", A["cnp"])
      rcnp = ctk.CTkFrame(c3, fg_color="transparent")
      rcnp.pack(fill="x", pady=(2, 0))
      ctk.CTkLabel(rcnp, text="CNP:", width=130, anchor="w",
                   font=ctk.CTkFont(size=10), text_color="#4A6070").pack(side="left")
      self.not_cnp_entry = ctk.CTkEntry(rcnp, width=140, height=26,
                                         placeholder_text="1234567890123",
                                         font=ctk.CTkFont(family="Courier", size=10))
      self.not_cnp_entry.pack(side="left", padx=(0, 6))
      ctk.CTkButton(rcnp, text="Caută", width=60, height=26,
                    fg_color=A["cnp"], hover_color="#4A4E8A",
                    command=self._cauta_cnp_notariat).pack(side="left")

      self.not_beneficiar_frame = ctk.CTkFrame(c3, fg_color="#F0F5FA",
                                                corner_radius=5, border_width=1,
                                                border_color="#D0DAE8")
      self.not_beneficiar_frame.pack(fill="x", pady=(4, 0))
      self.not_beneficiar_lbl = ctk.CTkLabel(
          self.not_beneficiar_frame,
          text="Introduceți CNP-ul și apăsați «Caută».",
          font=ctk.CTkFont(size=10), text_color="#7A8FA6",
          justify="left", anchor="w", wraplength=480
      )
      self.not_beneficiar_lbl.pack(anchor="w", padx=8, pady=4)

      # ── Butoane + status ────────────────────────────────────────────────
      self.btn_not_gen = ctk.CTkButton(
          wrap,
          text="Generează adeverință notarială",
          font=ctk.CTkFont(size=12, weight="bold"),
          fg_color="#744210", hover_color="#5A3210",
          height=34, corner_radius=7,
          state="disabled",
          command=self._on_notariat_generate
      )
      self.btn_not_gen.pack(fill="x", pady=(8, 2))

      self.not_status_lbl = ctk.CTkLabel(wrap, text="",
                                          font=ctk.CTkFont(size=10),
                                          text_color="#276749")
      self.not_status_lbl.pack(anchor="w")

      # Stoca referinta la CNP gasit
      self._not_cnp_found = None

      # Bind pentru activare buton
      for widget in [self.not_birou_entry, self.not_nr_sol_entry,
                     self.not_cnp_entry]:
          widget.bind("<KeyRelease>", lambda e: self.after(10, self._check_notariat_ready))
  ```

- [ ] **Step 2: Cheamă `_build_notariat_tab()` din `_build_ui()`**

  La finalul `_build_ui()`, după `self._build_omise_tab()`, adaugă:
  ```python
  self._build_notariat_tab()
  ```

- [ ] **Step 3: Adaugă metoda `_browse_notariat_template()`**

  ```python
  def _browse_notariat_template(self):
      path = filedialog.askopenfilename(
          title="Template adeverință notariat",
          filetypes=[("Word", "*.docx"), ("Toate", "*.*")]
      )
      if path:
          self.not_tpl_entry.delete(0, "end")
          self.not_tpl_entry.insert(0, path)
          self._check_notariat_ready()
  ```

- [ ] **Step 4: Verifică vizual**

  Rulează aplicația. Tab-ul Notariat trebuie să afișeze cardurile cu toate câmpurile.

- [ ] **Step 5: Commit**

  ```bash
  git add adeverinta_notar_app.py
  git commit -m "feat: tab Notariat - UI complet (template, PDF, birou, beneficiar)"
  ```

---

### Task 7: Tab Notariat — OCR via Claude Vision

**Files:**
- Modify: `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py` — adaugă `_browse_notariat_pdf()`, `_ocr_extract_pdf()`, `_apply_ocr_result()`

- [ ] **Step 1: Adaugă constanta `PROMPT_OCR_NOTARIAT` la nivel de modul**

  Adaugă după importuri, înainte de `LUNI_RO`:
  ```python
  PROMPT_OCR_NOTARIAT = (
      "Extrage din această imagine a unui document notarial românesc câmpurile de mai jos.\n"
      "Returnează DOAR un obiect JSON valid, fără text suplimentar:\n"
      "{\n"
      '  "denumire_birou": "denumirea completă a biroului/societății notariale",\n'
      '  "email": "adresa de email",\n'
      '  "nr_solicitare": "numărul din antet (doar cifre, ex: 498)",\n'
      '  "data_solicitare": "data din antet format DD.MM.YYYY",\n'
      '  "cnp_beneficiar": "CNP-ul persoanei decedate (13 cifre)",\n'
      '  "nume_beneficiar": "numele complet al defunctului"\n'
      "}\n"
      "Dacă un câmp nu poate fi identificat cu certitudine, folosește null."
  )
  ```

- [ ] **Step 2: Adaugă `_browse_notariat_pdf()`**

  ```python
  def _browse_notariat_pdf(self):
      path = filedialog.askopenfilename(
          title="Selectează solicitarea PDF de la biroul notarial",
          filetypes=[("PDF", "*.pdf"), ("Toate", "*.*")]
      )
      if not path:
          return
      name = os.path.basename(path)
      self.not_pdf_label.configure(text=name, text_color="#2D3748")
      self.not_ocr_status.configure(text="⏳ OCR în curs...", text_color="#744210")
      threading.Thread(
          target=self._ocr_extract_pdf,
          args=(path,),
          daemon=True
      ).start()
  ```

- [ ] **Step 3: Adaugă `_ocr_extract_pdf()`**

  ```python
  def _ocr_extract_pdf(self, pdf_path):
      """Extrage campuri din PDF via Claude Vision API (thread separat)."""
      import json
      api_key = os.environ.get('ANTHROPIC_API_KEY')
      if not api_key:
          self.after(0, lambda: self.not_ocr_status.configure(
              text="⚠ ANTHROPIC_API_KEY negăsit — completați manual.",
              text_color="#C53030"
          ))
          return
      try:
          import fitz
          import anthropic
          import base64
          doc = fitz.open(pdf_path)
          n_pages = doc.page_count
          pix = doc[0].get_pixmap(dpi=200)
          img_b64 = base64.standard_b64encode(pix.tobytes("png")).decode()
          doc.close()

          client = anthropic.Anthropic(api_key=api_key)
          response = client.messages.create(
              model="claude-haiku-4-5-20251001",
              max_tokens=512,
              messages=[{"role": "user", "content": [
                  {"type": "image", "source": {
                      "type": "base64", "media_type": "image/png", "data": img_b64
                  }},
                  {"type": "text", "text": PROMPT_OCR_NOTARIAT}
              ]}]
          )
          data = json.loads(response.content[0].text)
          note = f"pagina 1 din {n_pages}" if n_pages > 1 else "1 pagină"
          self.after(0, lambda d=data, nt=note: self._apply_ocr_result(d, nt))
      except json.JSONDecodeError:
          self.after(0, lambda: self.not_ocr_status.configure(
              text="⚠ OCR: răspuns non-JSON — completați manual.",
              text_color="#C53030"
          ))
      except Exception as ex:
          self.after(0, lambda e=str(ex): self.not_ocr_status.configure(
              text=f"⚠ OCR indisponibil: {e[:80]} — completați manual.",
              text_color="#C53030"
          ))
  ```

- [ ] **Step 4: Adaugă `_apply_ocr_result()`**

  ```python
  def _apply_ocr_result(self, data, page_note=""):
      """Pre-completeaza campurile din rezultatul OCR (thread principal)."""
      mapping = {
          'denumire_birou': self.not_birou_entry,
          'email':          self.not_email_entry,
          'nr_solicitare':  self.not_nr_sol_entry,
          'data_solicitare': self.not_data_sol_entry,
      }
      filled = 0
      for key, widget in mapping.items():
          val = data.get(key)
          if val:
              widget.delete(0, "end")
              widget.insert(0, str(val))
              filled += 1

      # CNP — auto-trigger cautare
      cnp_val = data.get('cnp_beneficiar')
      if cnp_val:
          self.not_cnp_entry.delete(0, "end")
          self.not_cnp_entry.insert(0, str(cnp_val))
          self._cauta_cnp_notariat()
          filled += 1

      note = f" ({page_note})" if page_note else ""
      self.not_ocr_status.configure(
          text=f"✔ OCR: {filled} câmpuri completate{note}.",
          text_color="#276749"
      )
      self._check_notariat_ready()
  ```

- [ ] **Step 5: Test OCR**

  1. Setează `ANTHROPIC_API_KEY` în mediu (`set ANTHROPIC_API_KEY=sk-ant-...` în terminal)
  2. Rulează aplicația din același terminal
  3. Mergi la tab Notariat → "Încarcă PDF"
  4. Selectează un PDF de solicitare notarială (ex: orice cerere primită de la un birou notarial)
  5. Așteptat: câmpurile se pre-completează (denumire birou, email, nr solicitare, dată, CNP)

  Dacă API_KEY lipsește: mesaj `"⚠ ANTHROPIC_API_KEY negăsit — completați manual."` fără crash.

  **Notă:** Folosește orice PDF de solicitare notarială disponibil. Formatul așteptat: antet cu denumire birou, email, număr/dată solicitare și CNP beneficiar.

- [ ] **Step 6: Commit**

  ```bash
  git add adeverinta_notar_app.py
  git commit -m "feat: tab Notariat - OCR via Claude Vision API (Haiku 4.5)"
  ```

---

### Task 8: Tab Notariat — Căutare CNP

**Files:**
- Modify: `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py` — adaugă `_cauta_cnp_notariat()` și `_check_notariat_ready()`

- [ ] **Step 1: Adaugă `_cauta_cnp_notariat()`**

  ```python
  def _cauta_cnp_notariat(self):
      """Cauta CNP in engine.raw_data si afiseaza datele beneficiarului."""
      cnp = self.not_cnp_entry.get().strip()
      if not cnp:
          self.not_beneficiar_lbl.configure(
              text="Introduceți CNP-ul și apăsați «Caută».",
              text_color="#7A8FA6"
          )
          self._not_cnp_found = None
          self._check_notariat_ready()
          return

      if not self.engine.raw_data:
          self.not_beneficiar_lbl.configure(
              text="⚠ Raportul Excel nu este încărcat. Mergeți la tab-ul Principal.",
              text_color="#C53030"
          )
          self._not_cnp_found = None
          self._check_notariat_ready()
          return

      p = self.engine.raw_data.get(cnp)
      if not p:
          self.not_beneficiar_lbl.configure(
              text=f"⚠ CNP {cnp} negăsit în raport.",
              text_color="#C53030"
          )
          self._not_cnp_found = None
          self._check_notariat_ready()
          return

      pd = self.engine.personal_data.get(cnp, {})
      outstanding = self.engine.compute_outstanding(cnp)

      lines = [
          f"Nume: {p.get('prenume', '')} {p.get('nume', '')}",
          f"Dosar: {p.get('nr_dosar', '—')} | Înregistrat: {format_data_inreg(p.get('data_inreg_raw'))}",
          f"CI: {pd.get('ci_serie', '[LIPSĂ]')} {pd.get('ci_nr', '')} | Data: {pd.get('data_ci', '[LIPSĂ]')}",
          f"Adresă: {pd.get('adresa', '[LIPSĂ]')}",
      ]
      if outstanding:
          lines.append(
              f"Restanțe: {outstanding['suma']} lei | Luni: {outstanding['luni_text']}"
          )
      else:
          lines.append("⚠ Fără restanțe detectate (adeverința se va genera cu suma 0).")

      self.not_beneficiar_lbl.configure(
          text="\n".join(lines),
          text_color="#2D3748"
      )
      self._not_cnp_found = cnp
      self._check_notariat_ready()
  ```

- [ ] **Step 2: Înlocuiește stub-ul `_check_notariat_ready()` (adăugat în Task 5, Chunk 1) cu implementarea completă**

  Găsește în cod metoda stub:
  ```python
  def _check_notariat_ready(self):
      pass  # implementat in Task 9
  ```
  **Șterge-o complet** și adaugă în locul ei:

  ```python
  def _check_notariat_ready(self):
      """Activează butonul Generează Notariat dacă toate condițiile sunt îndeplinite."""
      tpl_ok  = bool(self.not_tpl_entry.get().strip())
      out_ok  = bool(self.out_entry.get().strip())
      cnp_ok  = self._not_cnp_found is not None
      birou_ok = bool(self.not_birou_entry.get().strip())
      nr_ok   = bool(self.not_nr_sol_entry.get().strip())
      ready   = tpl_ok and out_ok and cnp_ok and birou_ok and nr_ok

      self.btn_not_gen.configure(state="normal" if ready else "disabled")

      if not out_ok:
          self.not_status_lbl.configure(
              text="Setați folderul output în tab-ul Principal.",
              text_color="#C53030"
          )
      else:
          self.not_status_lbl.configure(text="")
  ```

- [ ] **Step 3: Test căutare**

  1. Încarcă Excel din Principal
  2. Mergi la Notariat → tastează un CNP existent → apasă Caută
  3. Așteptat: datele beneficiarului apar în frame-ul read-only
  4. CNP inexistent → mesaj eroare roșu
  5. Completează birou + nr solicitare → butonul "Generează" devine activ

- [ ] **Step 4: Commit**

  ```bash
  git add adeverinta_notar_app.py
  git commit -m "feat: tab Notariat - cautare CNP si activare buton generare"
  ```

---

### Task 9: `generate_notariat()` în Engine + Generare din UI

**Files:**
- Modify: `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py` — adaugă `generate_notariat()` în `AdeverintaEngine`; adaugă `_on_notariat_generate()` și `_run_notariat_generation()` în `App`

- [ ] **Step 1: Adaugă `generate_notariat()` în `AdeverintaEngine`**

  Adaugă după metoda `generate_all()` (linia ~724):
  ```python
  def generate_notariat(self, cnp, template_path, out_dir,
                        outstanding, birou, email, salut,
                        nr_solicitare_ext, nr_inreg_int):
      """
      Genereaza o adeverinta pentru biroul notarial din template cu placeholdere.
      Returneaza calea fisierului generat.
      """
      from docx.oxml.ns import qn as _qn
      p = self.raw_data[cnp]
      pd = self.personal_data.get(cnp, {})

      if outstanding:
          suma      = str(outstanding['suma'])
          luni_text = outstanding['luni_text']
          perioada  = f"{outstanding['period_start']} – {outstanding['period_end']}"
          data_deces = outstanding['data_deces']
      else:
          suma = '0'; luni_text = '—'; perioada = '—'
          data_deces = format_data_deces(p['data_deces_raw']) or '—'

      replacements = {
          '{{DENUMIRE_BIROU}}':    birou,
          '{{EMAIL_BIROU}}':       email,
          '{{SALUT}}':             salut,
          '{{NR_SOLICITARE_EXT}}': nr_solicitare_ext,
          '{{NR_INREG_INT}}':      nr_inreg_int,
          '{{NUME_PRENUME}}':      f"{p['prenume']} {p['nume']}",
          '{{CNP}}':               cnp,
          '{{DOMICILIU}}':         pd.get('adresa', '[LIPSĂ]'),
          '{{CI_SERIE}}':          pd.get('ci_serie', '[LIPSĂ]'),
          '{{CI_NR}}':             pd.get('ci_nr', '[LIPSĂ]'),
          '{{CI_DATA}}':           pd.get('data_ci', '[LIPSĂ]'),
          '{{PERIOADA}}':          perioada,
          '{{SUMA}}':              suma,
          '{{LUNI_TEXT}}':         luni_text,
          '{{NR_DOSAR}}':          p['nr_dosar'],
          '{{DATA_INREG}}':        format_data_inreg(p['data_inreg_raw']),
          '{{DATA_DECES}}':        data_deces,
      }

      doc = Document(template_path)

      # Elimina mail merge (la fel ca _generate_doc)
      settings_el = doc.settings.element
      for tag in ('w:mailMerge', 'w:attachedTemplate'):
          el = settings_el.find(_qn(tag))
          if el is not None:
              settings_el.remove(el)

      def _subst_para(para):
          # Nota: se foloseste _set_para_text (nu _replace_para_content) deoarece
          # template-ul Notariat este creat de utilizator ca fisier nou, fara runs
          # fragmentate. _set_para_text pastreaza stilul primului run si inlocuieste
          # tot textul — suficient pentru un template curat.
          for ph, val in replacements.items():
              if ph in para.text:
                  _set_para_text(para, para.text.replace(ph, val))

      for para in doc.paragraphs:
          _subst_para(para)
      for table in doc.tables:
          for row in table.rows:
              for cell in row.cells:
                  for para in cell.paragraphs:
                      _subst_para(para)

      # Filename: NOTARIAT - Prenume Nume - nr_sol.docx
      nr_sol_raw = nr_solicitare_ext.replace('Nr.', '').split('/')[0].strip()
      nr_sol_raw = re.sub(r'[<>:"/\\|?*]', '_', nr_sol_raw)
      prenume_s = re.sub(r'[^\w\s-]', '', p['prenume']).strip()[:15]
      nume_s    = re.sub(r'[^\w\s-]', '', p['nume']).strip()[:20]
      filename  = f"NOTARIAT - {prenume_s} {nume_s} - {nr_sol_raw}.docx"
      out_path  = os.path.join(out_dir, filename)
      doc.save(out_path)
      return out_path
  ```

- [ ] **Step 2: Adaugă `_on_notariat_generate()` și `_run_notariat_generation()` în `App`**

  ```python
  def _on_notariat_generate(self):
      cnp = self._not_cnp_found
      if not cnp:
          return
      tpl = self.not_tpl_entry.get().strip()
      out = self.out_entry.get().strip()
      if not tpl or not os.path.isfile(tpl):
          self.not_status_lbl.configure(
              text="⚠ Template invalid.", text_color="#C53030"
          )
          return
      if not os.path.isdir(out):
          os.makedirs(out, exist_ok=True)

      birou    = self.not_birou_entry.get().strip()
      email    = self.not_email_entry.get().strip()
      notar    = self.not_notar_entry.get().strip()
      gen      = self.not_gen_seg.get()
      nr_sol   = self.not_nr_sol_entry.get().strip()
      data_sol = self.not_data_sol_entry.get().strip()
      nr_inreg = self.not_nr_inreg_entry.get().strip()
      data_inreg = self.not_data_inreg_entry.get().strip()

      # Construieste placeholderele compuse
      nr_sol_ext  = f"Nr.{nr_sol}/{data_sol}" if nr_sol and data_sol else f"Nr.{nr_sol}"
      nr_inreg_int = f"{nr_inreg}/{data_inreg}" if nr_inreg and data_inreg else nr_inreg
      prefix = "Stimată" if gen == "Doamnă" else "Stimate"
      salut  = f"{prefix} {gen} {notar}".strip()

      outstanding = self.engine.compute_outstanding(cnp)

      self._set_buttons(False)
      self.not_status_lbl.configure(text="⏳ Generez...", text_color="#744210")
      threading.Thread(
          target=self._run_notariat_generation,
          args=(cnp, tpl, out, outstanding, birou, email, salut, nr_sol_ext, nr_inreg_int),
          daemon=True
      ).start()

  def _run_notariat_generation(self, cnp, tpl, out, outstanding,
                                birou, email, salut, nr_sol_ext, nr_inreg_int):
      try:
          out_path = self.engine.generate_notariat(
              cnp=cnp, template_path=tpl, out_dir=out,
              outstanding=outstanding, birou=birou, email=email,
              salut=salut, nr_solicitare_ext=nr_sol_ext, nr_inreg_int=nr_inreg_int
          )
          self.after(0, lambda p=out_path: self.not_status_lbl.configure(
              text=f"✔ Generat: {os.path.basename(p)}", text_color="#276749"
          ))
          self.after(0, lambda: self.btn_open.configure(state="normal"))
      except Exception as ex:
          self.after(0, lambda e=str(ex): self.not_status_lbl.configure(
              text=f"✘ Eroare: {e[:100]}", text_color="#C53030"
          ))
      finally:
          self.after(0, lambda: self._set_buttons(True))
  ```

- [ ] **Step 3: Test end-to-end Tab Notariat**

  Pregătire: creează `Adeverinta Notariat master.docx` lângă script cu placeholder-ele `{{DENUMIRE_BIROU}}`, `{{SALUT}}`, `{{NR_SOLICITARE_EXT}}` etc. (copiază structura din exemplul `ADEV PTR NOTARIAT - Constantin Marin.docx` și înlocuiește conținutul cu placeholder-e).

  1. Pornește aplicația
  2. Încarcă raportul Excel din Principal
  3. Mergi la Notariat
  4. Încarcă PDF solicitare → verifică pre-completare câmpuri
  5. Completează manual câmpurile lipsă (dacă e cazul)
  6. Caută CNP → verifică afișarea datelor beneficiarului
  7. Apasă "Generează adeverință notarială"
  8. Așteptat: fișier `NOTARIAT - Prenume Nume - NrSol.docx` în folderul output

- [ ] **Step 4: Commit final**

  ```bash
  git add adeverinta_notar_app.py
  git commit -m "feat: tab Notariat - generate_notariat() + generare completa cu OCR"
  ```

---

## Note finale

**Dependință externă nouă:** `anthropic` SDK. Dacă nu e instalat:
```bash
pip install anthropic
```

**Template Notariat master.docx** trebuie creat de utilizator. Lista completă de placeholder-e:
- `{{DENUMIRE_BIROU}}`, `{{EMAIL_BIROU}}`, `{{SALUT}}`, `{{NR_SOLICITARE_EXT}}`, `{{NR_INREG_INT}}`
- `{{NUME_PRENUME}}`, `{{CNP}}`, `{{DOMICILIU}}`, `{{CI_SERIE}}`, `{{CI_NR}}`, `{{CI_DATA}}`
- `{{PERIOADA}}`, `{{SUMA}}`, `{{LUNI_TEXT}}`, `{{NR_DOSAR}}`, `{{DATA_INREG}}`, `{{DATA_DECES}}`

**ANTHROPIC_API_KEY:** setat ca variabilă de mediu înainte de lansarea aplicației.
