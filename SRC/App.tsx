import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import {
  Plus, Trash2, X, Search, Users, Package, FileText, Upload,
  Loader2, AlertCircle, ChevronLeft, Check, Ruler, Printer, Building2, Copy, Star, Download,
} from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

/* ============================================================
   KERN-DATAMODELLEN
   ============================================================ */
export type LettertypeKey = "plex" | "inter" | "grotesk";
export type MargeOptie = "compact" | "normaal" | "ruim";
export type UitlijningOptie = "links" | "gecentreerd";
export type TabelstijlOptie = "lijnen" | "minimal";
export type HandtekeningPositie = "onder-totalen" | "onder-klant";

export interface Klant {
  id: string;
  naam: string;
  adres: string;
  postcode: string;
  plaats: string;
  email: string;
  telefoon: string;
}

export interface Product {
  id: string;
  omschrijving: string;
  eenheid: string;
  /** Exclusief btw. */
  prijs: number;
  /** Percentage, doorgaans 0 | 9 | 21. */
  btw: number;
}

export interface OfferteRegel {
  id: string;
  productId: string | null;
  omschrijving: string;
  eenheid: string;
  aantal: number;
  prijs: number;
  /** Percentage. */
  korting: number;
  btw: number;
}

export interface Offerte {
  /** null = nog niet opgeslagen concept. */
  id: string | null;
  nummer: string;
  klantId: string;
  sjabloonId?: string | null;
  /** ISO yyyy-mm-dd. */
  datum: string;
  notities: string;
  regels: OfferteRegel[];
}

export interface Sjabloon {
  id: string;
  naam: string;
  isStandaard: boolean;
  bedrijfsnaam: string;
  adres: string;
  postcode: string;
  plaats: string;
  email: string;
  telefoon: string;
  website: string;
  kvk: string;
  btwNummer: string;
  iban: string;
  geldigheidsdagen: number;
  algemeneVoorwaarden: string;
  ondertekenaar: string;
  logo: string | null;
  accentKleur: string;
  lettertype: LettertypeKey;
  voettekst: string;
  secundaireKleur: string;
  marge: MargeOptie;
  uitlijning: UitlijningOptie;
  tabelstijl: TabelstijlOptie;
  offertePrefix: string;
  standaardBtw: number;
  valuta: string;
  handtekeningAfbeelding: string | null;
  handtekeningFunctie: string;
  handtekeningPositie: HandtekeningPositie;
  layoutTemplate: string;
}

/* De losstaande "Instellingen" uit Fase 5/6 bestaat niet meer als apart model sinds Fase 9 —
   die velden zitten nu allemaal in Sjabloon (elk sjabloon draagt zijn eigen bedrijfsgegevens/
   huisstijl/documentinstellingen). Instellingen is hier een alias, geen duplicaat datamodel,
   zodat de gevraagde naam toch als type bestaat. */
export type Instellingen = Sjabloon;

export interface ColumnMapping {
  omschrijving: string | null;
  eenheid: string | null;
  prijs: string | null;
  btw: string | null;
}

/* ============================================================
   DATABASE — schema & persistence (window.storage, personal scope)
   ============================================================ */

const KEYS: Record<"klanten" | "producten" | "offertes" | "instellingen" | "sjablonen", string> = {
  klanten: "klanten", producten: "producten", offertes: "offertes", instellingen: "instellingen", sjablonen: "sjablonen",
};
const FONT_LINK_ID = "offerte-app-fonts";
const BTW_OPTIONS: number[] = [21, 9, 0];

/* Fase 9 — default layout-template als platte tekst (geen JSX-hardcoding).
   {{Placeholder}} alleen op zijn eigen regel wordt vervangen door een tekstwaarde; regels die na
   substitutie leeg zijn worden overgeslagen (dus een lege KlantAdres-regel toont niets, zonder
   dat elk veld apart if-if-if hoeft te worden gecodeerd). Logo/OfferteRegels/Handtekening-* zijn
   "blok"-placeholders: die renderen een echt component (tabel, afbeelding, handtekeningregel) in
   plaats van tekst — nooit ruwe HTML, dus geen dangerouslySetInnerHTML en geen scriptinjectierisico. */
const DEFAULT_LAYOUT_TEMPLATE = `{{Logo}}
{{Bedrijfsnaam}}
{{Adres}}
{{PostcodePlaats}}
{{Email}}
{{Telefoon}}

Offerte {{OfferteNummer}}
Datum: {{Datum}}
Geldig tot: {{GeldigTot}}

Klant
{{KlantNaam}}
{{KlantAdres}}
{{KlantPostcodePlaats}}

{{OfferteRegels}}

{{Totalen}}

{{Notities}}

Algemene voorwaarden
{{AlgemeneVoorwaarden}}

{{HandtekeningKlant}}
{{HandtekeningOndertekenaar}}

{{Voettekst}}
{{Footer}}`;

const DEFAULT_SJABLOON_VELDEN: Omit<Sjabloon, "id" | "naam" | "isStandaard"> = {
  bedrijfsnaam: "", adres: "", postcode: "", plaats: "", email: "", telefoon: "", website: "",
  kvk: "", btwNummer: "", iban: "", geldigheidsdagen: 30, algemeneVoorwaarden: "", ondertekenaar: "",
  logo: null, accentKleur: "#C97A1A", lettertype: "plex", voettekst: "",
  // Fase 9 — huisstijl, documentinstellingen, handtekening
  secundaireKleur: "#EDEBE6", marge: "normaal", uitlijning: "links", tabelstijl: "lijnen",
  offertePrefix: "OFF-", standaardBtw: 21, valuta: "EUR",
  handtekeningAfbeelding: null, handtekeningFunctie: "", handtekeningPositie: "onder-totalen",
  layoutTemplate: DEFAULT_LAYOUT_TEMPLATE,
};
function nieuwSjabloon(naam: string, overrides: Partial<Sjabloon> = {}): Sjabloon {
  return { id: uid(), naam, isStandaard: false, ...DEFAULT_SJABLOON_VELDEN, ...overrides };
}

/* Fase 6 — curated lettertype-presets (geen vrij URL-veld: begrensde, betrouwbare set).
   Het mono-lettertype voor bedragen blijft altijd IBM Plex Mono, ongeacht de keuze hier —
   dat houdt prijzen/aantallen visueel consistent en is een bewuste scope-keuze. */
interface FontPreset {
  label: string;
  body: string;
  googleFontsUrl: string;
}
const FONT_PRESETS: Record<LettertypeKey, FontPreset> = {
  plex: { label: "IBM Plex (huidig)", body: "'IBM Plex Sans', sans-serif", googleFontsUrl: "https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" },
  inter: { label: "Inter", body: "'Inter', sans-serif", googleFontsUrl: "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" },
  grotesk: { label: "Space Grotesk", body: "'Space Grotesk', sans-serif", googleFontsUrl: "https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" },
};

function useFonts(lettertype: LettertypeKey | undefined): void {
  useEffect(() => {
    const preset = FONT_PRESETS[lettertype ?? "plex"] || FONT_PRESETS.plex;
    let link = document.getElementById(FONT_LINK_ID) as HTMLLinkElement | null;
    if (!link) {
      link = document.createElement("link");
      link.id = FONT_LINK_ID;
      link.rel = "stylesheet";
      document.head.appendChild(link);
    }
    link.href = preset.googleFontsUrl;
  }, [lettertype]);
}

/* Fase 10 — de fuzzy-zoekfunctie (Levenshtein over de hele catalogus) draaide voorheen op
   iedere toetsaanslag, zonder vertraging. Bij een grotere catalogus of een tragere telefoon
   is dat een reëel, merkbaar prestatierisico. 150ms is kort genoeg om niet traag te aanvoelen,
   lang genoeg om de meeste toetsaanslagen tijdens het typen te overslaan. */
function useDebouncedValue<T>(waarde: T, vertragingMs: number = 150): T {
  const [debounced, setDebounced] = useState<T>(waarde);
  useEffect(() => {
    const timer = setTimeout(() => setDebounced(waarde), vertragingMs);
    return () => clearTimeout(timer);
  }, [waarde, vertragingMs]);
  return debounced;
}

const uid = (): string => crypto.randomUUID();
const today = (): string => new Date().toISOString().slice(0, 10);
const money = (n: number | undefined): string =>
  (Number.isFinite(n) ? (n as number) : 0).toLocaleString("nl-NL", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDate = (d: string | undefined): string =>
  d ? new Date(d).toLocaleDateString("nl-NL", { day: "2-digit", month: "2-digit", year: "numeric" }) : "";

/* ---------- storage helpers ----------
   Belangrijk: window.storage bestaat alleen binnen een Claude-artifact-omgeving, niet in een
   echte browser/Netlify-deployment. localStorage is de universele, altijd-beschikbare vervanging
   met exact hetzelfde synchrone karakter — hier achter dezelfde async-signatuur gehouden zodat
   geen enkele aanroeper elders in dit bestand hoeft te veranderen. */
interface LoadResult<T> {
  value: T;
  /** true = de key bestond wél maar JSON.parse faalde (corrupte data) — een écht probleem,
      te onderscheiden van "nog geen data" (dat geeft corrupted:false terug). */
  corrupted: boolean;
}
async function loadKey<T>(key: string, fallback: T): Promise<LoadResult<T>> {
  let raw: string | null;
  try {
    raw = localStorage.getItem(key);
  } catch {
    return { value: fallback, corrupted: false };
  }
  if (raw === null) return { value: fallback, corrupted: false };
  try {
    return { value: JSON.parse(raw) as T, corrupted: false };
  } catch {
    return { value: fallback, corrupted: true };
  }
}
async function saveKey<T>(key: string, value: T): Promise<boolean> {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

/* ============================================================
   BUSINESS LOGIC
   ============================================================ */
/* Fase 10 — geklemd op de enige plek die telt: waar dan ook een negatief aantal/prijs of een
   korting buiten [0,100] vandaan komt (getypt, import), het bedrag kan hierna nooit
   meer negatief of onzinnig worden. HTML min/max op de invoervelden is aanvullend, niet de garantie. */
function regelBedrag(r: OfferteRegel): number {
  const aantal = Math.max(0, Number(r.aantal) || 0);
  const prijs = Math.max(0, Number(r.prijs) || 0);
  const kortingPct = Math.min(100, Math.max(0, Number(r.korting) || 0));
  const bruto = aantal * prijs;
  const korting = bruto * (kortingPct / 100);
  return bruto - korting;
}
export interface OfferteTotalen {
  subtotaal: number;
  btw: number;
  totaal: number;
}
function offerteTotalen(regels: OfferteRegel[]): OfferteTotalen {
  const subtotaal = regels.reduce((s, r) => s + regelBedrag(r), 0);
  const btw = regels.reduce((s, r) => s + regelBedrag(r) * ((Number(r.btw) || 0) / 100), 0);
  return { subtotaal, btw, totaal: subtotaal + btw };
}
function nextOfferteNummer(offertes: Offerte[], prefix: string = "OFF-"): string {
  const jaar = new Date().getFullYear();
  const volledigePrefix = `${prefix}${jaar}-`;
  const hoogsteVolgnummer = offertes.reduce((hoogste, o) => {
    if (!o.nummer?.startsWith(volledigePrefix)) return hoogste;
    const volgnummer = parseInt(o.nummer.slice(volledigePrefix.length), 10);
    return Number.isFinite(volgnummer) && volgnummer > hoogste ? volgnummer : hoogste;
  }, 0);
  return `${volledigePrefix}${String(hoogsteVolgnummer + 1).padStart(3, "0")}`;
}

/* ============================================================
   TEMPLATE-ENGINE — Fase 9
   Strikte scheiding: sjabloonWaarden() is pure data (geen React, los testbaar),
   renderSjabloon() is de generieke rendering-laag die ELK template-string + waarden-object
   kan omzetten naar output — herbruikbaar voor andere documenttypes (factuur, werkbon) door
   simpelweg een andere template-string en een ander waarden-object mee te geven.
   ============================================================ */
const BLOK_PLACEHOLDERS = new Set(["Logo", "OfferteRegels", "Totalen", "HandtekeningKlant", "HandtekeningOndertekenaar"]);

export interface SjabloonWaarden {
  Bedrijfsnaam: string;
  Adres: string;
  PostcodePlaats: string;
  Email: string;
  Telefoon: string;
  Website: string;
  KVK: string;
  BTWNummer: string;
  IBAN: string;
  KlantNaam: string;
  KlantAdres: string;
  KlantPostcodePlaats: string;
  OfferteNummer: string;
  Datum: string;
  GeldigTot: string;
  Subtotaal: string;
  BTW: string;
  Totaal: string;
  Notities: string;
  AlgemeneVoorwaarden: string;
  Voettekst: string;
  Footer: string;
  [key: string]: string; // toegang via willekeurige placeholder-naam bij het renderen
}

/* Pure data-laag: bouwt de platte-tekst-waarden voor alle {{Placeholder}}-tokens.
   Geen JSX, geen DOM — dit is losstaand testbaar met alleen objecten als input. */
function sjabloonWaarden(
  offerte: Offerte | undefined,
  klant: Klant | undefined,
  sjabloon: Sjabloon | undefined,
  totalen: OfferteTotalen | undefined
): SjabloonWaarden {
  return {
    Bedrijfsnaam: sjabloon?.bedrijfsnaam || "",
    Adres: sjabloon?.adres || "",
    PostcodePlaats: [sjabloon?.postcode, sjabloon?.plaats].filter(Boolean).join(" "),
    Email: sjabloon?.email || "",
    Telefoon: sjabloon?.telefoon || "",
    Website: sjabloon?.website || "",
    KVK: sjabloon?.kvk || "",
    BTWNummer: sjabloon?.btwNummer || "",
    IBAN: sjabloon?.iban || "",
    KlantNaam: klant?.naam || "",
    KlantAdres: klant?.adres || "",
    KlantPostcodePlaats: [klant?.postcode, klant?.plaats].filter(Boolean).join(" "),
    OfferteNummer: offerte?.nummer || "",
    Datum: fmtDate(offerte?.datum),
    GeldigTot: fmtDate(vervaldatum(offerte?.datum, sjabloon?.geldigheidsdagen)),
    Subtotaal: `€ ${money(totalen?.subtotaal)}`,
    BTW: `€ ${money(totalen?.btw)}`,
    Totaal: `€ ${money(totalen?.totaal)}`,
    Notities: offerte?.notities || "",
    AlgemeneVoorwaarden: sjabloon?.algemeneVoorwaarden || "",
    Voettekst: sjabloon?.voettekst || "",
    Footer: [
      sjabloon?.kvk && `KVK ${sjabloon.kvk}`,
      sjabloon?.btwNummer && `Btw ${sjabloon.btwNummer}`,
      sjabloon?.iban && `IBAN ${sjabloon.iban}`,
    ].filter(Boolean).join("   ·   "),
  };
}

/* Rendering-laag: tekstregel met inline {{token}}-vervanging. Puur tekst, nooit HTML —
   een geopende/gesloten tag in een sjabloon-tekst wordt letterlijk als tekst getoond,
   niet geïnterpreteerd. */
function renderSjabloonTekstregel(line: string, waarden: SjabloonWaarden, key: number): React.ReactNode {
  const parts = line.split(/(\{\{[A-Za-z]+\}\})/g).filter((p) => p !== "");
  const inhoud = parts.map((part) => {
    const m = part.match(/^\{\{([A-Za-z]+)\}\}$/);
    return m ? (waarden[m[1]] ?? "") : part;
  });
  if (inhoud.join("").trim() === "") return null; // lege regel na substitutie: overslaan
  return <div key={key} style={s.printSmall}>{inhoud.join("")}</div>;
}

interface RenderSjabloonContext {
  waarden: SjabloonWaarden;
  sjabloon: Sjabloon | undefined;
  offerte: Offerte | undefined;
}

/* Top-level renderer: splitst het template op regels; een regel die uit precies één
   blok-placeholder bestaat rendert een echt component (tabel/afbeelding/handtekening),
   elke andere regel gaat door de tekst-substitutie hierboven. */
function renderSjabloon(templateStr: string, { waarden, sjabloon, offerte }: RenderSjabloonContext): React.ReactNode[] {
  const regels = (templateStr || "").split("\n");
  return regels
    .map((line, i): React.ReactNode => {
      const trimmed = line.trim();
      const blokMatch = trimmed.match(/^\{\{([A-Za-z]+)\}\}$/);
      if (blokMatch && BLOK_PLACEHOLDERS.has(blokMatch[1])) {
        const naam = blokMatch[1];
        if (naam === "Logo") return sjabloon?.logo ? <img key={i} src={sjabloon.logo} alt="" style={s.printLogo} /> : null;
        if (naam === "OfferteRegels") return <OfferteRegelsTabel key={i} regels={offerte?.regels || []} tabelstijl={sjabloon?.tabelstijl} accentKleur={sjabloon?.accentKleur} secundaireKleur={sjabloon?.secundaireKleur} />;
        if (naam === "Totalen") return <TotalenBlok key={i} waarden={waarden} accentKleur={sjabloon?.accentKleur} />;
        if (naam === "HandtekeningKlant") return <HandtekeningBlok key={i} label="Akkoord klant" sub="Naam, datum en handtekening" />;
        if (naam === "HandtekeningOndertekenaar") {
          return (
            <HandtekeningBlok
              key={i}
              label={sjabloon?.ondertekenaar || sjabloon?.bedrijfsnaam || "—"}
              sub={sjabloon?.handtekeningFunctie || ""}
              afbeelding={sjabloon?.handtekeningAfbeelding}
            />
          );
        }
        return null;
      }
      return renderSjabloonTekstregel(line, waarden, i);
    })
    .filter((node): node is React.ReactNode => node !== null);
}
/* Fase 3 — slim zoeken: gedeeltelijke woorden, woordvolgorde-onafhankelijk,
   typefout-tolerant (Levenshtein), en een klein synoniemenwoordenboek beperkt
   tot afkortingen/termen die daadwerkelijk in de catalogus voorkomen (geen
   algemene NLP-synoniem-engine — dat is bewust buiten scope). */
const SYNONYMS: Record<string, string[]> = {
  pvc: ["vinyl"],
  vinyl: ["pvc"],
  marmo: ["marmoleum"],
  marmoleum: ["marmo"],
};
function normalizeText(str: string): string {
  return str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
function tokenize(str: string): string[] {
  return normalizeText(str).split(/[^a-z0-9]+/).filter(Boolean);
}
function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  let prev: number[] = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const curr: number[] = [i];
    for (let j = 1; j <= b.length; j++) {
      curr[j] = a[i - 1] === b[j - 1] ? prev[j - 1] : 1 + Math.min(prev[j - 1], prev[j], curr[j - 1]);
    }
    prev = curr;
  }
  return prev[b.length];
}
function tokenDistance(queryToken: string, word: string): number {
  if (word.includes(queryToken)) return 0; // gedeeltelijk woord / prefix, wint altijd
  const maxDist = queryToken.length <= 4 ? 1 : queryToken.length <= 7 ? 2 : 3;
  const dist = levenshtein(queryToken, word);
  return dist <= maxDist ? dist + 1 : Infinity;
}
function bestTokenScore(queryToken: string, wordTokens: string[]): number {
  let best = Infinity;
  for (const variant of [queryToken, ...(SYNONYMS[queryToken] || [])]) {
    for (const word of wordTokens) {
      const d = tokenDistance(variant, word);
      if (d < best) best = d;
    }
  }
  return best;
}
function zoekProducten(producten: Product[], query: string): Product[] {
  const queryTokens = tokenize(query);
  if (queryTokens.length === 0) return producten;
  const scored: { p: Product; total: number }[] = [];
  for (const p of producten) {
    const wordTokens = tokenize(p.omschrijving);
    let total = 0;
    let matchesAll = true;
    for (const qt of queryTokens) {
      const score = bestTokenScore(qt, wordTokens);
      if (score === Infinity) { matchesAll = false; break; }
      total += score;
    }
    if (matchesAll) scored.push({ p, total });
  }
  scored.sort((a, b) => a.total - b.total);
  return scored.map((entry) => entry.p);
}

/* Import: exacte kolomnaam-match (Omschrijving, Eenheid, Prijs, BTW) importeert direct.
   Wijken de kolomnamen af, dan koppelt de gebruiker ze handmatig (MappingReviewSheet)
   voordat er geïmporteerd wordt — geen AI, geen externe aanroepen. */
const IMPORT_HEADERS: Record<keyof ColumnMapping, string> = { omschrijving: "omschrijving", eenheid: "eenheid", prijs: "prijs", btw: "btw" };
interface MapImportResult {
  items: Product[];
  error: string | null;
}
function mapImportRows(rows: Record<string, unknown>[]): MapImportResult {
  if (!rows.length) return { items: [], error: "Bestand is leeg." };
  const headerKeys: Record<string, string> = Object.keys(rows[0]).reduce((acc: Record<string, string>, k) => {
    acc[k.trim().toLowerCase()] = k;
    return acc;
  }, {});
  const missing = Object.values(IMPORT_HEADERS).filter((h) => !(h in headerKeys));
  if (missing.length) {
    return {
      items: [],
      error: `Kolommen niet exact herkend. Ontbreekt: ${missing.join(", ")}.`,
    };
  }
  const mapping: ColumnMapping = { omschrijving: headerKeys.omschrijving, eenheid: headerKeys.eenheid, prijs: headerKeys.prijs, btw: headerKeys.btw };
  const items = applyMapping(rows, mapping);
  return { items, error: items.length ? null : "Geen geldige rijen gevonden." };
}

/* Past een kolom-mapping (doelveld -> originele kopnaam) toe op ruwe import-rijen. */
function parseNederlandsGetal(raw: unknown, fallback: number): number {
  const schoon = String(raw ?? "").trim().replace(/%/g, "").replace(/\./g, "").replace(",", ".");
  const getal = parseFloat(schoon);
  return Number.isFinite(getal) ? getal : fallback;
}

function applyMapping(rows: Record<string, unknown>[], mapping: ColumnMapping): Product[] {
  return rows
    .map((row): Product => {
      const omschrijving = String(mapping.omschrijving ? row[mapping.omschrijving] ?? "" : "").trim();
      const eenheid = String(mapping.eenheid ? row[mapping.eenheid] ?? "" : "").trim();
      const prijsRaw = mapping.prijs ? row[mapping.prijs] : "0";
      const prijs = parseNederlandsGetal(prijsRaw, 0);
      const btwRaw = mapping.btw ? row[mapping.btw] : "21";
      const btw = parseNederlandsGetal(btwRaw, 21);
      return { id: uid(), omschrijving, eenheid, prijs, btw };
    })
    .filter((p) => p.omschrijving);
}

/* ============================================================
   SHARED UI PRIMITIVES
   ============================================================ */
type IconComponent = React.ComponentType<{ size?: number; color?: string; className?: string }>;

interface SheetProps {
  title: string;
  onClose: () => void;
  children: React.ReactNode;
}
function Sheet({ title, onClose, children }: SheetProps) {
  return (
    <div style={s.sheetOverlay} onClick={onClose}>
      <div style={s.sheet} onClick={(e: React.MouseEvent) => e.stopPropagation()}>
        <div style={s.sheetHeader}>
          <span style={s.sheetTitle}>{title}</span>
          <button style={s.iconBtn} onClick={onClose} aria-label="Sluiten"><X size={20} /></button>
        </div>
        <div style={s.sheetBody}>{children}</div>
      </div>
    </div>
  );
}
interface FieldProps extends React.InputHTMLAttributes<HTMLInputElement> {
  label: string;
}
function Field({ label, ...props }: FieldProps) {
  return (
    <label style={s.field}>
      <span style={s.fieldLabel}>{label}</span>
      <input style={s.input} {...props} />
    </label>
  );
}
interface EmptyStateProps {
  icon: IconComponent;
  title: string;
  hint: string;
}
function EmptyState({ icon: Icon, title, hint }: EmptyStateProps) {
  return (
    <div style={s.empty}>
      <Icon size={28} color="#B8B2A7" />
      <div style={s.emptyTitle}>{title}</div>
      <div style={s.emptyHint}>{hint}</div>
    </div>
  );
}

/* Fase 10 — verwijderen was overal één tik zonder terugweg. Op mobiel, met vieze/gehandschoende
   vingers, is dat een reëel risico op onbedoeld en onherstelbaar dataverlies (geen undo, geen
   prullenbak). Eén herbruikbare bevestigingsstap in plaats van vier losse aanpassingen. */
interface ConfirmDeleteKnopProps {
  label: string;
  onConfirm: () => void;
  style: React.CSSProperties;
  fullWidth?: boolean;
}
function ConfirmDeleteKnop({ label, onConfirm, style, fullWidth }: ConfirmDeleteKnopProps) {
  const [asking, setAsking] = useState<boolean>(false);
  if (asking) {
    return (
      <div style={fullWidth ? s.confirmDeleteRowFull : s.confirmDeleteRow}>
        <span style={s.printSmall}>Zeker weten?</span>
        <button type="button" style={s.ghostBtnSmall} onClick={() => setAsking(false)}>Nee</button>
        <button type="button" style={style} onClick={onConfirm}><Trash2 size={15} /> Ja, verwijderen</button>
      </div>
    );
  }
  return (
    <button type="button" style={style} onClick={() => setAsking(true)}>
      <Trash2 size={16} /> {label}
    </button>
  );
}

/* ============================================================
   KLANTEN
   ============================================================ */
interface KlantRijProps {
  klant: Klant;
  onSelect: (klant: Klant | KlantDraft) => void;
}
const KlantRij = React.memo(function KlantRij({ klant, onSelect }: KlantRijProps) {
  return (
    <button style={s.listCard} onClick={() => onSelect(klant)}>
      <div style={s.cardTitle}>{klant.naam}</div>
      <div style={s.cardSub}>{[klant.plaats, klant.telefoon].filter(Boolean).join(" · ") || "—"}</div>
    </button>
  );
});

/** Een klant in bewerking kan nog geen id hebben (nieuw, nog niet opgeslagen). */
type KlantDraft = Omit<Klant, "id"> & { id: string | null };

interface KlantenViewProps {
  klanten: Klant[];
  onSave: (klant: Klant) => void;
  onDelete: (id: string) => void;
}
function KlantenView({ klanten, onSave, onDelete }: KlantenViewProps) {
  const [editing, setEditing] = useState<KlantDraft | null>(null);
  const empty: Omit<Klant, "id"> = { naam: "", adres: "", postcode: "", plaats: "", email: "", telefoon: "" };

  return (
    <div style={s.view}>
      <div style={s.viewHeader}>
        <span style={s.viewTitle}>Klanten</span>
        <button style={s.primaryBtnSmall} onClick={() => setEditing({ ...empty, id: null })}>
          <Plus size={16} /> Nieuw
        </button>
      </div>
      {klanten.length === 0 ? (
        <EmptyState icon={Users} title="Nog geen klanten" hint="Voeg je eerste klant toe om een offerte te kunnen maken." />
      ) : (
        <div style={s.list}>
          {klanten.map((k) => (
            <KlantRij key={k.id} klant={k} onSelect={setEditing} />
          ))}
        </div>
      )}
      {editing && (
        <Sheet title={editing.id ? "Klant bewerken" : "Nieuwe klant"} onClose={() => setEditing(null)}>
          <KlantForm
            klant={editing}
            onCancel={() => setEditing(null)}
            onDelete={editing.id ? () => { onDelete(editing.id as string); setEditing(null); } : null}
            onSubmit={(data) => { onSave(data); setEditing(null); }}
          />
        </Sheet>
      )}
    </div>
  );
}
interface KlantFormProps {
  klant: KlantDraft;
  onSubmit: (klant: Klant) => void;
  onCancel: () => void;
  onDelete: (() => void) | null;
}
function KlantForm({ klant, onSubmit, onCancel, onDelete }: KlantFormProps) {
  const [data, setData] = useState<KlantDraft>(klant);
  const set = (k: keyof KlantDraft) => (e: React.ChangeEvent<HTMLInputElement>) => setData({ ...data, [k]: e.target.value });
  return (
    <div>
      <Field label="Naam" value={data.naam} onChange={set("naam")} placeholder="Bedrijfs- of persoonsnaam" />
      <Field label="Adres" value={data.adres} onChange={set("adres")} placeholder="Straat en huisnummer" />
      <div style={s.fieldRow}>
        <Field label="Postcode" value={data.postcode} onChange={set("postcode")} placeholder="1234 AB" />
        <Field label="Plaats" value={data.plaats} onChange={set("plaats")} placeholder="Plaats" />
      </div>
      <Field label="E-mail" value={data.email} onChange={set("email")} type="email" placeholder="naam@voorbeeld.nl" />
      <Field label="Telefoon" value={data.telefoon} onChange={set("telefoon")} placeholder="06 12345678" />
      <div style={s.formActions}>
        {onDelete && <ConfirmDeleteKnop label="Verwijderen" onConfirm={onDelete} style={s.dangerBtn} />}
        <div style={{ flex: 1 }} />
        <button style={s.ghostBtn} onClick={onCancel}>Annuleren</button>
        <button
          style={s.primaryBtn}
          disabled={!data.naam.trim()}
          onClick={() => onSubmit({ ...data, id: data.id || uid() })}
        >
          <Check size={16} /> Opslaan
        </button>
      </div>
    </div>
  );
}

interface MappingField {
  key: keyof ColumnMapping;
  label: string;
  required: boolean;
}
const MAPPING_FIELDS: MappingField[] = [
  { key: "omschrijving", label: "Omschrijving", required: true },
  { key: "eenheid", label: "Eenheid", required: false },
  { key: "prijs", label: "Prijs (ex. btw)", required: true },
  { key: "btw", label: "Btw-percentage", required: false },
];

interface MappingReviewSheetProps {
  headers: string[];
  rows: Record<string, unknown>[];
  mapping: Partial<ColumnMapping>;
  onCancel: () => void;
  onConfirm: (mapping: ColumnMapping) => void;
}
function MappingReviewSheet({ headers, rows, mapping, onCancel, onConfirm }: MappingReviewSheetProps) {
  const [map, setMap] = useState<Partial<ColumnMapping>>(mapping || {});
  const valid = !!(map.omschrijving && map.prijs);
  const preview = useMemo(() => applyMapping(rows.slice(0, 3), map as ColumnMapping), [rows, map]);

  return (
    <div>
      <div style={s.notice}><AlertCircle size={15} /> Kolomnamen niet automatisch herkend — koppel ze hieronder handmatig.</div>
      {MAPPING_FIELDS.map((f) => (
        <label key={f.key} style={s.field}>
          <span style={s.fieldLabel}>
            {f.label}{f.required ? " *" : ""}
          </span>
          <select style={s.input} value={map[f.key] || ""} onChange={(e: React.ChangeEvent<HTMLSelectElement>) => setMap({ ...map, [f.key]: e.target.value || null })}>
            <option value="">— geen —</option>
            {headers.map((h) => <option key={h} value={h}>{h}</option>)}
          </select>
        </label>
      ))}
      {preview.length > 0 && (
        <div style={s.searchResults}>
          {preview.map((p, i) => (
            <div key={i} style={s.searchResultItem}>
              <span>{p.omschrijving || "—"}</span>
              <span style={s.mono}>€ {money(p.prijs)} · {p.btw}%</span>
            </div>
          ))}
        </div>
      )}
      <div style={s.formActions}>
        <div style={{ flex: 1 }} />
        <button style={s.ghostBtn} onClick={onCancel}>Annuleren</button>
        <button style={s.primaryBtn} disabled={!valid} onClick={() => onConfirm(map as ColumnMapping)}>
          <Check size={16} /> Importeren
        </button>
      </div>
    </div>
  );
}

/* Startcatalogus — diensten opgegeven door de gebruiker (prijslijst feb. 2026), excl. btw.
   13 regels uit de bron zijn hier bewust weggelaten: posten zonder vaste eenheidsprijs
   ("in overleg" / "bepalen op locatie") en twee volumekortingen (-5%/-7,5% boven 50/75 m²)
   die niet in het product-datamodel passen — die horen thuis als korting% op de offerteregel.
   Btw is niet in de bronlijst vermeld; hier op 21% (standaardtarief) gezet als aanname. */
const SEED_DIENSTEN = [
  { omschrijving: "Laminaat planken zwevend (incl. ondervloer & plakplint rondom kozijnen/pui)", eenheid: "m²", prijs: 10, btw: 21 },
  { omschrijving: "Laminaat visgraat zwevend (incl. ondervloer & plakplint rondom kozijnen/pui)", eenheid: "m²", prijs: 18.5, btw: 21 },
  { omschrijving: "Laminaat tegel zwevend (incl. ondervloer & plakplint rondom kozijnen/pui)", eenheid: "m²", prijs: 13, btw: 21 },
  { omschrijving: "PVC-Click planken zwevend (incl. ondervloer & plakplint)", eenheid: "m²", prijs: 11, btw: 21 },
  { omschrijving: "PVC-Click visgraat zwevend (incl. ondervloer & plakplint)", eenheid: "m²", prijs: 19.5, btw: 21 },
  { omschrijving: "PVC-Click tegels zwevend (incl. ondervloer & plakplint)", eenheid: "m²", prijs: 14, btw: 21 },
  { omschrijving: "PVC Plank Primeren/Egaliseren +/- 3mm en Vol verlijmen", eenheid: "m²", prijs: 18, btw: 21 },
  { omschrijving: "PVC Visgraat/Hongaarse punt/Egaliseren +/- 3mm en Vol verlijmen", eenheid: "m²", prijs: 24.5, btw: 21 },
  { omschrijving: "PVC Tegel Primeren/Egaliseren +/- 3mm en Vol verlijmen", eenheid: "m²", prijs: 21, btw: 21 },
  { omschrijving: "PVC Plak Band/bies", eenheid: "m1", prijs: 11, btw: 21 },
  { omschrijving: "Marmoleum Primeren/Egaliseren 3mm en Vol verlijmen", eenheid: "m²", prijs: 21.5, btw: 21 },
  { omschrijving: "Leggen gang (Marmoleum)", eenheid: "stuk", prijs: 150, btw: 21 },
  { omschrijving: "Leggen in trapkast (Marmoleum)", eenheid: "stuk", prijs: 100, btw: 21 },
  { omschrijving: "Leggen in toilet (Marmoleum)", eenheid: "stuk", prijs: 70, btw: 21 },
  { omschrijving: "Naad tussen verschillende ruimtes maken", eenheid: "stuk", prijs: 57.5, btw: 21 },
  { omschrijving: "Tapijt 400cm gespannen (incl. latten), randen verlijmd of volledig verlijmd", eenheid: "m1", prijs: 17, btw: 21 },
  { omschrijving: "(Textielrug) Vinyl 400cm vast (max 8m²)", eenheid: "m1", prijs: 20, btw: 21 },
  { omschrijving: "Etagetoeslag (ivm meer zaag/snij- + tilwerk)", eenheid: "m²", prijs: 1.75, btw: 21 },
  { omschrijving: "Extra tbv etagetoeslag: Rol tapijt naar etage (niet voor trappen)", eenheid: "rol", prijs: 8.5, btw: 21 },
  { omschrijving: "Extra tbv etagetoeslag: Rol vinyl/marmoleum naar etage (niet voor trappen)", eenheid: "rol", prijs: 20, btw: 21 },
  { omschrijving: "Plankplint plaatsen", eenheid: "m1", prijs: 2.5, btw: 21 },
  { omschrijving: "Staande plint <50ml (inclusief afkitten, exclusief kit)", eenheid: "m1", prijs: 6, btw: 21 },
  { omschrijving: "Staande plint >50ml (inclusief afkitten, exclusief kit)", eenheid: "m1", prijs: 5, btw: 21 },
  { omschrijving: "Alu profiel plaatsen", eenheid: "m1", prijs: 4, btw: 21 },
  { omschrijving: "Mat insnijden", eenheid: "stuk", prijs: 27.5, btw: 21 },
  { omschrijving: "Deur inkorten / Kozijn inzagen", eenheid: "stuk", prijs: 27.5, btw: 21 },
  { omschrijving: "Inkorten keukenplint (tot 3ml / boven 3ml als twee plinten rekenen)", eenheid: "stuk", prijs: 27.5, btw: 21 },
  { omschrijving: "Luik maken (excl. eventueel benodigd materiaal)", eenheid: "stuk", prijs: 70, btw: 21 },
  { omschrijving: "WC leggen, incl. transparant grijs afkitten (excl. kit)", eenheid: "stuk", prijs: 70, btw: 21 },
  { omschrijving: "Om convectorput heen leggen", eenheid: "stuk", prijs: 70, btw: 21 },
  { omschrijving: "Om kookeiland heen werken", eenheid: "stuk", prijs: 230, btw: 21 },
  { omschrijving: "Tegelvloer/anhydrietvloer diamantschuren en primeren ontvetten", eenheid: "m²", prijs: 4, btw: 21 },
  { omschrijving: "Lijmresten/vloer vlak schuren met diamant", eenheid: "m²", prijs: 7.5, btw: 21 },
  { omschrijving: "Afsmeren/dichtsmeren buizenstelsel vloerverwarming primeren", eenheid: "m²", prijs: 6.5, btw: 21 },
  { omschrijving: "Egaliseren (uitgaande van 3mm)", eenheid: "m²", prijs: 6.5, btw: 21 },
  { omschrijving: "Egaliseren (uitgaande van 6mm)", eenheid: "m²", prijs: 9, btw: 21 },
  { omschrijving: "Egaliseren (tussen 6-10 mm)", eenheid: "m²", prijs: 10, btw: 21 },
  { omschrijving: "Egaliseren (tussen 11-15 mm)", eenheid: "m²", prijs: 11, btw: 21 },
  { omschrijving: "Egaliseren (tussen 16-20 mm)", eenheid: "m²", prijs: 12, btw: 21 },
  { omschrijving: "Ondervloer op rol verlijmen of met kleeflaag", eenheid: "m²", prijs: 4.5, btw: 21 },
  { omschrijving: "Ondervloer platen 1-delig", eenheid: "m²", prijs: 2, btw: 21 },
  { omschrijving: "Ondervloer platen 2-delig (bv Fixfloor/Jumpax)", eenheid: "m²", prijs: 9.5, btw: 21 },
  { omschrijving: "Spaanplaat verlijmen/schroeven", eenheid: "m²", prijs: 7.5, btw: 21 },
  { omschrijving: "Trap (dicht) stofferen (inclusief lijm, rubber)", eenheid: "trede", prijs: 30, btw: 21 },
  { omschrijving: "Trap (dicht) stofferen exclusief stootbord (inclusief lijm, rubber)", eenheid: "trede", prijs: 30, btw: 21 },
  { omschrijving: "Trap (dicht) stofferen inclusief MDF stootbord (inclusief lijm, rubber)", eenheid: "trede", prijs: 40, btw: 21 },
  { omschrijving: "Trap (open) rondom stofferen (inclusief lijm, rubber)", eenheid: "trede", prijs: 40, btw: 21 },
  { omschrijving: "Trap (open) halfrond stofferen (inclusief lijm, rubber, profiel en strippen)", eenheid: "trede", prijs: 35, btw: 21 },
  { omschrijving: "Meerprijs Boucle", eenheid: "trede", prijs: 6.5, btw: 21 },
  { omschrijving: "Meerprijs halfronde latten 28mm", eenheid: "trede", prijs: 1.75, btw: 21 },
  { omschrijving: "Traploper (incl. plaatsen traproede)", eenheid: "trede", prijs: 31.5, btw: 21 },
  { omschrijving: "Zijkanten mee (stofferen)", eenheid: "kant", prijs: 70, btw: 21 },
  { omschrijving: "Trap bekleden met PVC/Vinyl textiel/Marmo, exclusief stootbord", eenheid: "trede", prijs: 25, btw: 21 },
  { omschrijving: "Trap bekleden met PVC/Vinyl textiel/Marmo inclusief stootbord", eenheid: "trede", prijs: 32.5, btw: 21 },
  { omschrijving: "Trap bekleden met PVC/Vinyl textiel/Marmo met MDF stootbord", eenheid: "trede", prijs: 35, btw: 21 },
  { omschrijving: "Meerprijs oude trap (niet recht en vlak)", eenheid: "trede", prijs: 6.5, btw: 21 },
  { omschrijving: "Meerprijs open trap met PVC profiel achterzijde", eenheid: "trede", prijs: 11, btw: 21 },
  { omschrijving: "Vloer laten bezorgen door de Woonzaken (Minimale tarief €50,-)", eenheid: "m²", prijs: 1.25, btw: 21 },
  { omschrijving: "Werk wat moet worden uitgevoerd op uurtarief", eenheid: "uur", prijs: 57.5, btw: 21 },
];

/* ============================================================
   PRODUCTEN
   ============================================================ */
interface ProductRijProps {
  product: Product;
  onSelect: (product: Product | ProductDraft) => void;
}
const ProductRij = React.memo(function ProductRij({ product, onSelect }: ProductRijProps) {
  return (
    <button style={s.listCard} onClick={() => onSelect(product)}>
      <div style={s.cardTitle}>{product.omschrijving}</div>
      <div style={s.cardSub}>€ {money(product.prijs)} / {product.eenheid || "stuk"} · {product.btw}% btw</div>
    </button>
  );
});

/** Een product in bewerking: prijs/btw zijn nog string terwijl het formulier getypt wordt. */
interface ProductDraft {
  id: string | null;
  omschrijving: string;
  eenheid: string;
  prijs: string | number;
  btw: string | number;
}
interface ImportState {
  busy: boolean;
  error: string | null;
  count: number;
  note?: string;
}
interface PendingImport {
  rows: Record<string, unknown>[];
  headers: string[];
  mapping: Partial<ColumnMapping>;
}

interface ProductenViewProps {
  producten: Product[];
  onSave: (product: Product) => void;
  onDelete: (id: string) => void;
  onImport: (items: Product[]) => void;
}
function ProductenView({ producten, onSave, onDelete, onImport }: ProductenViewProps) {
  const [editing, setEditing] = useState<ProductDraft | null>(null);
  const [query, setQuery] = useState("");
  const [importState, setImportState] = useState<ImportState | null>(null);
  const [pendingImport, setPendingImport] = useState<PendingImport | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const empty: ProductDraft = { id: null, omschrijving: "", eenheid: "", prijs: "", btw: 21 };
  const debouncedQuery = useDebouncedValue(query);
  const results = useMemo(() => zoekProducten(producten, debouncedQuery), [producten, debouncedQuery]);

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setImportState({ busy: true, error: null, count: 0 });
    try {
      let rows: Record<string, unknown>[];
      if (file.name.toLowerCase().endsWith(".csv")) {
        rows = await new Promise<Record<string, unknown>[]>((resolve, reject) =>
          Papa.parse(file, { header: true, skipEmptyLines: true, complete: (r: { data: Record<string, unknown>[] }) => resolve(r.data), error: reject })
        );
      } else {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      }
      if (!rows.length) { setImportState({ busy: false, error: "Bestand is leeg.", count: 0 }); return; }

      // Exacte kolomnamen herkend? Direct importeren. Zo niet: handmatig koppelen (geen AI meer).
      const fast = mapImportRows(rows);
      if (!fast.error) {
        onImport(fast.items);
        setImportState({ busy: false, error: null, count: fast.items.length });
        return;
      }

      const headers = Object.keys(rows[0]);
      setImportState(null);
      setPendingImport({ rows, headers, mapping: {} });
    } catch {
      setImportState({ busy: false, error: "Bestand kon niet gelezen worden. Controleer het formaat (.xlsx of .csv).", count: 0 });
    }
  };

  const confirmMapping = (map: ColumnMapping) => {
    if (!pendingImport) return;
    const items = applyMapping(pendingImport.rows, map);
    onImport(items);
    setImportState({ busy: false, error: null, count: items.length });
    setPendingImport(null);
  };

  const handleAddDiensten = () => {
    const bestaande = new Set(producten.map((p) => p.omschrijving.trim().toLowerCase()));
    const nieuw: Product[] = SEED_DIENSTEN.filter((p) => !bestaande.has(p.omschrijving.trim().toLowerCase())).map((p) => ({ ...p, id: uid() }));
    if (nieuw.length === 0) {
      setImportState({ busy: false, error: null, count: 0, note: "Deze diensten staan al allemaal in je catalogus." });
      return;
    }
    onImport(nieuw);
    setImportState({ busy: false, error: null, count: nieuw.length });
  };

  return (
    <div style={s.view}>
      <div style={s.viewHeader}>
        <span style={s.viewTitle}>Producten</span>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={s.ghostBtnSmall} onClick={handleAddDiensten}>
            <Plus size={15} /> Diensten
          </button>
          <button style={s.ghostBtnSmall} onClick={() => fileRef.current?.click()}>
            <Upload size={15} /> Import
          </button>
          <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={handleFile} />
          <button style={s.primaryBtnSmall} onClick={() => setEditing({ ...empty, id: null })}>
            <Plus size={16} /> Nieuw
          </button>
        </div>
      </div>

      {importState?.busy && <div style={s.notice}><Loader2 size={15} className="spin" /> Bestand verwerken…</div>}
      {importState?.error && <div style={s.noticeError}><AlertCircle size={15} /> {importState.error}</div>}
      {importState?.note && <div style={s.notice}><AlertCircle size={15} /> {importState.note}</div>}
      {importState && !importState.busy && !importState.error && !importState.note && importState.count > 0 && (
        <div style={s.noticeOk}>
          <Check size={15} /> {importState.count} producten geïmporteerd.
        </div>
      )}

      <div style={s.searchBar}>
        <Search size={16} color="#8A8377" />
        <input style={s.searchInput} placeholder="Zoek product…" value={query} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setQuery(e.target.value)} />
      </div>

      {producten.length === 0 ? (
        <EmptyState icon={Package} title="Nog geen producten" hint="Voeg producten toe of importeer een Excel/CSV-bestand. Wijken de kolomnamen af, dan koppel je ze zelf." />
      ) : (
        <div style={s.list}>
          {results.map((p) => (
            <ProductRij key={p.id} product={p} onSelect={setEditing} />
          ))}
        </div>
      )}

      {editing && (
        <Sheet title={editing.id ? "Product bewerken" : "Nieuw product"} onClose={() => setEditing(null)}>
          <ProductForm
            product={editing}
            onCancel={() => setEditing(null)}
            onDelete={editing.id ? () => { onDelete(editing.id as string); setEditing(null); } : null}
            onSubmit={(data) => { onSave(data); setEditing(null); }}
          />
        </Sheet>
      )}

      {pendingImport && (
        <Sheet title="Kolommen koppelen" onClose={() => setPendingImport(null)}>
          <MappingReviewSheet
            headers={pendingImport.headers}
            rows={pendingImport.rows}
            mapping={pendingImport.mapping}
            onCancel={() => setPendingImport(null)}
            onConfirm={confirmMapping}
          />
        </Sheet>
      )}
    </div>
  );
}
interface ProductFormProps {
  product: ProductDraft;
  onSubmit: (product: Product) => void;
  onCancel: () => void;
  onDelete: (() => void) | null;
}
function ProductForm({ product, onSubmit, onCancel, onDelete }: ProductFormProps) {
  const [data, setData] = useState<ProductDraft>(product);
  const set = (k: keyof ProductDraft) => (e: React.ChangeEvent<HTMLInputElement>) => setData({ ...data, [k]: e.target.value });
  const valid = data.omschrijving.trim() && Number(data.prijs) >= 0;
  return (
    <div>
      <Field label="Omschrijving" value={data.omschrijving} onChange={set("omschrijving")} placeholder="Bijv. Egaliseren vloer" />
      <div style={s.fieldRow}>
        <Field label="Eenheid" value={data.eenheid} onChange={set("eenheid")} placeholder="m², stuk, m1…" />
        <Field label="Prijs (ex. btw)" value={data.prijs} onChange={set("prijs")} type="number" step="0.01" placeholder="0,00" />
      </div>
      <label style={s.field}>
        <span style={s.fieldLabel}>BTW-percentage</span>
        <div style={s.btwRow}>
          {BTW_OPTIONS.map((pct) => (
            <button
              key={pct}
              type="button"
              style={{ ...s.btwOption, ...(Number(data.btw) === pct ? s.btwOptionActive : {}) }}
              onClick={() => setData({ ...data, btw: pct })}
            >
              {pct}%
            </button>
          ))}
        </div>
      </label>
      <div style={s.formActions}>
        {onDelete && <ConfirmDeleteKnop label="Verwijderen" onConfirm={onDelete} style={s.dangerBtn} />}
        <div style={{ flex: 1 }} />
        <button style={s.ghostBtn} onClick={onCancel}>Annuleren</button>
        <button
          style={s.primaryBtn}
          disabled={!valid}
          onClick={() => onSubmit({ ...data, id: data.id || uid(), prijs: Number(data.prijs), btw: Number(data.btw) })}
        >
          <Check size={16} /> Opslaan
        </button>
      </div>
    </div>
  );
}

/* ============================================================
   OFFERTES
   ============================================================ */
interface OffertesViewProps {
  offertes: Offerte[];
  klanten: Klant[];
  producten: Product[];
  sjablonen: Sjabloon[];
  onSave: (offerte: Offerte) => void;
  onDelete: (id: string) => void;
  onSaveKlant: (klant: Klant) => void;
}
function OffertesView({ offertes, klanten, producten, sjablonen, onSave, onDelete, onSaveKlant }: OffertesViewProps) {
  const [open, setOpen] = useState<Offerte | null>(null);
  const standaard = useMemo(() => sjablonen.find((sj) => sj.isStandaard) || sjablonen[0], [sjablonen]);

  if (open) {
    return (
      <OfferteBuilder
        offerte={open}
        klanten={klanten}
        producten={producten}
        sjablonen={sjablonen}
        onClose={() => setOpen(null)}
        onSave={(o) => { onSave(o); setOpen(null); }}
        onDelete={open.id && offertes.some((x) => x.id === open.id) ? () => { onDelete(open.id as string); setOpen(null); } : null}
      />
    );
  }

  const klantNaam = (id: string): string => klanten.find((k) => k.id === id)?.naam || "Onbekende klant";

  /* Testofferte: gebruikt de eerste klant, of maakt een voorbeeldklant aan als die er nog niet is.
     Regels komen uit SEED_DIENSTEN (exacte omschrijving-match, geen substring-gok), zodat de
     testofferte altijd met echte, bestaande catalogusprijzen werkt in plaats van verzonnen data. */
  const handleTestOfferte = () => {
    let klantId = klanten[0]?.id;
    if (!klantId) {
      const nieuweKlant: Klant = {
        id: uid(), naam: "Voorbeeldklant B.V.", adres: "Voorbeeldstraat 1",
        postcode: "1234 AB", plaats: "Voorbeeldstad", email: "info@voorbeeld.nl", telefoon: "06 12345678",
      };
      onSaveKlant(nieuweKlant);
      klantId = nieuweKlant.id;
    }
    const bron = ["PVC-Click planken zwevend (incl. ondervloer & plakplint)", "Egaliseren (uitgaande van 3mm)", "Plankplint plaatsen"]
      .map((naam) => SEED_DIENSTEN.find((p) => p.omschrijving === naam))
      .filter((p): p is Product => Boolean(p));
    const regels: OfferteRegel[] = bron.map((p, i) => ({
      id: uid(), productId: null, omschrijving: p.omschrijving, eenheid: p.eenheid,
      aantal: i === 2 ? 12 : 25, prijs: p.prijs, korting: i === 2 ? 10 : 0, btw: p.btw,
    }));
    const testOfferte: Offerte = {
      id: uid(), nummer: nextOfferteNummer(offertes, standaard?.offertePrefix), klantId, datum: today(),
      sjabloonId: standaard?.id || null,
      notities: "Testofferte — automatisch aangemaakt om de app te proberen.", regels,
    };
    onSave(testOfferte);
    setOpen(testOfferte);
  };

  /* Fase 8 — kopie maken: nieuw id/nummer/datum, regels met verse ids (anders identiek).
     Klant en aantallen kunnen daarna direct aangepast worden in de bouwer, zoals de spec vraagt. */
  const handleDupliceer = (bron: Offerte) => {
    const kopie: Offerte = {
      id: uid(),
      nummer: nextOfferteNummer(offertes, standaard?.offertePrefix),
      klantId: bron.klantId,
      datum: today(),
      sjabloonId: bron.sjabloonId || standaard?.id || null,
      notities: bron.notities,
      regels: bron.regels.map((r) => ({ ...r, id: uid() })),
    };
    onSave(kopie);
    setOpen(kopie);
  };

  return (
    <div style={s.view}>
      <div style={s.viewHeader}>
        <span style={s.viewTitle}>Offertes</span>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={s.ghostBtnSmall} onClick={handleTestOfferte}>
            <FileText size={15} /> Testofferte
          </button>
          <button
            style={s.primaryBtnSmall}
            disabled={klanten.length === 0}
            onClick={() => setOpen({ id: null, nummer: nextOfferteNummer(offertes, standaard?.offertePrefix), klantId: klanten[0]?.id, datum: today(), sjabloonId: standaard?.id || null, notities: "", regels: [] })}
          >
            <Plus size={16} /> Nieuw
          </button>
        </div>
      </div>
      {klanten.length === 0 && (
        <div style={s.notice}><AlertCircle size={15} /> Voeg eerst een klant toe voordat je een offerte maakt.</div>
      )}
      {offertes.length === 0 ? (
        <EmptyState icon={FileText} title="Nog geen offertes" hint="Maak binnen 60 seconden je eerste offerte." />
      ) : (
        <div style={s.list}>
          {offertes.slice().reverse().map((o) => {
            const { totaal } = offerteTotalen(o.regels);
            return (
              <div key={o.id} style={s.listCardRow}>
                <button style={s.listCardMain} onClick={() => setOpen(o)}>
                  <div style={s.cardTitle}>{o.nummer} — {klantNaam(o.klantId)}</div>
                  <div style={s.cardSub}>{fmtDate(o.datum)} · € {money(totaal)} · {o.regels.length} regel{o.regels.length === 1 ? "" : "s"}</div>
                </button>
                <button style={s.iconBtn} onClick={() => handleDupliceer(o)} aria-label="Kopie maken">
                  <Copy size={16} />
                </button>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

interface OfferteBuilderProps {
  offerte: Offerte;
  klanten: Klant[];
  producten: Product[];
  sjablonen: Sjabloon[];
  onClose: () => void;
  onSave: (offerte: Offerte) => void;
  onDelete: (() => void) | null;
}
function OfferteBuilder({ offerte, klanten, producten, sjablonen, onClose, onSave, onDelete }: OfferteBuilderProps) {
  const [data, setData] = useState<Offerte>(offerte);
  const [query, setQuery] = useState("");
  const [showPrint, setShowPrint] = useState(false);
  const debouncedQuery = useDebouncedValue(query);
  const results = useMemo(() => (debouncedQuery.trim() ? zoekProducten(producten, debouncedQuery).slice(0, 8) : []), [producten, debouncedQuery]);
  const totalen = useMemo(() => offerteTotalen(data.regels), [data.regels]);
  const klant = useMemo(() => klanten.find((k) => k.id === data.klantId), [klanten, data.klantId]);
  const standaardSjabloon = useMemo(() => sjablonen.find((sj) => sj.isStandaard) || sjablonen[0], [sjablonen]);
  const sjabloon = useMemo(() => sjablonen.find((sj) => sj.id === data.sjabloonId) || standaardSjabloon, [sjablonen, data.sjabloonId, standaardSjabloon]);

  if (showPrint) {
    return (
      <OffertePrintView
        offerte={data}
        klant={klant}
        sjabloon={sjabloon}
        totalen={totalen}
        onClose={() => setShowPrint(false)}
      />
    );
  }

  const addRegel = (p: Product) => {
    setData({
      ...data,
      regels: [...data.regels, { id: uid(), productId: p.id, omschrijving: p.omschrijving, eenheid: p.eenheid, aantal: 1, prijs: p.prijs, korting: 0, btw: p.btw }],
    });
    setQuery("");
  };
  const addCustomRegel = () => {
    setData({ ...data, regels: [...data.regels, { id: uid(), productId: null, omschrijving: "", eenheid: "stuk", aantal: 1, prijs: 0, korting: 0, btw: sjabloon?.standaardBtw ?? 21 }] });
  };
  const updateRegel = (id: string, patch: Partial<OfferteRegel>) =>
    setData({ ...data, regels: data.regels.map((r) => (r.id === id ? { ...r, ...patch } : r)) });
  const removeRegel = (id: string) => setData({ ...data, regels: data.regels.filter((r) => r.id !== id) });

  return (
    <div style={s.view}>
      <div style={s.viewHeader}>
        <button style={s.iconBtn} onClick={onClose} aria-label="Terug"><ChevronLeft size={22} /></button>
        <span style={s.viewTitle}>{data.nummer}</span>
        <button
          style={s.ghostBtnSmall}
          disabled={!data.klantId || data.regels.length === 0}
          onClick={() => setShowPrint(true)}
        >
          <Printer size={15} /> PDF
        </button>
        <button style={s.primaryBtnSmall} disabled={!data.klantId || data.regels.length === 0} onClick={() => onSave(data)}>
          <Check size={16} /> Opslaan
        </button>
      </div>

      <div style={s.builderBody}>
        <label style={s.field}>
          <span style={s.fieldLabel}>Klant</span>
          <select style={s.input} value={data.klantId} onChange={(e: React.ChangeEvent<HTMLSelectElement>) => setData({ ...data, klantId: e.target.value })}>
            {klanten.map((k) => <option key={k.id} value={k.id}>{k.naam}</option>)}
          </select>
        </label>

        {sjablonen.length > 1 && (
          <label style={s.field}>
            <span style={s.fieldLabel}>Sjabloon</span>
            <select style={s.input} value={sjabloon?.id || ""} onChange={(e: React.ChangeEvent<HTMLSelectElement>) => setData({ ...data, sjabloonId: e.target.value })}>
              {sjablonen.map((sj) => <option key={sj.id} value={sj.id}>{sj.naam}{sj.isStandaard ? " (standaard)" : ""}</option>)}
            </select>
          </label>
        )}

        <div style={s.searchBar}>
          <Search size={16} color="#8A8377" />
          <input style={s.searchInput} placeholder="Zoek en voeg product toe…" value={query} onChange={(e: React.ChangeEvent<HTMLInputElement>) => setQuery(e.target.value)} />
        </div>
        {results.length > 0 && (
          <div style={s.searchResults}>
            {results.map((p) => (
              <button key={p.id} style={s.searchResultItem} onClick={() => addRegel(p)}>
                <span>{p.omschrijving}</span>
                <span style={s.mono}>€ {money(p.prijs)}</span>
              </button>
            ))}
          </div>
        )}
        <button style={s.ghostBtnSmall} onClick={addCustomRegel}><Plus size={15} /> Losse regel toevoegen</button>

        <div style={s.regelList}>
          {data.regels.map((r) => (
            <div key={r.id} style={s.regelCard}>
              <div style={s.regelTop}>
                <input
                  style={s.regelOmschrijving}
                  value={r.omschrijving}
                  onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateRegel(r.id, { omschrijving: e.target.value })}
                  placeholder="Omschrijving"
                />
                <button style={s.iconBtnGhost} onClick={() => removeRegel(r.id)} aria-label="Regel verwijderen"><Trash2 size={16} /></button>
              </div>
              <div style={s.regelGrid}>
                <label style={s.miniField}>
                  <span style={s.miniLabel}>Aantal</span>
                  <input style={s.miniInput} type="number" step="0.01" min="0" value={r.aantal} onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateRegel(r.id, { aantal: e.target.value as unknown as number })} />
                </label>
                <label style={s.miniField}>
                  <span style={s.miniLabel}>Eenheid</span>
                  <input style={s.miniInput} value={r.eenheid} onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateRegel(r.id, { eenheid: e.target.value })} />
                </label>
                <label style={s.miniField}>
                  <span style={s.miniLabel}>Prijs</span>
                  <input style={s.miniInput} type="number" step="0.01" min="0" value={r.prijs} onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateRegel(r.id, { prijs: e.target.value as unknown as number })} />
                </label>
                <label style={s.miniField}>
                  <span style={s.miniLabel}>Korting %</span>
                  <input style={s.miniInput} type="number" step="1" min="0" max="100" value={r.korting} onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateRegel(r.id, { korting: e.target.value as unknown as number })} />
                </label>
                <label style={s.miniField}>
                  <span style={s.miniLabel}>Btw</span>
                  <select style={s.miniInput} value={r.btw} onChange={(e: React.ChangeEvent<HTMLSelectElement>) => updateRegel(r.id, { btw: Number(e.target.value) })}>
                    {BTW_OPTIONS.map((pct) => <option key={pct} value={pct}>{pct}%</option>)}
                  </select>
                </label>
              </div>
              <div style={s.regelBedrag}>€ {money(regelBedrag(r))}</div>
            </div>
          ))}
          {data.regels.length === 0 && (
            <div style={s.emptyInline}>Nog geen regels — zoek een product hierboven of voeg een losse regel toe.</div>
          )}
        </div>

        <label style={s.field}>
          <span style={s.fieldLabel}>Notities</span>
          <textarea
            style={{ ...s.input, minHeight: 70, resize: "vertical" }}
            value={data.notities}
            onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => setData({ ...data, notities: e.target.value })}
            placeholder="Interne of externe opmerkingen…"
          />
        </label>

        <div style={s.totalsCard}>
          <div style={s.totalsRow}><span>Subtotaal</span><span style={s.mono}>€ {money(totalen.subtotaal)}</span></div>
          <div style={s.totalsRow}><span>Btw</span><span style={s.mono}>€ {money(totalen.btw)}</span></div>
          <div style={{ ...s.totalsRow, ...s.totalsRowFinal }}><span>Totaal</span><span style={s.mono}>€ {money(totalen.totaal)}</span></div>
        </div>

        {onDelete && <ConfirmDeleteKnop label="Offerte verwijderen" onConfirm={onDelete} style={s.dangerBtnFull} fullWidth />}
      </div>
    </div>
  );
}

/* ============================================================
   INSTELLINGEN — minimaal subset van Fase 6, alleen wat Fase 5 (PDF) nodig heeft
   ============================================================ */
const PLACEHOLDER_HELP =
  "{{Bedrijfsnaam}}, {{Logo}}, {{Adres}}, {{PostcodePlaats}}, {{Email}}, {{Telefoon}}, {{Website}}, {{KVK}}, " +
  "{{BTWNummer}}, {{IBAN}}, {{KlantNaam}}, {{KlantAdres}}, {{KlantPostcodePlaats}}, {{OfferteNummer}}, {{Datum}}, " +
  "{{GeldigTot}}, {{OfferteRegels}}, {{Totalen}}, {{Subtotaal}}, {{BTW}}, {{Totaal}}, {{Notities}}, " +
  "{{AlgemeneVoorwaarden}}, {{Voettekst}}, {{Footer}}, {{HandtekeningKlant}}, {{HandtekeningOndertekenaar}}";

/* ============================================================
   SJABLONEN — Fase 9 (vervangt de losse Instellingen-tab uit Fase 5/6)
   ============================================================ */
interface SjablonenViewProps {
  sjablonen: Sjabloon[];
  onSave: (sjabloon: Sjabloon) => void;
  onDelete: (id: string) => void;
  onSetStandaard: (id: string) => void;
  onDupliceer: (sjabloon: Sjabloon) => void;
}
function SjablonenView({ sjablonen, onSave, onDelete, onSetStandaard, onDupliceer }: SjablonenViewProps) {
  const [editing, setEditing] = useState<Sjabloon | null>(null);

  if (editing) {
    const bestaatNogInLijst = sjablonen.some((sj) => sj.id === editing.id);
    return (
      <SjabloonEditor
        sjabloon={editing}
        onCancel={() => setEditing(null)}
        onSave={(sj) => { onSave(sj); setEditing(null); }}
        onDelete={bestaatNogInLijst && sjablonen.length > 1 ? () => { onDelete(editing.id); setEditing(null); } : null}
      />
    );
  }

  return (
    <div style={s.view}>
      <div style={s.viewHeader}>
        <span style={s.viewTitle}>Sjablonen</span>
        <button style={s.primaryBtnSmall} onClick={() => setEditing(nieuwSjabloon(`Sjabloon ${sjablonen.length + 1}`))}>
          <Plus size={16} /> Nieuw
        </button>
      </div>
      <div style={s.list}>
        {sjablonen.map((sj) => (
          <div key={sj.id} style={s.listCardRow}>
            <button style={s.listCardMain} onClick={() => setEditing(sj)}>
              <div style={s.cardTitle}>{sj.naam}{sj.isStandaard ? " · standaard" : ""}</div>
              <div style={s.cardSub}>{sj.bedrijfsnaam || "Nog geen bedrijfsnaam ingesteld"}</div>
            </button>
            {!sj.isStandaard && (
              <button style={s.iconBtn} onClick={() => onSetStandaard(sj.id)} aria-label="Als standaard instellen">
                <Star size={16} />
              </button>
            )}
            <button style={s.iconBtn} onClick={() => onDupliceer(sj)} aria-label="Sjabloon dupliceren">
              <Copy size={16} />
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

/** geldigheidsdagen/standaardBtw zijn string tijdens het typen in het formulier, net als
    aantal/prijs/korting bij offerteregels — dezelfde reden: niet bij elke toetsaanslag dwingend
    naar Number() converteren, dat breekt het natuurlijk intypen van een getal. */
type SjabloonDraft = Omit<Sjabloon, "geldigheidsdagen" | "standaardBtw"> & {
  geldigheidsdagen: string | number;
  standaardBtw: string | number;
};

const MARGE_OPTIES: MargeOptie[] = ["compact", "normaal", "ruim"];
const UITLIJNING_OPTIES: [UitlijningOptie, string][] = [["links", "Links"], ["gecentreerd", "Gecentreerd"]];
const TABELSTIJL_OPTIES: [TabelstijlOptie, string][] = [["lijnen", "Met lijnen"], ["minimal", "Minimaal"]];
const HANDTEKENING_POSITIE_OPTIES: [HandtekeningPositie, string][] = [["onder-totalen", "Onder totalen"], ["onder-klant", "Onder klantgegevens"]];

interface SjabloonEditorProps {
  sjabloon: Sjabloon;
  onCancel: () => void;
  onSave: (sjabloon: Sjabloon) => void;
  onDelete: (() => void) | null;
}
function SjabloonEditor({ sjabloon, onCancel, onSave, onDelete }: SjabloonEditorProps) {
  const [data, setData] = useState<SjabloonDraft>(sjabloon);
  const [logoError, setLogoError] = useState<string | null>(null);
  const [handtekeningError, setHandtekeningError] = useState<string | null>(null);
  const logoRef = useRef<HTMLInputElement>(null);
  const handtekeningRef = useRef<HTMLInputElement>(null);
  const set = (k: keyof SjabloonDraft) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => setData({ ...data, [k]: e.target.value });
  const valid = data.naam.trim().length > 0;

  const handleAfbeeldingFile = (
    e: React.ChangeEvent<HTMLInputElement>,
    veld: "logo" | "handtekeningAfbeelding",
    setError: (msg: string | null) => void,
    maxBytes: number
  ) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    if (!file.type.startsWith("image/")) { setError("Kies een afbeeldingsbestand (PNG, JPG of SVG)."); return; }
    if (file.size > maxBytes) { setError(`Afbeelding is te groot (max ~${Math.round(maxBytes / 1000)} KB) — comprimeer 'm of kies een kleinere versie.`); return; }
    const reader = new FileReader();
    reader.onload = () => { setData((d) => ({ ...d, [veld]: reader.result as string })); setError(null); };
    reader.onerror = () => setError("Kon de afbeelding niet lezen.");
    reader.readAsDataURL(file);
  };

  const opslaan = () => onSave({ ...data, geldigheidsdagen: Number(data.geldigheidsdagen) || 30, standaardBtw: Number(data.standaardBtw) || 21 });

  return (
    <div style={s.view}>
      <div style={s.viewHeader}>
        <button style={s.iconBtn} onClick={onCancel} aria-label="Terug"><ChevronLeft size={22} /></button>
        <span style={s.viewTitle}>{data.naam || "Sjabloon"}</span>
        <button style={s.primaryBtnSmall} disabled={!valid} onClick={opslaan}>
          <Check size={16} /> Opslaan
        </button>
      </div>
      <div style={s.builderBody}>
        <Field label="Sjabloonnaam" value={data.naam} onChange={set("naam")} placeholder="Bijv. Standaard, Spoedklus, Zakelijk" />

        <div style={s.sectieLabel}>Bedrijfsgegevens</div>
        <label style={s.field}>
          <span style={s.fieldLabel}>Logo</span>
          {data.logo && <img src={data.logo} alt="Logo" style={s.logoPreview} />}
          <div style={{ display: "flex", gap: 8 }}>
            <button type="button" style={s.ghostBtnSmall} onClick={() => logoRef.current?.click()}>
              <Upload size={15} /> {data.logo ? "Vervangen" : "Uploaden"}
            </button>
            {data.logo && <button type="button" style={s.dangerBtn} onClick={() => setData({ ...data, logo: null })}><Trash2 size={15} /> Verwijderen</button>}
          </div>
          <input ref={logoRef} type="file" accept="image/*" style={{ display: "none" }} onChange={(e: React.ChangeEvent<HTMLInputElement>) => handleAfbeeldingFile(e, "logo", setLogoError, 700_000)} />
          {logoError && <div style={s.noticeError}><AlertCircle size={15} /> {logoError}</div>}
        </label>
        <Field label="Bedrijfsnaam" value={data.bedrijfsnaam} onChange={set("bedrijfsnaam")} placeholder="Hinrichs Totaalservice" />
        <Field label="Adres" value={data.adres} onChange={set("adres")} placeholder="Straat en huisnummer" />
        <div style={s.fieldRow}>
          <Field label="Postcode" value={data.postcode} onChange={set("postcode")} placeholder="1234 AB" />
          <Field label="Plaats" value={data.plaats} onChange={set("plaats")} placeholder="Plaats" />
        </div>
        <div style={s.fieldRow}>
          <Field label="E-mail" value={data.email} onChange={set("email")} type="email" placeholder="naam@bedrijf.nl" />
          <Field label="Telefoon" value={data.telefoon} onChange={set("telefoon")} placeholder="06 12345678" />
        </div>
        <Field label="Website" value={data.website} onChange={set("website")} placeholder="www.bedrijf.nl" />
        <div style={s.fieldRow}>
          <Field label="KVK-nummer" value={data.kvk} onChange={set("kvk")} placeholder="12345678" />
          <Field label="Btw-nummer" value={data.btwNummer} onChange={set("btwNummer")} placeholder="NL000000000B00" />
        </div>
        <Field label="IBAN" value={data.iban} onChange={set("iban")} placeholder="NL00 BANK 0000 0000 00" />

        <div style={s.sectieLabel}>Huisstijl</div>
        <div style={s.fieldRow}>
          <label style={s.field}>
            <span style={s.fieldLabel}>Primaire kleur</span>
            <input type="color" value={data.accentKleur} onChange={set("accentKleur")} style={s.colorInput} />
          </label>
          <label style={s.field}>
            <span style={s.fieldLabel}>Secundaire kleur</span>
            <input type="color" value={data.secundaireKleur} onChange={set("secundaireKleur")} style={s.colorInput} />
          </label>
        </div>
        <label style={s.field}>
          <span style={s.fieldLabel}>Lettertype</span>
          <div style={s.btwRow}>
            {(Object.entries(FONT_PRESETS) as [LettertypeKey, FontPreset][]).map(([key, preset]) => (
              <button key={key} type="button" style={{ ...s.btwOption, ...(data.lettertype === key ? s.btwOptionActive : {}), fontFamily: preset.body }} onClick={() => setData({ ...data, lettertype: key })}>
                {preset.label}
              </button>
            ))}
          </div>
        </label>
        <label style={s.field}>
          <span style={s.fieldLabel}>Marge (PDF)</span>
          <div style={s.btwRow}>
            {MARGE_OPTIES.map((optie) => (
              <button key={optie} type="button" style={{ ...s.btwOption, ...(data.marge === optie ? s.btwOptionActive : {}) }} onClick={() => setData({ ...data, marge: optie })}>{optie}</button>
            ))}
          </div>
        </label>
        <label style={s.field}>
          <span style={s.fieldLabel}>Uitlijning</span>
          <div style={s.btwRow}>
            {UITLIJNING_OPTIES.map(([val, label]) => (
              <button key={val} type="button" style={{ ...s.btwOption, ...(data.uitlijning === val ? s.btwOptionActive : {}) }} onClick={() => setData({ ...data, uitlijning: val })}>{label}</button>
            ))}
          </div>
        </label>
        <label style={s.field}>
          <span style={s.fieldLabel}>Tabelstijl</span>
          <div style={s.btwRow}>
            {TABELSTIJL_OPTIES.map(([val, label]) => (
              <button key={val} type="button" style={{ ...s.btwOption, ...(data.tabelstijl === val ? s.btwOptionActive : {}) }} onClick={() => setData({ ...data, tabelstijl: val })}>{label}</button>
            ))}
          </div>
        </label>

        <div style={s.sectieLabel}>Documentinstellingen</div>
        <div style={s.fieldRow}>
          <Field label="Offertenummer-prefix" value={data.offertePrefix} onChange={set("offertePrefix")} placeholder="OFF-" />
          <Field label="Geldigheidsduur (dagen)" type="number" min="1" value={data.geldigheidsdagen} onChange={set("geldigheidsdagen")} />
        </div>
        <div style={s.fieldRow}>
          <label style={s.field}>
            <span style={s.fieldLabel}>Standaard btw%</span>
            <select style={s.input} value={data.standaardBtw} onChange={set("standaardBtw")}>
              {BTW_OPTIONS.map((pct) => <option key={pct} value={pct}>{pct}%</option>)}
            </select>
          </label>
          <Field label="Valuta" value={data.valuta} onChange={set("valuta")} placeholder="EUR" />
        </div>

        <div style={s.sectieLabel}>Algemene voorwaarden</div>
        <label style={s.field}>
          <textarea style={{ ...s.input, minHeight: 90, resize: "vertical" }} value={data.algemeneVoorwaarden} onChange={set("algemeneVoorwaarden")} placeholder="Tekst die onderaan elke offerte komt…" />
        </label>

        <div style={s.sectieLabel}>Handtekening</div>
        <label style={s.field}>
          <span style={s.fieldLabel}>Afbeelding (optioneel)</span>
          {data.handtekeningAfbeelding && <img src={data.handtekeningAfbeelding} alt="Handtekening" style={s.logoPreview} />}
          <div style={{ display: "flex", gap: 8 }}>
            <button type="button" style={s.ghostBtnSmall} onClick={() => handtekeningRef.current?.click()}>
              <Upload size={15} /> {data.handtekeningAfbeelding ? "Vervangen" : "Uploaden"}
            </button>
            {data.handtekeningAfbeelding && <button type="button" style={s.dangerBtn} onClick={() => setData({ ...data, handtekeningAfbeelding: null })}><Trash2 size={15} /> Verwijderen</button>}
          </div>
          <input ref={handtekeningRef} type="file" accept="image/*" style={{ display: "none" }} onChange={(e: React.ChangeEvent<HTMLInputElement>) => handleAfbeeldingFile(e, "handtekeningAfbeelding", setHandtekeningError, 300_000)} />
          {handtekeningError && <div style={s.noticeError}><AlertCircle size={15} /> {handtekeningError}</div>}
        </label>
        <Field label="Naam ondertekenaar" value={data.ondertekenaar} onChange={set("ondertekenaar")} placeholder="Bijv. eigen naam" />
        <Field label="Functietitel" value={data.handtekeningFunctie} onChange={set("handtekeningFunctie")} placeholder="Bijv. Eigenaar" />
        <label style={s.field}>
          <span style={s.fieldLabel}>Positie op document</span>
          <div style={s.btwRow}>
            {HANDTEKENING_POSITIE_OPTIES.map(([val, label]) => (
              <button key={val} type="button" style={{ ...s.btwOption, ...(data.handtekeningPositie === val ? s.btwOptionActive : {}) }} onClick={() => setData({ ...data, handtekeningPositie: val })}>{label}</button>
            ))}
          </div>
          <div style={s.printSmall}>Bepaalt waar de handtekeningblokken staan als je hieronder de standaardlayout opnieuw genereert; in een handmatig aangepast template verplaats je de plek zelf.</div>
        </label>
        <Field label="Voettekst" value={data.voettekst} onChange={set("voettekst")} placeholder="Bijv. een slogan of dankbetuiging" />

        <div style={s.sectieLabel}>PDF-layout (template)</div>
        <div style={s.notice}><AlertCircle size={15} /> Platte tekst, geen hardcoded layout. Beschikbare velden: {PLACEHOLDER_HELP}</div>
        <label style={s.field}>
          <textarea
            style={{ ...s.input, minHeight: 220, resize: "vertical", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}
            value={data.layoutTemplate}
            onChange={set("layoutTemplate")}
          />
        </label>
        <button type="button" style={s.ghostBtnSmall} onClick={() => setData({ ...data, layoutTemplate: DEFAULT_LAYOUT_TEMPLATE })}>
          Terugzetten naar standaardlayout
        </button>

        <div style={s.formActions}>
          {onDelete && <ConfirmDeleteKnop label="Verwijderen" onConfirm={onDelete} style={s.dangerBtn} />}
          <div style={{ flex: 1 }} />
          <button style={s.ghostBtn} onClick={onCancel}>Annuleren</button>
          <button style={s.primaryBtn} disabled={!valid} onClick={opslaan}>
            <Check size={16} /> Opslaan
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   PDF / AFDRUKKEN — Fase 5
   Er is geen PDF-bibliotheek beschikbaar in deze artifact-omgeving (niet in de
   toegestane library-lijst), dus dit gebruikt de browser's eigen afdrukfunctie
   op een A4-opgemaakte weergave — "Opslaan als PDF" in het printdialoog geeft
   het gevraagde eindresultaat, zonder een niet-beschikbare library aan te nemen.
   ============================================================ */
function vervaldatum(datum: string | undefined, geldigheidsdagen: number | undefined): string {
  const d = new Date(datum ?? "");
  if (isNaN(d.getTime())) return "";
  d.setDate(d.getDate() + (Number(geldigheidsdagen) || 30));
  return d.toISOString().slice(0, 10);
}

/* Blok-componenten voor de template-engine. Elk is puur data-in/JSX-uit — herbruikbaar
   voor toekomstige documenttypes zolang je regels/waarden in hetzelfde vorm aanlevert. */
interface OfferteRegelsTabelProps {
  regels: OfferteRegel[];
  tabelstijl: TabelstijlOptie | undefined;
  accentKleur: string | undefined;
  secundaireKleur: string | undefined;
}
function OfferteRegelsTabel({ regels, tabelstijl, accentKleur, secundaireKleur }: OfferteRegelsTabelProps) {
  const metLijnen = tabelstijl !== "minimal";
  const cel: React.CSSProperties = metLijnen ? {} : { borderBottom: "none" };
  return (
    <table style={s.printTable}>
      <thead>
        <tr style={{ background: secundaireKleur || undefined }}>
          <th style={{ ...s.printTh, ...cel, color: accentKleur || s.printTh.color }}>Omschrijving</th>
          <th style={{ ...s.printThRight, ...cel, color: accentKleur || s.printThRight.color }}>Aantal</th>
          <th style={{ ...s.printThRight, ...cel, color: accentKleur || s.printThRight.color }}>Prijs</th>
          <th style={{ ...s.printThRight, ...cel, color: accentKleur || s.printThRight.color }}>Bedrag</th>
        </tr>
      </thead>
      <tbody>
        {regels.map((r) => (
          <tr key={r.id}>
            <td style={{ ...s.printTd, ...cel }}>{r.omschrijving}</td>
            <td style={{ ...s.printTdRight, ...cel }}>{r.aantal} {r.eenheid}</td>
            <td style={{ ...s.printTdRight, ...cel }}>€ {money(r.prijs)}</td>
            <td style={{ ...s.printTdRight, ...cel }}>€ {money(regelBedrag(r))}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}
interface TotalenBlokProps {
  waarden: SjabloonWaarden;
  accentKleur: string | undefined;
}
function TotalenBlok({ waarden, accentKleur }: TotalenBlokProps) {
  return (
    <div style={s.printTotals}>
      <div style={s.printTotalsRow}><span>Subtotaal</span><span>{waarden.Subtotaal}</span></div>
      <div style={s.printTotalsRow}><span>Btw</span><span>{waarden.BTW}</span></div>
      <div style={{ ...s.printTotalsRow, ...s.printTotalsFinal, color: accentKleur || s.printTotalsFinal.color }}>
        <span>Totaal</span><span>{waarden.Totaal}</span>
      </div>
    </div>
  );
}
interface HandtekeningBlokProps {
  label: string;
  sub?: string;
  afbeelding?: string | null;
}
function HandtekeningBlok({ label, sub, afbeelding }: HandtekeningBlokProps) {
  return (
    <div style={s.printHandtekeningBlok}>
      <div style={s.printLabel}>{label}</div>
      {afbeelding ? <img src={afbeelding} alt="" style={s.printLogo} /> : <div style={s.printHandtekeningLijn} />}
      {sub && <div style={s.printSmall}>{sub}</div>}
    </div>
  );
}

const MARGE_MM: Record<MargeOptie, string> = { compact: "10mm", normaal: "14mm", ruim: "20mm" };

interface OffertePrintViewProps {
  offerte: Offerte;
  klant: Klant | undefined;
  sjabloon: Sjabloon | undefined;
  totalen: OfferteTotalen;
  onClose: () => void;
}
function OffertePrintView({ offerte, klant, sjabloon, totalen, onClose }: OffertePrintViewProps) {
  const printAreaRef = useRef<HTMLDivElement>(null);
  const [downloadError, setDownloadError] = useState<string | null>(null);
  const waarden = useMemo(() => sjabloonWaarden(offerte, klant, sjabloon, totalen), [offerte, klant, sjabloon, totalen]);
  const inhoud = useMemo(
    () => renderSjabloon(sjabloon?.layoutTemplate || DEFAULT_LAYOUT_TEMPLATE, { waarden, sjabloon, offerte }),
    [sjabloon, waarden, offerte]
  );
  const printAreaStyle: React.CSSProperties = {
    ...s.printArea,
    padding: MARGE_MM[sjabloon?.marge ?? "normaal"] || MARGE_MM.normaal,
    textAlign: sjabloon?.uitlijning === "gecentreerd" ? "center" : "left",
  };

  /* Fallback naast window.print(): leest de al-gerenderde DOM van print-area rechtstreeks uit
     (React's inline styles staan daar al helemaal in) en downloadt 'm als losstaand HTML-bestand.
     Dat bestand kan daarna in ELKE browser geopend worden — buiten de artifact-sandbox om — en
     die browser heeft altijd zijn eigen "opslaan als PDF". Geen aparte HTML-generator nodig,
     dus geen duplicatie van de layout-logica die al in renderSjabloon zit. */
  const handleDownload = () => {
    try {
      const inhoudHtml = printAreaRef.current?.outerHTML;
      if (!inhoudHtml) { setDownloadError("Kon de offerte-inhoud niet lezen."); return; }
      const document_ = `<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="utf-8" />
<title>Offerte ${offerte?.nummer || ""}</title>
<style>
  body { margin: 0; font-family: sans-serif; }
  @page { size: A4; margin: 15mm; }
</style>
</head>
<body>${inhoudHtml}</body>
</html>`;
      const blob = new Blob([document_], { type: "text/html" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `offerte-${offerte?.nummer || "concept"}.html`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setDownloadError(null);
    } catch (err) {
      setDownloadError(`Downloaden is niet gelukt (${err instanceof Error ? err.message : String(err)}).`);
    }
  };

  return (
    <div style={s.printOverlay}>
      <style>{`
        @media print {
          .no-print { display: none !important; }
          body * { visibility: hidden; }
          .print-area, .print-area * { visibility: visible; }
          .print-area { position: absolute; top: 0; left: 0; width: 100%; }
          @page { size: A4; margin: 15mm; }
        }
      `}</style>
      <div className="no-print" style={s.printBar}>
        <button style={s.ghostBtn} onClick={onClose}>Sluiten</button>
        <button style={s.ghostBtn} onClick={handleDownload}><Download size={16} /> Download</button>
        <button style={s.primaryBtn} onClick={() => window.print()}><Printer size={16} /> Afdrukken / PDF</button>
      </div>
      {downloadError && <div className="no-print" style={s.noticeError}><AlertCircle size={15} /> {downloadError}</div>}
      <div className="print-area" style={printAreaStyle} ref={printAreaRef}>
        {inhoud}
      </div>
    </div>
  );
}

/* ============================================================
   APP SHELL
   ============================================================ */
type TabKey = "klanten" | "producten" | "offertes" | "sjablonen";
interface TabDef {
  key: TabKey;
  label: string;
  icon: IconComponent;
}
const TABS: TabDef[] = [
  { key: "klanten", label: "Klanten", icon: Users },
  { key: "producten", label: "Producten", icon: Package },
  { key: "offertes", label: "Offertes", icon: FileText },
  { key: "sjablonen", label: "Sjablonen", icon: Building2 },
];

/* Fase 10 — een onverwachte crash ergens in de boom liet voorheen de hele app blanco/kapot
   achter (we hebben zelf twee zulke crashes gevonden tijdens dit traject: een undefined-clobber
   in dupliceerSjabloon en een Invalid-Date-crash in vervaldatum). Een error boundary is de enige
   plek in React waar een class component nog verplicht is — er bestaat geen hook-equivalent van
   componentDidCatch/getDerivedStateFromError. Dat is dus een bewuste, onvermijdelijke uitzondering
   op de rest van dit bestand, niet een stijlbreuk. */
interface ErrorBoundaryProps {
  children: React.ReactNode;
}
interface ErrorBoundaryState {
  error: Error | null;
}
class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error };
  }
  render() {
    if (this.state.error) {
      return (
        <div style={s.appLoading}>
          <AlertCircle size={28} color="#B3402A" />
          <span style={{ textAlign: "center", padding: "0 24px" }}>
            Er ging iets mis bij het weergeven van de app. Je opgeslagen gegevens zijn hierdoor niet aangetast.
          </span>
          <button style={s.ghostBtn} onClick={() => this.setState({ error: null })}>Opnieuw proberen</button>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function App() {
  return (
    <ErrorBoundary>
      <AppInner />
    </ErrorBoundary>
  );
}

function AppInner() {
  const [tab, setTab] = useState<TabKey>("offertes");
  const [loading, setLoading] = useState(true);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [klanten, setKlanten] = useState<Klant[]>([]);
  const [producten, setProducten] = useState<Product[]>([]);
  const [offertes, setOffertes] = useState<Offerte[]>([]);
  const [sjablonen, setSjablonen] = useState<Sjabloon[]>([]);
  const standaardSjabloon = useMemo(() => sjablonen.find((sj) => sj.isStandaard) || sjablonen[0], [sjablonen]);
  const accent = standaardSjabloon?.accentKleur || DEFAULT_SJABLOON_VELDEN.accentKleur;
  const fontBody = (FONT_PRESETS[standaardSjabloon?.lettertype ?? "plex"] || FONT_PRESETS.plex).body;
  useFonts(standaardSjabloon?.lettertype);

  useEffect(() => {
    (async () => {
      const [k, p, o, sj] = await Promise.all([
        loadKey<Klant[]>(KEYS.klanten, []),
        loadKey<Product[]>(KEYS.producten, []),
        loadKey<Offerte[]>(KEYS.offertes, []),
        loadKey<Sjabloon[] | null>(KEYS.sjablonen, null),
      ]);
      const corrupted = [k, p, o, sj].some((r) => r.corrupted);
      setKlanten(k.value); setProducten(p.value); setOffertes(o.value);
      if (sj.value && sj.value.length > 0) {
        setSjablonen(sj.value);
      } else {
        // Migratie: bestaande Fase 5/6 "instellingen" (enkelvoud) wordt het eerste, standaard sjabloon.
        // Veldnamen zijn bewust ongewijzigd gebleven t.o.v. Fase 5/6, dus dit is een simpele spread —
        // geen enkele bedrijfsgegeven/logo/instelling die je eerder invulde gaat hierbij verloren.
        const oud = await loadKey<Partial<Sjabloon> | null>(KEYS.instellingen, null);
        setSjablonen([nieuwSjabloon("Standaard", { isStandaard: true, ...(oud.value || {}) })]);
      }
      if (corrupted) {
        setSaveError("Een deel van je opgeslagen gegevens kon niet worden gelezen en is mogelijk beschadigd. Wat wél leesbaar was, is geladen — controleer je klanten, producten, offertes en sjablonen voor je verder werkt.");
      }
      setLoading(false);
    })();
  }, []);

  const persist = useCallback(async <T,>(key: string, value: T) => {
    const ok = await saveKey(key, value);
    if (!ok) setSaveError("Opslaan is niet gelukt. Wijzigingen blijven zichtbaar in deze sessie, maar zijn mogelijk niet bewaard.");
  }, []);

  function upsert<T extends { id: string | null }>(list: T[], item: T): T[] {
    const exists = list.some((x) => x.id === item.id);
    return exists ? list.map((x) => (x.id === item.id ? item : x)) : [...list, item];
  }

  const saveKlant = (k: Klant) => { const next = upsert(klanten, k); setKlanten(next); persist(KEYS.klanten, next); };
  const deleteKlant = (id: string) => { const next = klanten.filter((k) => k.id !== id); setKlanten(next); persist(KEYS.klanten, next); };

  const saveProduct = (p: Product) => { const next = upsert(producten, p); setProducten(next); persist(KEYS.producten, next); };
  const deleteProduct = (id: string) => { const next = producten.filter((p) => p.id !== id); setProducten(next); persist(KEYS.producten, next); };
  const importProducts = (items: Product[]) => { const next = [...producten, ...items]; setProducten(next); persist(KEYS.producten, next); };

  const saveOfferte = (o: Offerte) => { const next = upsert(offertes, o); setOffertes(next); persist(KEYS.offertes, next); };
  const deleteOfferte = (id: string) => { const next = offertes.filter((o) => o.id !== id); setOffertes(next); persist(KEYS.offertes, next); };

  const saveSjabloon = (sj: Sjabloon) => {
    const next = upsert(sjablonen, sj);
    const geschatteGrootte = new Blob([JSON.stringify(next)]).size;
    if (geschatteGrootte > 4_500_000) {
      setSaveError("Je sjablonen (met logo's/handtekeningen) zijn samen te groot om op te slaan (limiet ~5 MB per opslagsleutel). Verklein een afbeelding of verwijder een sjabloon.");
      return;
    }
    setSjablonen(next);
    persist(KEYS.sjablonen, next);
  };
  const deleteSjabloon = (id: string) => {
    const verwijderd = sjablonen.find((sj) => sj.id === id);
    let next = sjablonen.filter((sj) => sj.id !== id);
    if (verwijderd?.isStandaard && next.length > 0) next = next.map((sj, i) => (i === 0 ? { ...sj, isStandaard: true } : sj));
    setSjablonen(next); persist(KEYS.sjablonen, next);
  };
  const setStandaardSjabloon = (id: string) => {
    const next = sjablonen.map((sj) => ({ ...sj, isStandaard: sj.id === id }));
    setSjablonen(next); persist(KEYS.sjablonen, next);
  };
  const dupliceerSjabloon = (bron: Sjabloon) => {
    const { id, naam, isStandaard, ...velden } = bron;
    const kopie = nieuwSjabloon(`${naam} (kopie)`, velden);
    const next = [...sjablonen, kopie];
    setSjablonen(next); persist(KEYS.sjablonen, next);
  };

  const themeStyle = (
    <style>{`
      :root { --accent: ${accent}; --font-body: ${fontBody}; }
      * { box-sizing: border-box; }
      input:focus, select:focus, textarea:focus, button:focus-visible {
        outline: 2px solid var(--accent, #C97A1A); outline-offset: 1px;
      }
      .spin { animation: spin 0.9s linear infinite; }
      @keyframes spin { to { transform: rotate(360deg); } }
      @media (prefers-reduced-motion: reduce) { .spin { animation: none; } }
    `}</style>
  );

  if (loading) {
    return (
      <>
        {themeStyle}
        <div style={s.appLoading}>
          <Loader2 size={22} className="spin" />
          <span>Gegevens laden…</span>
        </div>
      </>
    );
  }

  return (
    <div style={s.app}>
      {themeStyle}

      <div style={s.topBar}>
        {standaardSjabloon?.logo ? <img src={standaardSjabloon.logo} alt="" style={s.topBarLogo} /> : <Ruler size={18} color={accent} />}
        <span style={s.topBarTitle}>Offerte</span>
      </div>

      {saveError && <div style={s.noticeError}><AlertCircle size={15} /> {saveError}</div>}

      <div style={s.content}>
        {tab === "klanten" && <KlantenView klanten={klanten} onSave={saveKlant} onDelete={deleteKlant} />}
        {tab === "producten" && <ProductenView producten={producten} onSave={saveProduct} onDelete={deleteProduct} onImport={importProducts} />}
        {tab === "offertes" && <OffertesView offertes={offertes} klanten={klanten} producten={producten} sjablonen={sjablonen} onSave={saveOfferte} onDelete={deleteOfferte} onSaveKlant={saveKlant} />}
        {tab === "sjablonen" && <SjablonenView sjablonen={sjablonen} onSave={saveSjabloon} onDelete={deleteSjabloon} onSetStandaard={setStandaardSjabloon} onDupliceer={dupliceerSjabloon} />}
      </div>

      <nav style={s.bottomNav}>
        {TABS.map(({ key, label, icon: Icon }) => (
          <button key={key} style={{ ...s.navItem, ...(tab === key ? s.navItemActive : {}) }} onClick={() => setTab(key)}>
            <Icon size={20} color={tab === key ? accent : "#8A8377"} />
            <span style={{ color: tab === key ? accent : "#8A8377" }}>{label}</span>
          </button>
        ))}
      </nav>
    </div>
  );
}

/* ============================================================
   STYLES — mobiel-first, industrieel/werkplaats palet
   ============================================================ */
const s: Record<string, React.CSSProperties> = {
  app: { display: "flex", flexDirection: "column", height: "100vh", maxWidth: 480, margin: "0 auto", background: "#EDEBE6", fontFamily: "var(--font-body, 'IBM Plex Sans', sans-serif)", color: "#2B2926" },
  appLoading: { height: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, color: "#8A8377", fontFamily: "var(--font-body, 'IBM Plex Sans', sans-serif)" },
  topBar: { display: "flex", alignItems: "center", gap: 8, padding: "16px 16px 12px", borderBottom: "1px solid #DEDACE" },
  topBarTitle: { fontWeight: 700, fontSize: 17, letterSpacing: 0.2 },
  content: { flex: 1, overflowY: "auto", WebkitOverflowScrolling: "touch" },
  view: { display: "flex", flexDirection: "column", minHeight: "100%" },
  viewHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px", position: "sticky", top: 0, background: "#EDEBE6", zIndex: 2, borderBottom: "1px solid #DEDACE" },
  viewTitle: { fontWeight: 700, fontSize: 16, flex: 1 },
  list: { display: "flex", flexDirection: "column", gap: 8, padding: 12 },
  listCard: { textAlign: "left", background: "#FFFFFF", border: "1px solid #DEDACE", borderRadius: 10, padding: "12px 14px", cursor: "pointer" },
  listCardRow: { display: "flex", alignItems: "center", gap: 2, background: "#FFFFFF", border: "1px solid #DEDACE", borderRadius: 10, paddingRight: 8 },
  listCardMain: { flex: 1, textAlign: "left", background: "none", border: "none", cursor: "pointer", padding: "12px 14px" },
  cardTitle: { fontWeight: 600, fontSize: 14.5 },
  cardSub: { fontSize: 12.5, color: "#8A8377", marginTop: 2 },
  empty: { display: "flex", flexDirection: "column", alignItems: "center", textAlign: "center", gap: 6, padding: "48px 32px", color: "#8A8377" },
  emptyTitle: { fontWeight: 600, color: "#4A463E", fontSize: 14.5 },
  emptyHint: { fontSize: 13, lineHeight: 1.4 },
  emptyInline: { fontSize: 13, color: "#8A8377", padding: "16px 4px", textAlign: "center" },
  bottomNav: { display: "flex", borderTop: "1px solid #DEDACE", background: "#FFFFFF" },
  navItem: { flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 3, padding: "9px 0 10px", background: "none", border: "none", fontSize: 11, fontWeight: 500, cursor: "pointer" },
  navItemActive: {},
  primaryBtn: { display: "flex", alignItems: "center", gap: 6, background: "var(--accent, #C97A1A)", color: "#FFF", border: "none", borderRadius: 8, padding: "10px 16px", fontWeight: 600, fontSize: 14, cursor: "pointer" },
  primaryBtnSmall: { display: "flex", alignItems: "center", gap: 5, background: "var(--accent, #C97A1A)", color: "#FFF", border: "none", borderRadius: 8, padding: "7px 12px", fontWeight: 600, fontSize: 13, cursor: "pointer" },
  ghostBtn: { background: "none", border: "1px solid #DEDACE", borderRadius: 8, padding: "10px 16px", fontWeight: 600, fontSize: 14, cursor: "pointer", color: "#4A463E" },
  ghostBtnSmall: { display: "flex", alignItems: "center", gap: 5, background: "none", border: "1px dashed #C7C1B4", borderRadius: 8, padding: "8px 12px", fontWeight: 600, fontSize: 12.5, cursor: "pointer", color: "#5C5747", margin: "0 12px 8px" },
  dangerBtn: { display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", color: "#B3402A", fontWeight: 600, fontSize: 13.5, cursor: "pointer", padding: "10px 4px" },
  dangerBtnFull: { display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#FBEAE5", border: "1px solid #EAC6BA", color: "#B3402A", borderRadius: 8, padding: "11px", fontWeight: 600, fontSize: 13.5, cursor: "pointer", margin: "4px 12px 20px" },
  confirmDeleteRow: { display: "flex", alignItems: "center", gap: 8 },
  confirmDeleteRowFull: { display: "flex", alignItems: "center", justifyContent: "center", gap: 8, margin: "4px 12px 20px" },
  iconBtn: { background: "none", border: "none", padding: 14, cursor: "pointer", color: "#4A463E", display: "flex" },
  iconBtnGhost: { background: "none", border: "none", padding: 14, cursor: "pointer", color: "#B3402A", display: "flex" },
  formActions: { display: "flex", alignItems: "center", gap: 8, marginTop: 8 },
  field: { display: "flex", flexDirection: "column", gap: 5, marginBottom: 12, flex: 1 },
  fieldRow: { display: "flex", gap: 10 },
  fieldLabel: { fontSize: 12, fontWeight: 600, color: "#5C5747" },
  input: { border: "1px solid #DEDACE", borderRadius: 8, padding: "10px 11px", fontSize: 14.5, fontFamily: "var(--font-body, 'IBM Plex Sans', sans-serif)", background: "#FBFAF8", color: "#2B2926" },
  btwRow: { display: "flex", gap: 8 },
  btwOption: { flex: 1, border: "1px solid #DEDACE", background: "#FBFAF8", borderRadius: 8, padding: "9px 0", fontWeight: 600, fontSize: 13.5, cursor: "pointer", color: "#5C5747" },
  btwOptionActive: { background: "var(--accent, #C97A1A)", borderColor: "var(--accent, #C97A1A)", color: "#FFF" },
  sheetOverlay: { position: "fixed", inset: 0, background: "rgba(30,28,24,0.4)", display: "flex", alignItems: "flex-end", zIndex: 10 },
  sheet: { width: "100%", maxWidth: 480, margin: "0 auto", background: "#EDEBE6", borderRadius: "16px 16px 0 0", maxHeight: "88vh", display: "flex", flexDirection: "column" },
  sheetHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px", borderBottom: "1px solid #DEDACE" },
  sheetTitle: { fontWeight: 700, fontSize: 15.5 },
  sheetBody: { padding: 16, overflowY: "auto" },
  searchBar: { display: "flex", alignItems: "center", gap: 8, background: "#FBFAF8", border: "1px solid #DEDACE", borderRadius: 8, padding: "9px 11px", margin: "0 12px 10px" },
  searchInput: { border: "none", background: "none", flex: 1, fontSize: 14, outline: "none", fontFamily: "var(--font-body, 'IBM Plex Sans', sans-serif)" },
  searchResults: { display: "flex", flexDirection: "column", gap: 4, margin: "0 12px 10px", background: "#FFF", border: "1px solid #DEDACE", borderRadius: 8, overflow: "hidden" },
  searchResultItem: { display: "flex", justifyContent: "space-between", padding: "9px 12px", background: "none", border: "none", borderBottom: "1px solid #EFEDE7", fontSize: 13.5, cursor: "pointer", textAlign: "left", color: "#2B2926" },
  notice: { display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: "#5C5747", background: "#F1EEE6", margin: "10px 12px 0", padding: "9px 11px", borderRadius: 8 },
  noticeError: { display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: "#B3402A", background: "#FBEAE5", margin: "10px 12px 0", padding: "9px 11px", borderRadius: 8 },
  noticeOk: { display: "flex", alignItems: "center", gap: 7, fontSize: 13, color: "#2E6B3E", background: "#E7F3E8", margin: "10px 12px 0", padding: "9px 11px", borderRadius: 8 },
  builderBody: { padding: 12, display: "flex", flexDirection: "column" },
  regelList: { display: "flex", flexDirection: "column", gap: 8, margin: "4px 0 14px" },
  regelCard: { background: "#FFFFFF", border: "1px solid #DEDACE", borderRadius: 10, padding: 12 },
  regelTop: { display: "flex", alignItems: "center", gap: 8, marginBottom: 8 },
  regelOmschrijving: { flex: 1, border: "none", borderBottom: "1px solid #EFEDE7", background: "none", fontSize: 14, fontWeight: 600, padding: "2px 0", fontFamily: "var(--font-body, 'IBM Plex Sans', sans-serif)" },
  regelGrid: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8 },
  miniField: { display: "flex", flexDirection: "column", gap: 3 },
  miniLabel: { fontSize: 10.5, color: "#8A8377", fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.3 },
  miniInput: { border: "1px solid #DEDACE", borderRadius: 6, padding: "6px 7px", fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", background: "#FBFAF8" },
  regelBedrag: { textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, fontSize: 13.5, marginTop: 8, color: "#4A463E" },
  totalsCard: { background: "#FFFFFF", border: "1px solid #DEDACE", borderRadius: 10, padding: 14, margin: "6px 0 16px" },
  totalsRow: { display: "flex", justifyContent: "space-between", fontSize: 13.5, padding: "4px 0", color: "#5C5747" },
  totalsRowFinal: { borderTop: "1px solid #EFEDE7", marginTop: 4, paddingTop: 8, fontWeight: 700, fontSize: 15, color: "#2B2926" },
  mono: { fontFamily: "'IBM Plex Mono', monospace" },

  printOverlay: { position: "fixed", inset: 0, background: "#FFFFFF", zIndex: 20, overflowY: "auto" },
  printBar: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 16px", borderBottom: "1px solid #DEDACE", position: "sticky", top: 0, background: "#EDEBE6", zIndex: 1 },
  printArea: { maxWidth: "210mm", margin: "0 auto", padding: "14mm", color: "#1E1C18", fontFamily: "var(--font-body, 'IBM Plex Sans', sans-serif)", fontSize: 13 },
  printHeader: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 24, gap: 16 },
  printBedrijfsnaam: { fontWeight: 700, fontSize: 16, marginBottom: 4 },
  printSmall: { fontSize: 12, color: "#5C5747", lineHeight: 1.5 },
  printMeta: { textAlign: "right" },
  printTitle: { fontWeight: 700, fontSize: 16, marginBottom: 4 },
  printKlant: { marginBottom: 22 },
  printLabel: { fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.4, color: "#8A8377", marginBottom: 4 },
  printTable: { width: "100%", borderCollapse: "collapse", marginBottom: 18 },
  printTh: { textAlign: "left", fontSize: 10.5, textTransform: "uppercase", letterSpacing: 0.3, color: "#8A8377", borderBottom: "1.5px solid #2B2926", padding: "0 6px 8px 0" },
  printThRight: { textAlign: "right", fontSize: 10.5, textTransform: "uppercase", letterSpacing: 0.3, color: "#8A8377", borderBottom: "1.5px solid #2B2926", padding: "0 0 8px 6px" },
  printTd: { padding: "7px 6px 7px 0", borderBottom: "1px solid #EFEDE7", fontSize: 13 },
  printTdRight: { padding: "7px 0 7px 6px", borderBottom: "1px solid #EFEDE7", fontSize: 13, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" },
  printTotals: { marginLeft: "auto", width: "60%", marginBottom: 26 },
  printTotalsRow: { display: "flex", justifyContent: "space-between", padding: "3px 0", fontSize: 13 },
  printTotalsFinal: { borderTop: "1.5px solid #2B2926", marginTop: 4, paddingTop: 7, fontWeight: 700, fontSize: 14.5 },
  printBlok: { marginBottom: 20 },
  printOndertekening: { display: "flex", gap: 40, marginTop: 36, marginBottom: 30 },
  printHandtekeningBlok: { flex: 1 },
  printHandtekeningLijn: { borderBottom: "1px solid #8A8377", height: 40, marginBottom: 4 },
  printFooter: { fontSize: 10.5, color: "#8A8377", borderTop: "1px solid #EFEDE7", paddingTop: 10, textAlign: "center" },
  printVoettekst: { fontSize: 11.5, color: "#5C5747", textAlign: "center", marginBottom: 6, fontStyle: "italic" },
  printLogo: { maxHeight: 40, maxWidth: 160, objectFit: "contain", marginBottom: 8, display: "block" },

  topBarLogo: { height: 22, maxWidth: 120, objectFit: "contain" },
  logoPreview: { maxHeight: 60, maxWidth: 200, objectFit: "contain", background: "#FBFAF8", border: "1px solid #DEDACE", borderRadius: 8, padding: 8, marginBottom: 8, alignSelf: "flex-start" },
  colorInput: { width: 48, height: 32, border: "1px solid #DEDACE", borderRadius: 6, padding: 2, background: "#FBFAF8", cursor: "pointer" },
  sectieLabel: { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.4, color: "#8A8377", margin: "18px 0 6px" },
};
