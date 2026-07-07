# Offerte-app

Mobiel-first offerte-tool voor afbouw-vakmensen. React + TypeScript + Vite. Volledig statisch —
geen backend, geen API-kosten, geen accounts nodig.

## Lokaal draaien

```
npm install
npm run dev
```

## Deployen naar Netlify

1. Zet dit project in een git-repository (GitHub/GitLab/Bitbucket).
2. Nieuwe site op Netlify, koppel de repository.
3. Build command en publish directory staan al in `netlify.toml` — Netlify pikt dat automatisch op.
4. Deploy.

Dat is alles. Geen environment variables, geen API-sleutel, geen serverless functions — een
gratis Netlify-account is voldoende.

## Wat is er veranderd t.o.v. de vorige versie

Alle AI-functionaliteit is verwijderd, op eigen verzoek — geen API-kosten, blijft gratis:

- **AI-kolomherkenning bij import** (Fase 2) is weg. Wijken kolomnamen in een geïmporteerd
  Excel/CSV-bestand af van de verwachte namen, dan koppel je ze nu altijd zelf in een
  eenvoudig scherm (`MappingReviewSheet`) — die handmatige koppeling bestond al als vangnet en
  werkt op zichzelf prima, alleen de automatische AI-suggestie ervoor is eruit.
- **"Beschrijf de klus" AI-voorstel** (Fase 7) is volledig verwijderd — geen vervanging, want
  daar bestaat geen niet-AI-equivalent van.
- De twee Netlify Functions en de bijbehorende `/api/ai/*`-routes zijn weg.
- `ANTHROPIC_API_KEY` is nergens meer nodig.

Alle andere functionaliteit (klanten, producten, fuzzy zoeken, offertes, sjablonen, PDF/print,
handmatige import) is ongewijzigd.

## Wat ik heb geverifieerd, en wat niet

**Wel:**
- `tsc --strict` compileert schoon na het verwijderen van alle AI-code; geen nieuwe fouten
  geïntroduceerd t.o.v. de versie mét AI.
- Elke referentie naar de verwijderde AI-functies/componenten/types is nagelopen en weg —
  geen dode code, geen dangling imports.

**Niet, en dat kán ik hier niet:**
- Een echte `npm install`, `vite build` of Netlify-deploy — geen netwerktoegang in deze sandbox.

De eerste echte `npm install && npm run build` bij jou is het moment waarop dit verder wordt
bevestigd.
