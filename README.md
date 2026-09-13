# Paliwo PF

Dziennik tankowań i kosztów eksploatacji samochodu (RA-STER). Jeden plik
`index.html` (frontend) + Google Apps Script (`Kod.gs`) jako backend
zapisujący dane w Arkuszu Google + GitHub Pages jako hosting.

Backend wdraża się automatycznie przy każdym pushu na `main` przez
`.github/workflows/deploy-gas.yml` (clasp). Konfiguracja opisana w
`.claude/skills/gas-clasp-autodeploy/SKILL.md`.
