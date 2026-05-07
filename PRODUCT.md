# PCC Integrity — CIPS Dashboard

## Product Purpose
Internal operations tool for cathodic protection engineers at Protección Catódica de Colombia. Users inspect pipeline segments on the Ocensa system in Colombia and need to identify which segments fail the NACE criterion (pipe potential < -850 mV Off) and act on that information fast.

## Users
- Field engineers: reviewing post-inspection results, looking for problem segments
- Supervisors: comparing historical vs current runs, prioritizing maintenance
- Both access the tool from office desks on wide monitors (1920px+), not mobile
- Spanish-speaking; technical domain; not impressed by decoration

## Register
product

## Core Tasks (in order of frequency)
1. Identify which pipeline segments are critical (failure rate, score)
2. See where on the pipeline the failures concentrate (PK chart + GPS map)
3. Compare segments against each other and against prior inspections
4. Export or communicate findings to maintenance teams

## Anti-references
- Colorful SaaS marketing dashboards (Mixpanel, Amplitude visual style)
- Heavy glassmorphism/blur cards
- Dark mode that's dark for aesthetics, not function
- Over-animated page-load sequences that make the tool feel slow
- Generic BI templates (Tableau, Power BI default themes)

## Strategic Principles
- Speed of recognition beats speed of interaction. The user must parse the critical situation in under 3 seconds.
- Semantic color is load-bearing. Green/amber/red mean NACE compliance states; nothing else should use these hues.
- Data density is a feature. Engineers expect tables and charts with actual numbers, not summaries.
- No chrome for chrome's sake. Every visual element must earn its place by communicating data or affording an action.

## Semantic Color System (non-negotiable)
- Protegido (criterion met, -850 to -1200 mV): #16A34A green
- Sobreprotegido (over-protected, < -1200 mV): #D97706 amber
- Desprotegido (unprotected, > -850 mV): #D50032 red
- Sin dato: #94A3B8 slate
- These four colors must ONLY appear in data contexts, never as chrome decoration

## Brand
- PCC Integrity / Protección Catódica de Colombia
- Primary red: #D50032 (brand only; also maps to Desprotegido state)
- No blues anywhere
