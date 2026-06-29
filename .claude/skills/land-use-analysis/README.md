# land-use-analysis (Claude Skill)

Competitive land-use / zoning / multifamily-supply-threat analysis for **any US property**.
See `SKILL.md` for the full workflow. Generic, config-driven scripts + reference playbooks;
the locale-specific crosswalks and the threat ranking are reasoned per run.

## Install
**Claude Code (this machine):** copy this folder into your personal skills dir so it's always
available:
```
cp -r "land-use-analysis" ~/.claude/skills/
```
(or a project's `.claude/skills/`). Then it's invocable as the `land-use-analysis` skill.

**claude.ai (browser):** zip this folder and upload it under Settings → Capabilities → Skills
(Pro/Team/Enterprise). The model reads `SKILL.md` and runs the scripts in the code sandbox;
keep open internet enabled so it can pull the GIS data.

## Deps
`pip install -r requirements.txt`

## Run (summary)
Make a run directory per property, author `<dir>/config.json` (see
`references/config-schema.md` + `templates/config.example.json`), then:
```
python scripts/run_all.py classify --root <dir>
python scripts/run_all.py prep     --root <dir>
#  ... reason -> write <dir>/in/reasoned_ranking.json + <dir>/Tables/decisions_log.md ...
python scripts/run_all.py deliver  --root <dir>
```
Deliverables land in `<dir>/`: the Excel workbook + the interactive HTML viewer.
