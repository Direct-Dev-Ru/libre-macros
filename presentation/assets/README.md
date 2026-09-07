# Картинки для презентации

Положите сюда файлы с именами из раскадровки (`../STORYBOARD.md`, `../STORYBOARD_short.md`).

## Полная (`full`)

| Имя файла | Слайд |
|-----------|-------|
| `01_title_hero.png` | 01 титул |
| `03_risk_gap.png` | 03 разрыв процесса |
| `04_excel_chaos.png` | 04 хаос Excel |
| `05_pipeline.png` | 05 pipeline |
| `07_architecture.png` | 07 архитектура |
| `09_scenario.png` | 09 сценарий |
| `10_wizard_ui.png` | 10 визард (лучше скриншот) |
| `11_deploy.png` | 11 внедрение |
| `14_closing.png` | 14 финал |

## Краткая (`short`)

| Имя файла | Слайд | Fallback |
|-----------|-------|----------|
| `s01_title.png` | 01 титул | `01_title_hero.png` |
| `s02_vba_wall.png` | 02 стена VBA | `03_risk_gap.png` |
| `s03_pipeline.png` | 03 killer / pipeline | `05_pipeline.png` |
| `s04_practice.png` | 04 на практике | `09_scenario.png` |
| `s05_ask.png` | 05 предложение | `14_closing.png` / `s04_ask.png` |

Допустимы `.png` / `.jpg` / `.jpeg` / `.webp` с тем же basename.

Промпты — в соответствующих STORYBOARD*.md.
Если файла нет, `build_pptx.py` рисует заглушку (для short — пробует fallback).
