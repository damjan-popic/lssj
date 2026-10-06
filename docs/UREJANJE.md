# Urejanje gradiva

`LSSJ_skripta.md` je glavni besedilni vir. Ob gradnji se iz nje preberejo poglavja, razdelki, vaje, rešitve, dodatki in bibliografija. Polne razlage se vključijo v opombe prosojnic; besedila vaj in rešitev se ne vzdržujejo še v ločenem podatkovnem izvodu.

Datoteke `predavanja/chapter_1.json`–`chapter_8.json` vsebujejo kratke projekcije razlag, primere in navodila predavatelju. Te projekcije so uredniško oblikovane, zato jih program ne povzema samodejno. **Vsebinski popravek skripte lahko zahteva tudi popravek ustreznega JSON-a.** Polja `source_sections` povezujejo prosojnice z razdelki skripte. Polji `short_title` in po potrebi `display_title` določata kratke prikazane naslove poglavij.

`LSSJ_predavanja.pptx` je sestavljeni izdelek. Neposredne popravke v PowerPointu ob naslednji gradnji nadomesti programsko ustvarjena različica; trajne popravke zato prenesite v skripto, projekcije ali gradnik.

## Lokalna gradnja

Potrebujete Python 3.10 ali novejši. Iz korena repozitorija:

```sh
python -m venv .venv
```

Aktivirajte okolje z `source .venv/bin/activate` na Linuxu/macOS ali `.venv\Scripts\Activate.ps1` v PowerShellu. Nato:

```sh
python -m pip install -r requirements.txt
python scripts/build_slides.py
```

Ukaz obnovi `LSSJ_predavanja.pptx` v korenu repozitorija in izpiše povzetek preverjanja. Ne potrebuje povezave v omrežje, naloženih referenčnih knjig ali prejšnjih delovnih map. Izhodne datoteke ne presojajte samo po uspešno zaključenem ukazu: pred objavo odprite tudi nekaj zahtevnejših prosojnic, daljše rešitve in opombe.

### Pisave

Objavljena postavitev uporablja Arial, besedilo za izračun prelomov pa je bilo izmerjeno z združljivo pisavo Nimbus Sans. Za enake odločitve o delitvi prosojnic uporabite Nimbus Sans; na Debianu/Ubuntuju jo vsebuje paket `fonts-urw-base35`.

Gradnik sam poišče Nimbus Sans, Arial ali Liberation Sans prek Fontconfiga, nato še običajne namestitve Ariala na Windowsu in macOS. Pisava ni vključena v repozitorij ali vdelana v PPTX. Če je program ne najde, podajte obe datoteki:

```sh
python scripts/build_slides.py --font-regular pot/do/Regular.otf --font-bold pot/do/Bold.otf
```

Možnost `--font-family "Ime pisave"` spremeni tudi ime pisave v predstavitvi. Druge pisave lahko spremenijo prelome in število prosojnic; preverite izpis opozoril ter dejanski prikaz na računalniku za predavanje.

## Običajni postopek spremembe

1. Popravite `LSSJ_skripta.md`, nato preglejte povezane projekcije v `predavanja/`.
2. Ohranite strukturo naslovov `## 1. Naslov` in `### 1.1 Naslov`. Razdelek vaj ima naslov `Vaje`, `Naloge` ali `Sklepne vaje`, rešitve pa `Rešitve z razlago`. Vaje in rešitve naj imajo ujemajoče se neprekinjeno številčenje od 1.
3. Ponovno zgradite PPTX in preverite povzetek. Pri trenutni vsebini in Nimbus Sans pričakujemo **362 prosojnic, 103 sklope vaj in 103 pripadajoče rešitve**. Število prosojnic se ob vsebinskih spremembah lahko upravičeno spremeni.
4. Preglejte besedilne spremembe ter v isti spremembi shranite prenovljeno skripto, spremenjene projekcije in novo predstavitev.

Gradnik preveri številčenje, ujemanje vaj z rešitvami, sklice projekcij na razdelke in vključitev vseh razlag v opombe. Razlike v vsebinskem pomenu med skripto in kratko projekcijo mora še vedno preveriti urednik.

Za kratek vzorec postavitev uporabite `--sample`. Z `--output pot/datoteka.pptx` določite drug izhod. `--report pot/porocilo.json` po želji shrani tehnični povzetek; takšno začasno poročilo ni del učnega gradiva. Vse možnosti prikaže `python scripts/build_slides.py --help`.
