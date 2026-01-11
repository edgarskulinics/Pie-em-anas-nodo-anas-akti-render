# Lejupielādēt EXE instalācijas dati var šeit: https://failiem.lv/f/57k8pfpm3x

# Pieņemšanas-Nodošanas Akta Ģenerators

Šī ir daudzfunkcionāla darbvirsmas lietojumprogramma, kas izstrādāta, izmantojot Python un PySide6, lai vienkāršotu pieņemšanas-nodošanas aktu (un līdzīgu dokumentu) ģenerēšanu. Tā piedāvā plašas pielāgošanas iespējas PDF un DOCX formātos, ietverot detalizētus iestatījumus dokumenta izkārtojumam, saturam un drošībai.

## Satura rādītājs

- [Pieņemšanas-Nodošanas Akta Ģenerators](#pieņemšanas-nodošanas-akta-ģenerators)
  - [Satura rādītājs](#satura-rādītājs)
  - [Funkcijas](#funkcijas)
  - [Priekšnoteikumi](#priekšnoteikumi)
  - [Instalācija un palaišana](#instalācija-un-palaišana)
    - [1. Klonējiet repozitoriju](#1-klonējiet-repozitoriju)
    - [2. Izveidojiet virtuālo vidi (ieteicams)](#2-izveidojiet-virtuālo-vidi-ieteicams)
    - [3. Instalējiet atkarības](#3-instalējiet-atkarības)
    - [4. Poppler instalācija (obligāta PDF priekšskatījumam un drukāšanai)](#4-poppler-instalācija-obligāta-pdf-priekšskatījumam-un-drukāšanai)
    - [5. Palaidiet lietojumprogrammu](#5-palaidiet-lietojumprogrammu)
  - [Lietošana](#lietošana)
    - [Galvenās cilnes](#galvenās-cilnes)
      - [1. Pamata dati](#1-pamata-dati)
      - [2. Puses](#2-puses)
      - [3. Pozīcijas](#3-pozīcijas)
      - [4. Fotogrāfijas](#4-fotogrāfijas)
      - [5. Iestatījumi & Eksports](#5-iestatījumi--eksports)
      - [6. Papildu iestatījumi](#6-papildu-iestatījumi)
      - [7. Šabloni](#7-šabloni)
      - [8. Adrešu grāmata](#8-adrešu-grāmata)
      - [9. Dokumentu vēsture](#9-dokumentu-vēsture)
      - [10. Karte](#10-karte)
    - [Dokumentu ģenerēšana un eksportēšana](#dokumentu-ģenerēšana-un-eksportēšana)
    - [Projekta saglabāšana un ielāde](#projekta-saglabāšana-un-ielāde)
    - [Noklusējuma iestatījumi](#noklusējuma-iestatījumi)
    - [Teksta bloku pārvaldība](#teksta-bloku-pārvaldība)
  - [Failu struktūra un saglabāšanas vietas](#failu-struktūra-un-saglabāšanas-vietas)
  - [Problēmu novēršana](#problēmu-novēršana)
  - [Licence](#licence)
  - [Autors](#autors)

## Funkcijas

*   **Daudzpusīga dokumentu ģenerēšana:** Izveidojiet pieņemšanas-nodošanas aktus PDF un DOCX formātos.
*   **Detalizēta informācijas ievade:**
    *   Akta numurs, datums, vieta, pasūtījuma un līguma numuri.
    *   Izpildes, pieņemšanas un nodošanas datumi.
    *   Strīdu risināšanas kārtība, konfidencialitātes klauzulas, soda naudas procenti, piegādes nosacījumi, apdrošināšana un papildu nosacījumi.
    *   Atsauces dokumenti un akta statuss.
*   **Pušu pārvaldība:** Ievadiet detalizētu informāciju par pieņēmēju un nodevēju (nosaukums, reģ. nr., adrese, kontaktpersona, tālrunis, e-pasts, bankas konts, juridiskais statuss).
*   **Adrešu grāmata:** Saglabājiet un ielādējiet bieži izmantoto pušu datus.
*   **Pozīciju saraksts:** Pievienojiet, rediģējiet un dzēsiet pozīcijas ar aprakstu, daudzumu, vienību, cenu, sērijas numuru, garantiju un piezīmēm. Automātisks summu aprēķins.
*   **Attēlu pievienošana:** Iekļaujiet attēlus ar parakstiem dokumentā.
*   **PDF un DOCX pielāgošana:**
    *   **Vispārīgi:** Logotips, fontu izvēle, parakstu attēli, PVN aprēķins.
    *   **PDF specifiski:** Lapas izmērs un orientācija, malas, fontu izmēri (galvenei, normālam tekstam, mazam tekstam, tabulām), logotipa un parakstu attēlu platums.
    *   **DOCX specifiski:** Attēlu un parakstu attēlu platums.
    *   **Papildu stila un izkārtojuma iestatījumi:**
        *   Galvenes/kājenes teksta krāsas, tabulas galvenes fona un režģa krāsas.
        *   Rindu atstarpes, lapu numuri, ģenerēšanas laika zīmogs.
        *   Valūtas simbola pozīcija, datuma formāts.
        *   Paraksta līnijas garums un biezums, paraksta fonta izmērs un atstarpes.
        *   Dokumenta un sadaļu virsrakstu fonta izmēri un krāsas.
        *   Tabulas apmales stils un biezums, alternatīvo rindu krāsa.
        *   Kopsumma vārdos (ar valodas izvēli), PVN sadalījums.
        *   Titullapa ar pielāgojamu virsrakstu un logotipu.
*   **QR kodu atbalsts:**
    *   **Individuāls QR kods:** Pievienojiet pielāgotu QR kodu ar jebkādiem datiem, izmēru, pozīciju un krāsu.
    *   **Automātisks QR kods:** Automātiski ģenerējiet QR kodu ar akta ID, izmēru, pozīciju un krāsu.
*   **Drošības funkcijas:**
    *   **PDF šifrēšana:** Iespējojiet PDF šifrēšanu ar lietotāja un īpašnieka parolēm.
    *   **Atļauju kontrole:** Iestatiet atļaujas drukāšanai, kopēšanai, modificēšanai un anotēšanai.
    *   **Ūdenszīmes:** Pievienojiet pielāgotu ūdenszīmi (teksts, fonta izmērs, krāsa, rotācija).
    *   **Digitālā paraksta lauks:** Iespējojiet digitālā paraksta lauku PDF dokumentā.
*   **Šablonu sistēma:** Saglabājiet un ielādējiet aktu konfigurācijas kā šablonus, ieskaitot paroles aizsardzību.
*   **Teksta bloku pārvaldība:** Saglabājiet un ielādējiet bieži izmantotus teksta blokus (piemēram, piezīmēm, strīdu risināšanai) no iebūvētas bibliotēkas.
*   **Dokumentu vēsture:** Ātra piekļuve pēdējiem projektiem.
*   **PDF priekšskatījums:** Reāllaika PDF priekšskatījums ar lapu navigāciju un tālummaiņu.
*   **Kartes integrācija:** Izmantojiet interaktīvu karti, lai ģeokodētu adresi un iestatītu "Vieta" lauku.
*   **Automātiska akta numura ģenerēšana:** Konfigurējiet automātisku akta numuru ģenerēšanu.
*   **Noklusējuma iestatījumi:** Saglabājiet un ielādējiet noklusējuma iestatījumus.

## Priekšnoteikumi

Pirms lietojumprogrammas palaišanas pārliecinieties, ka jūsu sistēmā ir instalēts:

*   **Python 3.8+**
*   **Poppler** (nepieciešams PDF priekšskatījumam un drukāšanai, īpaši Windows sistēmās).

## Instalācija un palaišana

### 1. Klonējiet repozitoriju

Atveriet termināli vai komandrindas uzvedni un izpildiet:

```bash
git clone https://github.com/jusu-lietotajvards/akta-generators.git
cd akta-generators
```
### 2. Izveidojiet virtuālo vidi (ieteicams)

Lai izvairītos no konfliktiem ar citām Python paketēm, ieteicams izveidot virtuālo vidi:

```bash
python -m venv venv
```

Aktivizējiet virtuālo vidi:

```bash
.\venv\Scripts\activate
```

### 3. Instalējiet atkarības

```bash
pip install PySide6 reportlab python-docx Pillow pdf2image qrcode PyPDF2 requests
```

### 4. Poppler instalācija (obligāta PDF priekšskatījumam un drukāšanai)

```bash
pdf2image bibliotēka, ko izmanto PDF priekšskatījumam un drukāšanai, ir atkarīga no Poppler utilītprogrammām.
```

Windows:

Lejupielādējiet Poppler bināros failus no oficiālās vietnes vai, biežāk, no xpdf vietnes (meklējiet "Poppler for Windows" vai "xpdf-tools"). Ieteicams izmantot šos bināros failus.
Izpakojiet arhīvu (piemēram, poppler-23.05.0-win64.zip) uz vietu, kurai varat viegli piekļūt (piemēram, C:\poppler).
Lietojumprogrammā, cilnē "Papildu iestatījumi", norādiet ceļu uz Poppler bin direktoriju (piemēram, C:\poppler\Library\bin).


## Lietošana

Lietojumprogrammas saskarne ir sadalīta vairākās cilnēs, lai organizētu dažādas funkcijas.

### Galvenās cilnes

#### 1. Pamata dati
Šeit ievadiet galveno informāciju par aktu:

*  Akta Nr.: Dokumenta numurs. Var ģenerēt automātiski, nospiežot "Ģenerēt Nr.".
*  Datums: Akta sastādīšanas datums.
*  Vieta: Akta sastādīšanas vieta.
*  Pasūtījuma Nr.: Saistītā pasūtījuma numurs.
*  Līguma Nr.: Saistītā līguma numurs.
*  Izpildes termiņš: Darbu izpildes termiņš.
*  Pieņemšanas datums: Akta pieņemšanas datums.
*  Nodošanas datums: Akta nodošanas datums.
*  Strīdu risināšana: Teksts par strīdu risināšanas kārtību.
*  Konfidencialitāte: Atzīmējiet, ja aktā jāiekļauj konfidencialitātes klauzula.
*  Soda nauda (%): Soda naudas procentu likme.
*  Piegādes nosacījumi: Teksts par piegādes nosacījumiem.
*  Apdrošināšana: Atzīmējiet, ja aktā jāiekļauj apdrošināšanas klauzula.
*  Papildu nosacījumi: Jebkādi citi papildu nosacījumi.
*  Atsauces dokumenti: Norādes uz citiem dokumentiem.
*  Akta statuss: Izvēlieties akta statusu (Melnraksts, Apstiprināts utt.).
*  Valūta: Dokumentā izmantotā valūta.
*  Piezīmes: Vispārīgas piezīmes par aktu.
*  Elektroniskais paraksts: Atzīmējiet, ja dokuments tiks parakstīts elektroniski (ignorē fiziskos parakstus).
*  Rādīt elektroniskā paraksta tekstu PDF dokumentā: Atzīmējiet, ja vēlaties, lai PDF dokumentā būtu redzams paziņojums par elektronisko parakstu.

#### 2. Puses
Šajā cilnē ievadiet informāciju par Pieņēmēju un Nodevēju. Katrai pusei varat norādīt:

*  Nosaukumu / Vārdu, Uzvārdu
*  Reģ. Nr. / personas kodu
*  Adresi
*  Kontaktpersonu
*  Tālruni
*  E-pastu
*  Bankas kontu
*  Juridisko statusu
*  Ielādēt no adrešu grāmatas: Ielādē datus no saglabātajiem kontaktiem.
*  Saglabāt adrešu grāmatā: Saglabā pašreizējos datus adrešu grāmatā.

#### 3. Pozīcijas
Šeit pārvaldiet aktu pozīciju sarakstu:

*  Pievienot pozīciju: Pievieno jaunu rindu tabulā.
*  Dzēst izvēlēto: Dzēš atlasīto rindu.
*  Tabulas kolonnas:
    *  Apraksts
    *  Daudzums
    *  Vienība
    *  Cena
    *  Summa (aprēķinās automātiski)
    *  Seriālais Nr. (redzamība konfigurējama "Papildu iestatījumos")
    *  Garantija (redzamība konfigurējama "Papildu iestatījumos")
    *  Piezīmes pozīcijai (redzamība konfigurējama "Papildu iestatījumos")

#### 4. Fotogrāfijas
Pievienojiet attēlus, kas tiks iekļauti dokumentā:

*  Pievienot foto…: Atver failu dialogu, lai izvēlētos attēlus.
*  Uz augšu / Uz leju: Maina attēlu secību.
*  Dzēst: Dzēš atlasīto attēlu.
*  Dubultklikšķis uz attēla: Rediģē attēla parakstu.

#### 5. Iestatījumi & Eksports
Konfigurējiet vispārīgos dokumenta ģenerēšanas iestatījumus un eksportējiet dokumentus:

*  Aprēķināt PVN: Iespējo PVN aprēķinu.
*  PVN (%): PVN likme.
*  Iekļaut parakstu rindas: Atzīmējiet, ja vēlaties fiziskās parakstu rindas (ignorē, ja ieslēgts elektroniskais paraksts).
*  Logotips (neobligāti): Ceļš uz logotipa attēlu.
*  Fonts TTF/OTF (ieteicams latviešu diakritikai): Ceļš uz pielāgotu fonta failu.
*  Paraksts pieņēmējs attēls: Ceļš uz pieņēmēja paraksta attēlu.
*  Paraksts nodevējs attēls: Ceļš uz nodevēja paraksta attēlu.
*  Valoda: (Pašlaik tikai latviešu, angļu nav implementēta).
*  Saglabāt kā noklusējumu: Saglabā pašreizējos iestatījumus kā noklusējuma iestatījumus nākamajām reizēm.
*  Saglabāt kā šablonu: Saglabā pašreizējo konfigurāciju kā šablonu.
*  Ģenerēt PDF…: Ģenerē aktu PDF formātā.
*  Ģenerēt DOCX…: Ģenerē aktu DOCX formātā.
*  Drukāt PDF…: Atver PDF drukas priekšskatījumu.

#### 6. Papildu iestatījumi
Šī cilne piedāvā plašas pielāgošanas iespējas dokumenta izkārtojumam un drošībai:

*  PDF lapas izmērs un orientācija: A4, Letter, Legal, A3, A5; Portrets vai Ainava.
*  PDF malas (mm): Kreisā, labā, augšējā, apakšējā.
*  PDF fonta izmēri: Galvenes, normāla, maza, tabulas.
*  PDF logotipa platums (mm): Logotipa attēla platums PDF dokumentā.
*  PDF paraksta attēla platums un augstums (mm): Parakstu attēlu izmēri PDF dokumentā.
*  DOCX attēlu platums (collas): Attēlu platums DOCX dokumentā.
*  DOCX paraksta attēla platums (collas): Parakstu attēlu platums DOCX dokumentā.
*  Pozīciju tabulas kolonnu platumi: Komatiem atdalīti platumi milimetros.
*  Automātiski ģenerēt akta numuru: Iespējo automātisku akta numura ģenerēšanu.
*  Noklusējuma valūta un vienība: Noklusējuma vērtības jaunām pozīcijām.
*  Noklusējuma PVN likme (%): Noklusējuma PVN likme.
*  Poppler bin direktorijas ceļš (Windows): Obligāti norādīt ceļu uz Poppler bin direktoriju Windows sistēmās.
*  Krāsu iestatījumi: Galvenes/kājenes teksta krāsa, tabulas galvenes fona un režģa krāsa (Hex formātā).
*  Atstarpes: Tabulas rindu atstarpe (mm), rindu atstarpes reizinātājs.
*  Rādīt lapu numurus: Iespējo lapu numerāciju.
*  Rādīt ģenerēšanas laiku: Iespējo dokumenta ģenerēšanas laika zīmogu.
*  Valūtas simbola pozīcija: Pirms vai pēc summas.
*  Datuma formāts: Pielāgojams datuma formāts.
*  Paraksta līnijas garums (mm) un biezums (pt): Pielāgo paraksta līnijas izskatu.

Titullapa:
*  Pievienot titullapu: Iespējo titullapas ģenerēšanu.
*  Titullapas virsraksts: Pielāgojams titullapas virsraksts.
*  Titullapas logo platums (mm): Logotipa platums titullapā.

Individuālais QR kods:
*  Iekļaut individuālu QR kodu: Iespējo pielāgota QR koda iekļaušanu.
*  Individuālā QR koda dati: Teksts, URL vai citi dati QR kodam.
*  Individuālā QR koda izmērs (mm): QR koda izmērs.
*  Individuālā QR koda pozīcija: bottom_right, bottom_left, top_right, top_left, custom.
*  Individuālā QR koda X/Y pozīcija (mm): Pielāgota pozīcija.
*  Individuālā QR koda krāsa (Hex): QR koda krāsa.
*  Automātiskais QR kods (akta ID):
*  Iekļaut automātisku QR kodu (Akta ID): Iespējo QR koda ģenerēšanu no akta numura.
*  Automātiskā QR koda izmērs (mm): QR koda izmērs.
*  Automātiskā QR koda pozīcija: bottom_left, bottom_right, top_right, top_left, custom.
*  Automātiskā QR koda X/Y pozīcija (mm): Pielāgota pozīcija.
*  Automātiskā QR koda krāsa (Hex): QR koda krāsa.

Ūdenszīme:
*  Pievienot ūdenszīmi: Iespējo ūdenszīmes iekļaušanu.
*  Ūdenszīmes teksts: Pielāgojams ūdenszīmes teksts.
*  Ūdenszīmes fonta izmērs, krāsa (Hex), rotācija (grādi): Pielāgo ūdenszīmes izskatu.

PDF šifrēšana:
*  Iespējot PDF šifrēšanu: Iespējo PDF faila aizsardzību ar parolēm.
*  PDF lietotāja parole: Parole, kas nepieciešama, lai atvērtu dokumentu.
*  PDF īpašnieka parole: Parole, kas nepieciešama, lai mainītu dokumenta atļaujas.
*  Atļaujas: Drukāšana, kopēšana, modificēšana, anotēšana.
*  Noklusējuma valsts un pilsēta: Noklusējuma vērtības.
*  Rādīt kontaktinformāciju galvenē: Iespējo kontaktinformācijas rādīšanu dokumenta galvenē.
*  Kontaktu detaļu galvenes fonta izmērs: Fonta izmērs kontaktinformācijai galvenē.
*  Pozīciju attēlu platums (mm) un paraksta fonta izmērs: Attēlu izmēri un parakstu fonta izmēri pie pozīcijām.
*  Rādīt pozīciju piezīmes, sērijas Nr., garantiju tabulā: Kontrolē šo kolonnu redzamību pozīciju tabulā.
*  Tabulas šūnu polsterējums (mm): Atstarpe šūnu iekšpusē.
*  Tabulas galvenes fonta stils: Bold, italic, normal.
*  Tabulas satura izlīdzināšana: Left, center, right.
*  Paraksta fonta izmērs un atstarpe (mm): Pielāgo parakstu izskatu.
*  Dokumenta virsraksta fonta izmērs un krāsa: Pielāgo galvenā virsraksta izskatu.
*  Sadaļas virsraksta fonta izmērs un krāsa: Pielāgo sadaļu virsrakstu izskatu.
*  Paragrāfa rindu atstarpes reizinātājs: Pielāgo rindu atstarpes paragrāfos.
*  Tabulas apmales stils un biezums (pt): Solid, dashed, none.
*  Tabulas alternatīvās rindas krāsa (Hex): Krāsa katrai otrajai tabulas rindai.
*  Rādīt kopsummu vārdos: Iespējo kopsummas attēlošanu vārdos.
*  Kopsummas vārdos valoda: Valoda kopsummai vārdos (lv, en).
*  Noklusējuma PVN aprēķina metode: Exclusive vai inclusive.
*  Rādīt PVN sadalījumu: Iespējo PVN sadalījuma rādīšanu.
*  Iespējot digitālā paraksta lauku (PDF): Pievieno interaktīvu digitālā paraksta lauku PDF dokumentā.
*  Digitālā paraksta lauka nosaukums, izmērs (mm), pozīcija: Pielāgo digitālā paraksta lauka izskatu un novietojumu.
*  Šablonu direktorijs: Norāda direktoriju, kurā tiek saglabāti un ielādēti šabloni.

#### 7. Šabloni
Pārvaldiet saglabātās aktu konfigurācijas:

*  Pieejamie šabloni: Saraksts ar saglabātajiem šabloniem.
*  Šabloni ar paroli tiks atzīmēti ar krāsu un tooltip.
*  Iebūvētais šablons "Pappus dati (piemērs)" ir pieejams vienmēr.
*  Ielādēt izvēlēto šablonu: Ielādē atlasītā šablona datus. Ja šablonam ir parole, tā tiks pieprasīta.
*  Dzēst atlasītos šablonus: Dzēš atlasītos šablonus. Ja šablonam ir parole, tā tiks pieprasīta.
*  Mainīt/Pievienot paroli: Iestata vai maina paroli atlasītajam šablonam.
*  Noņemt paroli: Noņem paroli no atlasītā šablona.

#### 8. Adrešu grāmata
Pārvaldiet saglabātos pušu kontaktus:

*  Saglabātās personas: Saraksts ar saglabātajām personām/uzņēmumiem.
*  Ielādēt izvēlēto: Ielādē atlasītās personas datus kā Pieņēmēju vai Nodevēju.
*  Dzēst izvēlēto: Dzēš atlasīto personu no adrešu grāmatas.

#### 9. Dokumentu vēsture
Piekļūstiet pēdējiem saglabātajiem projektiem:

*  Pēdējie projekti: Saraksts ar pēdējiem 10 saglabātajiem projektiem.
*  Ielādēt izvēlēto projektu: Ielādē atlasītā projekta datus.
*  Notīrīt vēsturi: Dzēš visus ierakstus no vēstures.
*  Atvērt mapi: Atver failu pārlūkā mapi, kurā saglabāts atlasītais projekts.

#### 10. Karte
Izmantojiet interaktīvu karti, lai precizētu atrašanās vietu:

*  Interaktīva karte: Parāda OpenStreetMap karti.
*  Kartes klikšķis: Noklikšķinot uz kartes, tiek iegūtas koordinātes un tās parādītas ievades laukos.
*  Platums (Latitude) / Garums (Longitude): Ievades lauki koordinātēm.
*  Iestatīt marķieri: Novieto marķieri kartē atbilstoši ievadītajām koordinātēm.
*  Iegūt atrašanās vietu: Iegūst kartes centra koordinātes un parāda tās ievades laukos.
*  Iestatīt 'Vieta' lauku: Veic reversās ģeokodēšanas pieprasījumu (izmantojot Nominatim API) ar ievadītajām koordinātēm un aizpilda "Vieta" lauku cilnē "Pamata dati".

### Dokumentu ģenerēšana un eksportēšana
Pēc visu nepieciešamo datu ievades un iestatījumu konfigurēšanas, dodieties uz cilni "Iestatījumi & Eksports".
Nospiediet "Ģenerēt PDF…" vai "Ģenerēt DOCX…", lai saglabātu dokumentu.
Lietojumprogramma automātiski izveidos jaunu mapi AktaGenerators_Output jūsu dokumentu mapē un saglabās tajā dokumentu kopā ar JSON failu, kas satur visus ievadītos datus.
Nospiediet "Drukāt PDF…", lai atvērtu PDF drukas priekšskatījumu un nosūtītu dokumentu uz printeri.

### Projekta saglabāšana un ielāde
Saglabāt projektu…: Izvēlnē Fails -> Saglabāt projektu… ļauj saglabāt visu pašreizējo aktu konfigurāciju (visas cilnes) JSON failā.
Ielādēt projektu…: Izvēlnē Fails -> Ielādēt projektu… ļauj ielādēt iepriekš saglabātu JSON projektu.

### Noklusējuma iestatījumi
Cilnē "Iestatījumi & Eksports" nospiediet "Saglabāt kā noklusējumu", lai saglabātu pašreizējos iestatījumus (izņemot pozīcijas un attēlus) kā noklusējuma iestatījumus. Tie tiks automātiski ielādēti katru reizi, kad palaidīsiet lietojumprogrammu.

### Teksta bloku pārvaldība
Laukiem, kas atbalsta teksta blokus (piemēram, "Piezīmes", "Strīdu risināšana"), blakus ievades laukam ir pogas "Saglabāt bloku" un "Dzēst bloku", kā arī nolaižamais saraksts ar saglabātajiem blokiem.
Saglabāt bloku: Saglabā pašreizējo ievades lauka saturu ar norādītu nosaukumu.
Ielādēt bloku: Ielādē izvēlēto bloku ievades laukā.
Dzēst bloku: Dzēš izvēlēto bloku.

### Failu struktūra un saglabāšanas vietas
Lietojumprogramma izmanto specifiskas direktorijas, lai saglabātu datus un iestatījumus:

APP_DATA_DIR (piemēram, C:\Users\JūsuLietotājs\AppData\Local\AktaGenerators Windows sistēmās vai ~/.local/share/AktaGenerators Linux sistēmās):

*  AktaGenerators/history.json: Saglabā pēdējo projektu vēsturi.
*  AktaGenerators/address_book.json: Saglabā adrešu grāmatas ierakstus.
*  AktaGenerators/default_settings.json: Saglabā noklusējuma iestatījumus.
*  AktaGenerators/text_blocks.json: Saglabā pielāgotos teksta blokus.
*  AktaGenerators_Projects/: Direktorijs saglabātajiem projektu JSON failiem.
*  AktaGenerators_Templates/: Direktorijs saglabātajiem šablonu JSON failiem (ceļš konfigurējams "Papildu iestatījumos").
*  DOCUMENTS_DIR (piemēram, C:\Users\JūsuLietotājs\Documents Windows sistēmās):
