# Readme - LaTeX Vorlage

In der Datei `Bachelorarbeit.tex` sind einige Macros am Anfang,
die zum generieren der ersten Paar Seiten (Deckblatt, Erklärung zur Arbeit, etc...)
verwendet werden.

# Formelverzeichniss

Falls ein Formelverzeichniss gewünscht ist, dann können Sie in
`setup/listofequations.tex` eine Möglichkeit sehen, wie es möglich ist
eines Anzulegen, indem man einfach ein neues Macro z.B. \myeqations
definiert.

# Struktur

```
.
├── Bachelorarbeit.tex
├── Bilder <- Bilder ist als Standart Suchverzeichniss
│   │         für Grafiken konfiguriert
│   └── thab_logo.png
├── Kapitel
│   ├── Einleitung.tex
│   └── ZusammenfassungundAusblick.tex
├── Literatur
│   └── Literatur.bib
├── Makefile
├── Seiten
│   ├── Acronyme.tex     <- Abkürzungsverzeichniss
│   ├── Anhang.tex
│   ├── Deckblaetter.tex <- Bitte Ungeändert lassen
│   └── Vorwort.tex
└── setup
    ├── commands.tex
    ├── custom_packages.tex <- Hier eigene Pakte einpflegen, wenn benötigt
    │                          Vereinfacht es, die Vorlagee zu Updaten
    ├── listofequations.tex
    ├── macros.tex
    ├── packages.tex
    └── page.tex
```

# Abhängigeiten

Momentan wird nur pdfLaTeX als LaTeX-Typesetter unterstützt.
Wenn Sie z.B. xetex verwenden wollen, dann suchen sie alle
pdtex Erwähnungen, und ändern sie diese, wenn von den Pakten
unterstützt auf ihren gewollten Typesetter ab.
