# Programmering-Eksamen
Mit eksamensprojekt i Programmering på C niveau på HTX i Holstebro 2023.

## Om programmet
Programmet er et simpelt python **CLI**-program, som opretter journaler / rapporter i HTX-fagende: Fysik, Kemi & Teknologi. Journalerne bliver oprettet på baggrund af skabeloner formatteret i .JSON filer, som oprettes af brugeren. Når en skabelon er oprettet i .JSON format, vil programmet konvertere filen til en .docx fil.

## Karakter
Karakter for projektet: _**12**_

## Dokumentationen
Programmets dokumentation er baseret på teorien i undervisningsforløbet op til eksamensprojektet


## Startparametre for programmet i et CLI
Startparameter | Datatype | Påkrævet (Ja/Nej) | Standardværdi |
:---: | :---: | :---: | :---: |
--subject & -s | string | Ja | None |
--help & -h | boolean | Nej | False |
--title & -t | string | Nej | Indsæt titel her |
--front & -t | boolean | Nej | True |
--front-picture & -p | string (url / file-path) | Nej | None |
--out & -o | string | Nej | ./output |
