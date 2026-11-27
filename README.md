# PPT Cavalieri

Generatore di "cavalieri elettronici": una presentazione PowerPoint con una slide per relatore.

Apri `index.html` (anche offline, tutto gira nel browser):

1. **Template** – carica il .pptx/.potx dell'agenzia o il tuo. Nella slide modello scrivi `{{NOME}}` e `{{RUOLO}}` dove devono comparire. Senza segnaposto uso il titolo della slide, altrimenti aggiungo una casella di testo (posizione in %). Senza template: sfondo a tinta unita o immagine.
2. **Relatori** – carica il programma (PDF, Word .docx o testo): l'app ricava i nomi (anche titoli e ruoli) e te li fa controllare prima di inserirli. Oppure scrivi un nome per riga, `Nome Cognome | Ruolo` (o incolla due colonne da Excel).
3. **Testo** – font, dimensione, colore, allineamento, riduzione automatica dei nomi lunghi.
4. **Transizione** – dissolvenza, spingi, wipe, zoom… con direzione, velocità e avanzamento automatico.
5. **Genera** – scarica il `.pptx` (una slide per relatore, opzionalmente con le altre slide del template).

**Modelli salvati**: salvi template, stile, banner e posizioni sotto un nome (cliente/evento). Si riapre da solo l'ultimo usato; alla volta dopo basta caricare il programma e premere «Inserisci e genera subito». I modelli si possono esportare/importare per usarli su un altro PC.

Il font scelto deve essere installato sul PC che apre il file.

Librerie incluse in `lib/` (JSZip, PptxGenJS, pdf.js): nessuna connessione necessaria.
