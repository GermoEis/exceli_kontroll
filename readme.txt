Exceli Kontroll - Kasutusjuhend ja Ülevaade
See on väike utiliit, mis koondab mitmed Exceli ja XML failide kontrolli- ning võrdlusvahendid ühte aknasse. 
Peamine eesmärk on lihtsustada meie töövoogu, kus tuleb tihti teha suuri Exceli võrdlusi, kuupäevade kontrolli ja XML failide muutmist.

Mis see on?
Exceli Kontroll on GUI-põhine programm, mis käivitab erinevad tööriistad (exe failid), mis teevad spetsiifilisi tööülesandeid:

XML failide muutmine

Suure Exceli tabeli võrdlus ja kontroll

Kuupäevade kontroll nädalate alusel

Üldine Exceli kontroll

Lisaks kontrollib programm automaatselt, kas on saadaval uuem versioon ja annab võimaluse selle kiiresti alla laadida ja paigaldada.

Kuidas see töötab?
Kui käivitad programmi, avaneb lihtne aknake, kus saad valida vajalikku tööriista.

Programm kontrollib taustal internetist, kas on uus versioon (vaadates teksti failist versioon.txt GitLabis).

Kui uuendus on olemas, ilmub punane märge ja "Uuenda" nupp, mille vajutades programm:

Laeb GitLabist alla uue versiooni ZIP failina.

Pakkib selle lahti ajutises kaustas.

Kopeerib failid üle jooksva programmi kausta, asendades vanad.

Värskendab versiooninumbrit kohalikus failis.

Kui uuendus ei ole saadaval või oled juba viimase versiooniga, kuvatakse sellest info.

Mida peaks teadma?
Programmi tööriistad töötavad kui eraldi .exe failid, seega peab neid olema sama kaustas, kus GUI.

Internetiühendus on vajalik versiooni kontrollimiseks ja uuenduse allalaadimiseks.

Uuendamine on automaatne ja lihtne, ei pea käsitsi midagi lahti pakkima või kopeerima.

Kui midagi ei tööta, annab programm kas terminali aknas või hüpikaknas veateate.

Kuidas käivitada?
Ava fail kontroll.exe (või käivita skript, kui kasutad .py versiooni).

Vali sobiv tööriist nupuvajutusega.

Kasuta "Uuenda" nuppu, kui ilmub teade uuenduse kohta.

Sulge programm nupuga "Sulge".
