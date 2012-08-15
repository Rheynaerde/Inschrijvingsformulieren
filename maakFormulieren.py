try: 
    from xml.etree import ElementTree
except ImportError:
    from elementtree import ElementTree
import gdata.spreadsheet.service
import gdata.service
import atom.service
import gdata.spreadsheet
import atom

from datetime import date

startjaarNieuwSeizoen = 2012

eindeSeizoen = date(startjaarNieuwSeizoen+1, 6, 30)

from private import *

gd_client.source = 'inschrijvingsformulierscript'
gd_client.ProgrammaticLogin()

documentLaTeX = """\\documentclass{{article}}

\\usepackage[margin=1.5in]{{geometry}}
\\usepackage{{booktabs}}
\\usepackage{{tikz}}
\\usetikzlibrary{{intersections}}

\\pagestyle {{empty}}
\\begin{{document}}

{inhoud}

\\end{{document}}
"""

paginaEindeLaTeX = """

\\clearpage

"""

paginaLaTeX = """
\\begin{{center}}\\bf\\Large
Inschrijvingsformulieren seizoen 2012-2013
\\end{{center}}

\\noindent Gelieve dit formulier na te lezen. Indien er ontbrekende of foutieve gegevens zijn, 
pas deze dan aan op een duidelijke manier. Bezorg daarna het ondertekende formulier samen met 
de benodigde bijlagen terug aan het secretariaat van Degencentrum Rheynaerde vzw.

\\bigskip

\\noindent\\textbf{{Gegevens}}
\\begin{{quote}}
\\noindent\\begin{{tabular}}{{@{{}}ll}}
Naam: & {lid.naam}\\\\
& \\\\
Voornaam: & {lid.voornaam}\\\\
& \\\\
Geboortedatum: & {lid.geboortedatumFormat}\\\\
& \\\\
Geslacht: & {lid.geslacht}\\\\
& \\\\
Adres: & {lid.adres.regel1}\\\\
       & {lid.adres.regel2}\\\\
& \\\\
E-mail (lid): & {eigenmail}\\\\
& \\\\
E-mail ({label_anderemail}): & {anderemail}\\\\
& \\\\
Telefoon: & {lid.telefoon}\\\\
& \\\\
T-shirtmaat: & 1/2 \\  3/4 \\  5/6 \\  7/8 \\  9/11 \\  12/14 \\  S \\  M \\  L \\  XL \\  XXL \\  XXL \\\\
\\end{{tabular}}
\\end{{quote}}

\\medskip

\\noindent\\textbf{{Medisch attest}}
\\begin{{quote}}
\\noindent {medischAttest}
\\end{{quote}}

\\medskip

\\noindent\\textbf{{Lidgeld}}
\\begin{{quote}}
\\noindent Gelieve voor 1 oktober het lidgeld ten bedrage van {lid.lidgeld} euro\\footnote[1]{{Het vermelde bedrag 
houdt geen rekening met eventuele familiekorting waarop je recht zou kunnen hebben. Indien
meerdere personen uit uw gezin lid zijn, gelieve dan na te vragen wat het correcte bedrag is.}} te storten op 
rekeningnummer BE98 1420 6537 5193 (BIC: GEBA BE BB) op naam van \\textbf{{Degencentrum Rheynaerde vzw}} met als mededeling:
LIDGELD + $<$NAAM$>$.
\\end{{quote}}

\\bigskip
\\bigskip

\\noindent Door te ondertekenen gaat u akkoord dat Degencentrum Rheynaerde vzw een schermlicentie
aanvraagt voor {lid.volledigenaam} en bevestigt u dat u akkoord bent met het huishoudelijk reglement.

\\bigskip

\\noindent \\textbf{{Datum}}: \\underline{{\\ \\ \\ \\ \\ \\ }}/\\underline{{\\ \\ \\ \\ \\ \\ }}/\\underline{{\\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ }}

\\medskip

\\noindent \\textbf{{Naam en handtekening {handtekening}}}:
\\bigskip
\\bigskip

\\clearpage

\\begin{{center}}

\\begin{{tikzpicture}}[>=latex, scale = 2]

\\useasboundingbox (0,0) rectangle (4.5,5);

\\coordinate (rCollar) at (1.5,3.6);
\\coordinate (rSleeveTop) at (0,3.5);
\\coordinate (rSleeveBottom) at (0,2.5);
\\coordinate (rSleeveArmpit) at (1.1,2.5);
\\coordinate (rBottom) at (1,0);

\\coordinate (lCollar) at (3,3.6);
\\coordinate (lSleeveTop) at (4.5,3.5);
\\coordinate (lSleeveBottom) at (4.5,2.5);
\\coordinate (lSleeveArmpit) at (3.4,2.5);
\\coordinate (lBottom) at (3.5,0);

\\coordinate (lShoulderAbove) at (3.4,4);
\\coordinate (lSleeveAbove) at (4.5,4);
\\coordinate (lShoulderBeside) at (5,3.6);
\\coordinate (lBottomBeside) at (5,0);

\\draw [ultra thick, name path=tshirt] (rCollar) -- (rSleeveTop) -- (rSleeveBottom) -- (rSleeveArmpit) -- (rBottom) -- (lBottom) -- (lSleeveArmpit) -- (lSleeveBottom) -- (lSleeveTop) -- (lCollar) -- (rCollar) --cycle;

\\draw [name path=shoulderVertical] (lSleeveArmpit) -- (lShoulderAbove);
\\draw (lSleeveTop) -- (lSleeveAbove);
\\draw [name intersections={{of=tshirt and shoulderVertical}}] (intersection-3) -- (lShoulderBeside);
\\draw (lBottom) -- (lBottomBeside);

\\path (rCollar) edge [ultra thick, double, out=-90, in=-90] (lCollar);

\\draw [very thick, gray,<->] (rSleeveArmpit.east) to node[below] {{A}} (lSleeveArmpit.west);
\\draw [very thick, gray,<->] (lShoulderBeside.south) to node[right] {{B}} (lBottomBeside.north);
\\draw [very thick, gray,<->] (lShoulderAbove.east) to node[above] {{C}} (lSleeveAbove.west);

\\end{{tikzpicture}}

\\bigskip

Kinderen:

\\begin{{tabular}}{{lcccccc}}
& \\rule{{1cm}}{{0pt}} & \\rule{{1cm}}{{0pt}} & \\rule{{1cm}}{{0pt}}& \\rule{{1cm}}{{0pt}} & \\rule{{1cm}}{{0pt}} & \\rule{{1cm}}{{0pt}}\\\\
\\toprule
& 1/2 & 3/4 & 5/6 & 7/8 & 9/11 & 12/14\\\\
\\midrule
A Halve Borst & 30 & 33 & 36 & 39 & 42 & 45\\\\
B Lichaamslengte & 40 & 43 & 46 & 51 & 56 & 63\\\\
C Mouwlengte & 12 & 13 & 14 & 15 & 16,5 & 18\\\\
\\bottomrule
\\end{{tabular}}

\\bigskip

Volwassenen:

\\begin{{tabular}}{{lcccccc}}
& \\rule{{1cm}}{{0pt}} & \\rule{{1cm}}{{0pt}} & \\rule{{1cm}}{{0pt}}& \\rule{{1cm}}{{0pt}} & \\rule{{1cm}}{{0pt}} & \\rule{{1cm}}{{0pt}}\\\\
\\toprule
& S & M & L & XL & XXL & XXXL\\\\
\\midrule
A Halve Borst & 50 & 53 & 56 & 59 & 62 & 65\\\\
B Lichaamslengte & 69 & 72 & 74 & 76 & 78 & 81\\\\
C Mouwlengte & 20 & 20 & 20 & 21 & 21 & 21\\\\
\\bottomrule
\\end{{tabular}}

\\end{{center}}

"""

medischAttestOK = """We beschikken nog over een geldig medisch attest."""
medischAttestGeen = """We beschikken niet over een geldig medisch attest.
Gelieve zo snel mogelijk een medisch attest binnen te brengen.
Dit is nodig om verzekerd te zijn tijdens het schermen."""
medischAttestNodig = """Het laatste medisch attest dateert van {datum}.
Gelieve zo snel mogelijk een nieuw medisch attest binnen te brengen.
Dit is nodig om verzekerd te zijn tijdens het schermen."""

class adres:
    def __init__(self, straat, nummer, postcode, stad):
        if straat == None: straat = ''
        if nummer == None: nummer = ''
        if postcode == None: postcode = ''
        if stad == None: stad = ''
        self.straat = straat
        self.nummer = nummer
        self.postcode = postcode
        self.stad = stad
        self.regel1 = straat + ' ' + nummer
        self.regel2 = postcode + ' ' + stad


class lid:

    geslachten = {'M':'Man', 'V':'Vrouw'}
    maanden = ['januari', 'februari', 'maart', 'april', 'mei', 'juni', 'juli', 'augustus', 'september', 'oktober', 'november', 'december']
    
    @classmethod
    def createLidFromFeedEntry(cls, feedEntry):
        voornaam = feedEntry.custom['voornaam'].text
        naam = feedEntry.custom['naam'].text
        adress = adres(feedEntry.custom['straat'].text,feedEntry.custom['nummer'].text,feedEntry.custom['plaats'].text,feedEntry.custom['gemeente'].text)
        nationaliteit = feedEntry.custom['nationaliteit'].text
        geslacht = lid.geslachten[feedEntry.custom['geslacht'].text]
        if feedEntry.custom['geboortedatum'].text:
            datumList = feedEntry.custom['geboortedatum'].text.split('/')
            geboortedatum = date(int(datumList[2]), int(datumList[0]), int(datumList[1]))
        else:
            geboortedatum = None
        if feedEntry.custom['medischattest'].text == None:
            medischattest = None
        else:
            datumList = feedEntry.custom['medischattest'].text.split('/')
            medischattest = date(int(datumList[2]), int(datumList[0]), int(datumList[1]))
        if feedEntry.custom['eigenmailadres'].text == None:
            eigenmail = 'Geen'
        else:
            eigenmail = feedEntry.custom['eigenmailadres'].text
        if feedEntry.custom['anderemailadres'].text == None:
            anderemail = 'Geen'
        else:
            anderemail = feedEntry.custom['anderemailadres'].text
        if feedEntry.custom['telefoon'].text == None:
            telefoon = 'Geen'
        else:
            telefoon = feedEntry.custom['telefoon'].text
        return lid(voornaam, naam, adress, nationaliteit, geslacht, geboortedatum, medischattest, eigenmail, anderemail, telefoon)

    def __init__(self, voornaam, naam, adres, nationaliteit, geslacht, geboortedatum, medischattest, eigenmail, anderemail, telefoon):
        self.voornaam = voornaam
        self.naam = naam
        self.volledigenaam = self.voornaam + ' ' + self.naam
        self.adres = adres
        self.nationaliteit = nationaliteit
        self.geslacht = geslacht
        self.geboortedatum = geboortedatum
        if geboortedatum:
            self.geboortedatumFormat = '%d %s %d' % (geboortedatum.day, lid.maanden[geboortedatum.month - 1], geboortedatum.year)
        else:
            self.geboortedatumFormat = ''
        self.medischattest = medischattest
        self.eigenmail = eigenmail
        self.anderemail = anderemail
        self.telefoon = telefoon
        
        if self.leeftijd(eindeSeizoen) < 18:
            self.lidgeld = 180
        else:
            self.lidgeld = 240
        
    def __str__(self):
        return self.volledigenaam
    
    def __repr__(self):
        return self.volledigenaam
   
    def leeftijd(self, datum=date.today()):
        if self.geboortedatum == None:
            return
        try: # raised when birth date is February 29 and the current year is not a leap year
            verjaardag = self.geboortedatum.replace(year=datum.year)
        except ValueError:
            verjaardag = self.geboortedatum.replace(year=datum.year, day=self.geboortedatum.day-1)
        if verjaardag > datum:
            return datum.year - self.geboortedatum.year - 1
        else:
            return datum.year - self.geboortedatum.year


def verjaring(fromDate, datum=date.today()):
    try: # raised when birth date is February 29 and the current year is not a leap year
        verjaardag = fromDate.replace(year=datum.year)
    except ValueError:
        verjaardag = fromDate.replace(year=datum.year, day=fromDate.day-1)
    if verjaardag > datum:
        return datum.year - fromDate.year - 1
    else:
        return datum.year - fromDate.year
    
rowsFeed = gd_client.GetListFeed(spreadsheetKey, worksheetId)

def printLedenlijst(feed):
    assert isinstance(feed, gdata.spreadsheet.SpreadsheetsListFeed)
    for i, entry in enumerate(feed.entry):
        if entry.custom['seizoen2011'].text == '1':
            l = lid.createLidFromFeedEntry(entry)
            print '%d) %s %s (%d)' % (i, l.voornaam, l.naam, l.leeftijd())

def feed2Lijst(feed, seizoen=None):
    return [lid.createLidFromFeedEntry(entry) for entry in feed.entry if seizoen==None or entry.custom['seizoen%d' % seizoen].text == '1']

#print rowsFeed.entry[0].custom.keys()
#for key in rowsFeed.entry[0].custom.keys():
#    print key, rowsFeed.entry[0].custom[key].text


#printLedenlijst(rowsFeed)


def maakFormulier(lid):
    if lid.leeftijd() < 18:
        labelAnderemail = 'ouder'
        handtekening = 'ouder of wettelijke voogd'
    else:
        labelAnderemail = 'extra'
        handtekening = 'schermer'
    if lid.medischattest == None:
        medischAttest = medischAttestGeen
    elif lid.leeftijd(lid.medischattest) < 18 or verjaring(lid.medischattest, eindeSeizoen) > 5:
        medischAttest = medischAttestNodig.format(datum = '%d/%d/%d' % (lid.medischattest.day, lid.medischattest.month, lid.medischattest.year))
    else:
        medischAttest = medischAttestOK
    return paginaLaTeX.format(lid=lid, medischAttest=medischAttest, eigenmail=lid.eigenmail.replace('_', '\_'), anderemail=lid.anderemail.replace('_', '\_'), label_anderemail=labelAnderemail, handtekening=handtekening)

paginas = [maakFormulier(lid) for lid in feed2Lijst(rowsFeed, startjaarNieuwSeizoen-1)] #maak lijst op basis van leden van vorig jaar

print documentLaTeX.format(inhoud=paginaEindeLaTeX.join(paginas))


