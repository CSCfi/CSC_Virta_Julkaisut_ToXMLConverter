using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;
using Excel = Microsoft.Office.Interop.Excel; 

// Viimeisin päivitetty sorsa ohjelmaversio 1.2.2 2017-03-12 
// Virhenavigaattori
// XML validointia varten oma virheloki
// Tiedote tai huomautus jos on Generoitu julkaisunorganisaatiotunnus
// Tuki 2017 julkaisuille
// JosJulkaisuTyyppille_LehdenNimi_tai_EmojulkaisunNimi_pakollinen
// TarkistaKonferenssinNimiPakollisuus
// Orcid virhe korjattu
// EmojulkaisunToimittajatTeksti 500 merkin rajoite
// TarkistaKustantajaPakollisuus vuodelle 2017
// Ei JulkaisunTunnus riittää OrganisaationJulkaisunTunnus
// Tuki 2016 csv-mallitiedostolle
// Uusi sarake "julkaisun tila" ja molemilla "julkaisun organisaatiokohtainen id"
// Indexointi meni uusiksi, uusia sarakkeita ja poistettuja sarakkeita
// csv-tarkistustiedosto

// URN pysyvä osoite, mikäli URN:ISBN:978-951-51-1829-5 niin eteen http://urn.fi => http://urn.fi/URN:ISBN:978-951-51-1829-5
// Rivinro missä korjataan isbn maakoodi jne 2016-03-18


/* *******************************************************************
 * 
 *  
 * CSC - Virta-Julkaisutietojen CSV-XML muuntotyökalu (C) 2016, 2017
 * 
 * 
 * Työkalu voi olla avuksi kun halutaan muuntaa Virta-Julkaisutietovarantoa 
 * varten xml-tiedosto csv-muodossa olevasta lähdetiedostosta 
 * 
 * Muuhun kuin Julkaisutietojen xml-tiedostojen tuottamiseen ohjelma ei ole tarkoitettu
 * 
 * 
 * Ohjelmaa on tällä hetkellä ensisijassa tarkoitettu vuoden 2015-2016 tiedonkeruita varten.
 * Lisäominaisuuksia tehdään tarpeen mukaan vuoden 2016 aikana, ennen seuraavaa tilastointivuotta 2017.
 * 
 * Korjaukset käynnissä olevaan tiedonkeruun tarpeisiin tehdään niin pian kuin mahdollista
 * Bugi-ilmoitukset voi lähettää osoitteeseen virta-julkaisut@postit.csc.fi 
 * 
 * Valmis käännetty ohjelman exe-tiedosto ja ohjeet löytyy Virta-Julkaisutietopalvelujen sivuilta.
 * Lisäksi sivuilla löytyy tiedonkeruujen malli csv-tiedostot ja viimeisin XML-skeema (xsd)
 * 
 * 
 * Viimeisimmän version tästä lähdekoodista csv-xml-ohjelmalle, 
 * löydät 
 * https://github.com/CSCfi/CSC_Virta_Julkaisut_ToXMLConverter
 * 
 * 
 * Ottakaa rohkeasti yhteyttä jos on kysymyksiä lähdekoodista tai itse ohjelman käytöstä   
 * fredrik.finnberg@csc.fi tai virta-julkaisut@postit.csc.fi 
 * 
 * Kiitos kiinnostuksesta!
 * 
 * *******************************************************************
 */


// csv tiedostojen merkistön, UTF-8 ilman BOM (ByteOrderMark), tarkistukseen tarvitaan tämä kirjastoluokka 
// jota käytetään KlerksSoft.dll assemlblyn kautta, käännetty 
// lähdekoodista TextFileEncodingDetector.cs tässä projektissa) 

using KlerksSoft;   // KlerksSoft TextFileEncodingDetector luokka https://gist.github.com/TaoK/945127
                    // metodeja jonka avulla ohjelma analysoi jos lähdetiedosto on UTF8 BOM tai ilman w/o BOM 
                    // Esimerkki UTF8 ohjeet CSC:n sivulta: https://confluence.csc.fi/pages/viewpage.action?pageId=36604464



// Julkaisut-CSV-XML-Työkalu luokka 

namespace CSC_Virta_Julkaisut_ToXMLConverter
{
    public partial class CSC_VIRTA_JulkaisutForm : Form
    {
        public CSC_VIRTA_JulkaisutForm()
        {
            InitializeComponent();
            TallennaXMLButton.Enabled = false;
            ValidoiButton.Enabled = false;
            AvaaXMLButton.Enabled = false;
            MuunnaXMLButton.Enabled = false;
            ValidoiXMLButton.Enabled = false;

            numericUpDownVirhe.Enabled = false;


            System.Windows.Forms.ToolTip ToolTip_XMLButton = new System.Windows.Forms.ToolTip();
            ToolTip_XMLButton.SetToolTip(this.TallennaXMLButton, "Tallenna XML-tiedosto");

            System.Windows.Forms.ToolTip ToolTip_ValidoiButton = new System.Windows.Forms.ToolTip();
            ToolTip_ValidoiButton.SetToolTip(this.ValidoiButton, "Tarkista CSV aineisto");

            System.Windows.Forms.ToolTip toolTipVirheNavigaattori = new System.Windows.Forms.ToolTip();
            toolTipVirheNavigaattori.SetToolTip(this.numericUpDownVirhe, "Siirry seuraavaan virhekohtaan...");

            


            errorTextBox.Clear();
            LokiTiedosto("loki.txt");
            CSV_virhe_LokiTiedosto("tarkistusloki.csv");

            Poista_XML_virheLokiTiedosto("xmlvirheloki.txt");

        }


        // Luodaan uusi lokitiedosto jos ei ole jo olemassa
        private void LokiTiedosto(string nimi)
        {
            string lokitiedostonimi = nimi.Trim();

            // Ei ole olemassa jo?
            if (!File.Exists(lokitiedostonimi))
            {
                // Luodaan masterloki loki.txt jos ei ole olemassa
                try
                {
                    StreamWriter lokifile = new StreamWriter("loki.txt", true);

                    lokifile.WriteLine("--------------------------------------------------------------");
                    lokifile.WriteLine("Uusi loki.txt tiedosto luotiin.");
                    lokifile.WriteLine("Ohjelma käynnistyi: " + "{0} klo {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString() );
                    lokifile.WriteLine("--------------------------------------------------------------");
                    
                    lokifile.Close();
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Lokitiedostoa ei voitu luoda." + ex.Message);
                }

            }
            else   // Avataan olemassa oleva masterloki
            {
                try
                {
                    StreamWriter lokifile = new StreamWriter("loki.txt", true);

                    lokifile.WriteLine("--------------------------------------------------------------");
                    lokifile.WriteLine("Ohjelma käynnistyi uudestaan: " + "{0} klo {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString());
                                           
                    lokifile.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lokitiedostoa ei voitu avata." + ex.Message);
                }
            }

        }
        
        // Virheet talteen, rivi ja sarake
        public List<int> virheRivi   = new List<int>();
        public List<int> virheSarake = new List<int>();

        // Generoidaanko Julkaisun organisaatiotunnus true = kyllä, false = ei
        public static bool GeneroiJulkaisunOrganisaatioTunnus = true;
        
        // Ilmoitusvuosi, joka vaihdetaan ohjelmassa Työkalut/Asetukset kautta
        public static int vuosiIlmo = 2016;
       
        // Mikä vuosi luetussa csv-ainestossa, vakio arvo 
        public int vuosiMuuttuja = 2016;

        // Tulevan XML tiedoston nimi
        public static string UusiXMLFile = "";

        // Ohjelman masterloki-tiedoston nimi
        public static string masterloki = "loki.txt";

        // CSV muotoinen tarkistusloki
        public static string csvtarkistusloki = "tarkistusloki.csv";

        // Ohjelman XML-virheloki
        public static string xmlvirheloki = "xmlvirheloki.txt";


        // Organisaatio koodi ainestosta
        public string organisaatiokoodi = "";

        public DataTable virheTaulukko = new DataTable("Virheet");

        // Luodaan uusi CSV_virhe_LokiTiedosto jos ei ole jo olemassa
        private void CSV_virhe_LokiTiedosto(string nimi)
        {
            string lokitiedostonimi = nimi.Trim();

            // olemassa jo, poistetaan
            if (File.Exists(lokitiedostonimi))
            {
                try
                {
                    File.Delete(lokitiedostonimi);
                }
                catch (Exception delex)
                {
                    MessageBox.Show("tarkistusloki.csv tiedostoa ei voitu poistaa." + delex.Message);
                }
            }

            // Luodaan tarkistusloki.csv jos ei ole olemassa
            try
            {
                StreamWriter csvlokifile = new StreamWriter("tarkistusloki.csv", false);    
                csvlokifile.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show("tarkistusloki.csv tiedostoa ei voitu luoda." + ex.Message);
            }

        }


        // Poistetaan jos olemassa
        private void Poista_XML_virheLokiTiedosto(string nimi)
        {
            string xmlvirheloki_nimi = nimi.Trim();

            // olemassa jo, poistetaan
            if (File.Exists(xmlvirheloki_nimi))
            {
                try
                {
                    File.Delete(xmlvirheloki_nimi);
                }
                catch (Exception delex)
                {
                    MessageBox.Show("xmlvirheloki.txt tiedostoa ei voitu poistaa." + delex.Message);
                }
            }

        }	



        // Hae mikä vuosi
        private int getVuosi()
        {
            return vuosiMuuttuja;
        }            

        // Aseta vuosi
        private void setVuosi( int vuosi)
        {
            vuosiMuuttuja = vuosi;

            mikaVuosilabel.Text = vuosi.ToString();

            vuosiIlmo = vuosiMuuttuja;

        }

        // Hae XML-tiedoston nimi
        private string getXMLTiedosto()
        {
            return UusiXMLFile;
        }

        // Aseta XML-tiedoston nimi
        private void setXMLTiedostoNimi(string nimi)
        {
            UusiXMLFile = nimi.Trim();

        }

        // Hae organisaatiokoodi
        private string getOrganisaatioKoodi()
        {
            return organisaatiokoodi;
        }


        // Aseta organisaatiokoodi  
        private void setOrganisaatioKoodi(string koodi)
        {
            organisaatiokoodi = koodi.Trim();
        }

        // Kerätään virheitä virhetaulukkoon
        private void RekisteroiVirheet(string virhe_puute, string julkorgtunnus, string rivi, string tieto, string viesti)
        {
            object[] virheLista = new object[5];

            virheLista[0] = virhe_puute + ";";
            virheLista[1] = julkorgtunnus + ";";
            virheLista[2] = rivi   + ";";
            virheLista[3] = tieto  + ";";
            virheLista[4] = viesti + ";"; 

            virheTaulukko.Rows.Add(virheLista);
        }

     
        // Jos lokia ei ole olemassa odotetussa paikassa, eli juuressa, ei voi kirjoittaa lokiin
        private static bool LokiOlemassa(string loki)
        {
            string lokitiedostonimi = loki.Trim();

            if (File.Exists(lokitiedostonimi))
                return true;
            else
                return false;

        }

        // Avataan loki ja kirjoitetaan lokitietoja		
        private static void KirjoitaLokiin(string lokitieto, string lokitiedosto)
        {
            StreamWriter loki = new StreamWriter(lokitiedosto, true);		  
		    KirjoitaLokiTieto( lokitieto,  loki);		  		  
        }
		
		
		// Kirjoitetaan tietoja lokitiedostoon
        private static void KirjoitaLokiTieto(string lokitieto, TextWriter tw)
        {
            try
            {
                tw.WriteLine("--------------------------------------------------------------");
            
                tw.WriteLine("{0} {1}:", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString());
                tw.WriteLine("{0}", lokitieto.Trim());                                
                
                tw.Flush();
                tw.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lokitiedostoa ei voitu päivittää." + ex.Message);
            } 			
        }		


        // Montako saraketta meillä on ensimmäisellä rivillä csv-aineistossa?
        private int MontakoSaraketta(string tiedosto, char separattori)
        {
            var otsikko = File.ReadAllLines(tiedosto);
            int sarakkeet = otsikko[0].Split(separattori).Length;
                        
            return sarakkeet;
        }

        // Montako riviä meillä on csv-tiedostossa?
        private int MontakoRivia(string tiedosto )
        {
            int riveja = File.ReadAllLines(tiedosto).Length;          

            return riveja;
        }



        // Tarkista että csv sarakkeet ovat ok, eli nimet ja niiden järjestys pitää olla mallitiedoston  mukaiset
        private Boolean TarkistaSarakkeetCSVsta(List<string[]> parsedData, int num )
        {            
            int maximi = 0;

            if (getVuosi() == 2018)
            {
                maximi = 62;
            }


            if (getVuosi() == 2017)
            {
                maximi = 62;
            }


            if (getVuosi() == 2016)
            {
                maximi = 62;
            }

            if (getVuosi() == 2015)
            {
                maximi = 66;
            }

            if (getVuosi() == 2014)
            {
                maximi = 66;
            }

            if (getVuosi() == 2013)
            {
                maximi = 46;
            }

            //MessageBox.Show(num.ToString(), "num Parsed");

			// Luetaan parsedDatan sarakkeet				
			String[] sarakkeetCSV = new String[maximi]; 			
            
            for (int i = 0; i < maximi; i++) {			
	
                var sb = new StringBuilder(parsedData[0][i]);				
                sb.Replace("\"", "");  // Siivotaan "-merkkejä jos löytyy nimessä
                sarakkeetCSV[i] = sb.ToString();  // Sarakkeet ensimmäiseltä riviltä csv parsedData				
            }           
			
			// Tiedostonmukaiset sarakenimet vuosille 2015 ja 2016, eikä sitä vanhemmille tiedoille tässä vaiheessa
            String[] CSVSarakkeet2015_amk = { "Ammattikorkeakoulu", "Vuosi", "Julkaisutyyppi", "I Tieteenala", "II Tieteenala", "III Tieteenala", "IV Tieteenala", "V Tieteenala", "VI Tieteenala", "I Koulutusala", "II Koulutusala", "III Koulutusala", "IV Koulutusala", "V Koulutusala", "VI Koulutusala", "Organisaation tekijät", "I Organisaation alayksikkö", "II Organisaation alayksikkö", "III Organisaation alayksikkö", "IV Organisaation alayksikkö", "V Organisaation alayksikkö", "VI Organisaation alayksikkö", "VII Organisaation alayksikkö", "VIII Organisaation alayksikkö", "IX Organisaation alayksikkö", "X Organisaation alayksikkö", "XI Organisaation alayksikkö", "XII Organisaation alayksikkö", "XIII Organisaation alayksikkö", "XIV Organisaation alayksikkö", "XV Organisaation alayksikkö", "XVI Organisaation alayksikkö", "XVII Organisaation alayksikkö", "XVIII Organisaation alayksikkö", "XIX Organisaation alayksikkö", "XX Organisaation alayksikkö", "Julkaisun tekijät", "Julkaisun tekijöiden lukumäärä", "Kansainvälinen yhteisjulkaisu", "Yliopistollinen sairaanhoitopiiri, yhteisjulkaisu", "Valtion sektoritutkimuslaitos, yhteisjulkaisu", "Muu kotimainen tutkimusorganisaatio, yhteisjulkaisu", "Julkaisun nimi", "Julkaisuvuosi", "Volyymi", "Numero", "Sivut", "Artikkelinumero", "Julkaisun kieli", "Lehden/sarjan nimi", "ISSN", "ISBN", "Emojulkaisun nimi", "Emojulkaisun toimittajat", "Kustantaja", "Julkaisun kustannuspaikka", "Julkaisumaa", "Julkaisun kansainvälisyys", "DOI-tunniste", "Pysyvä verkko-osoite", "Avoin saatavuus", "Lähdetietokannan koodi",  "Julkaisun julkaisukanava (JUFO-ID)", "Julkaisun id", "Konferenssin vakiintunut nimi", "Avainsanat" };
            String[] CSVSarakkeet2016_amk = { "Ammattikorkeakoulu", "Vuosi", "Julkaisutyyppi", "I Tieteenala", "II Tieteenala", "III Tieteenala", "IV Tieteenala", "V Tieteenala", "VI Tieteenala", "Organisaation tekijät", "I Organisaation alayksikkö", "II Organisaation alayksikkö", "III Organisaation alayksikkö", "IV Organisaation alayksikkö", "V Organisaation alayksikkö", "VI Organisaation alayksikkö", "VII Organisaation alayksikkö", "VIII Organisaation alayksikkö", "IX Organisaation alayksikkö", "X Organisaation alayksikkö", "XI Organisaation alayksikkö", "XII Organisaation alayksikkö", "XIII Organisaation alayksikkö", "XIV Organisaation alayksikkö", "XV Organisaation alayksikkö", "XVI Organisaation alayksikkö", "XVII Organisaation alayksikkö", "XVIII Organisaation alayksikkö", "XIX Organisaation alayksikkö", "XX Organisaation alayksikkö", "Julkaisun tekijät", "Julkaisun tekijöiden lukumäärä", "Kansainvälinen yhteisjulkaisu", "Yhteisjulkaisu yrityksen kanssa", "Julkaisun nimi", "Julkaisuvuosi", "Volyymi", "Numero", "Sivut", "Artikkelinumero", "Julkaisun kieli", "Lehden/sarjan nimi", "ISSN", "ISBN", "Emojulkaisun nimi", "Emojulkaisun toimittajat", "Kustantaja", "Julkaisun kustannuspaikka", "Julkaisumaa", "Julkaisun kansainvälisyys", "DOI-tunniste", "Pysyvä verkko-osoite", "Avoin saatavuus", "Lähdetietokannan koodi", "Julkaisun julkaisukanava (JUFO-ID)", "Julkaisun organisaatiokohtainen id", "Konferenssin vakiintunut nimi", "Avainsanat", "Julkaisu rinnakkaistallennettu", "Rinnakkaistallennetun version verkko-osoite", "ORCID", "Julkaisun tila" };

    
            String[] CSVSarakkeet2015 = { "Organisaatiotunnus", "Vuosi", "Julkaisutyyppi", "I Tieteenala", "II Tieteenala", "III Tieteenala", "IV Tieteenala", "V Tieteenala", "VI Tieteenala", "I Koulutusala", "II Koulutusala", "III Koulutusala", "IV Koulutusala", "V Koulutusala", "VI Koulutusala", "Organisaation tekijät", "I Organisaation alayksikkö", "II Organisaation alayksikkö", "III Organisaation alayksikkö", "IV Organisaation alayksikkö", "V Organisaation alayksikkö", "VI Organisaation alayksikkö", "VII Organisaation alayksikkö", "VIII Organisaation alayksikkö", "IX Organisaation alayksikkö", "X Organisaation alayksikkö", "XI Organisaation alayksikkö", "XII Organisaation alayksikkö", "XIII Organisaation alayksikkö", "XIV Organisaation alayksikkö", "XV Organisaation alayksikkö", "XVI Organisaation alayksikkö", "XVII Organisaation alayksikkö", "XVIII Organisaation alayksikkö", "XIX Organisaation alayksikkö", "XX Organisaation alayksikkö", "Julkaisun tekijät", "Julkaisun tekijöiden lukumäärä", "Kansainvälinen yhteisjulkaisu", "Yliopistollinen sairaanhoitopiiri, yhteisjulkaisu", "Valtion sektoritutkimuslaitos, yhteisjulkaisu", "Muu kotimainen tutkimusorganisaatio, yhteisjulkaisu", "Julkaisun nimi", "Julkaisuvuosi", "Volyymi", "Numero", "Sivut", "Artikkelinumero", "Julkaisun kieli", "Lehden/sarjan nimi", "ISSN", "ISBN", "Emojulkaisun nimi", "Emojulkaisun toimittajat", "Kustantaja", "Julkaisun kustannuspaikka", "Julkaisumaa", "Julkaisun kansainvälisyys", "DOI-tunniste", "Pysyvä verkko-osoite", "Avoin saatavuus", "Lähdetietokannan koodi",  "Julkaisun julkaisukanava (JUFO-ID)", "Julkaisun organisaatiokohtainen id", "Konferenssin vakiintunut nimi", "Avainsanat" };
            String[] CSVSarakkeet2016 = { "Organisaatiotunnus", "Vuosi", "Julkaisutyyppi", "I Tieteenala", "II Tieteenala", "III Tieteenala", "IV Tieteenala", "V Tieteenala", "VI Tieteenala", "Organisaation tekijät", "I Organisaation alayksikkö", "II Organisaation alayksikkö", "III Organisaation alayksikkö", "IV Organisaation alayksikkö", "V Organisaation alayksikkö", "VI Organisaation alayksikkö", "VII Organisaation alayksikkö", "VIII Organisaation alayksikkö", "IX Organisaation alayksikkö", "X Organisaation alayksikkö", "XI Organisaation alayksikkö", "XII Organisaation alayksikkö", "XIII Organisaation alayksikkö", "XIV Organisaation alayksikkö", "XV Organisaation alayksikkö", "XVI Organisaation alayksikkö", "XVII Organisaation alayksikkö", "XVIII Organisaation alayksikkö", "XIX Organisaation alayksikkö", "XX Organisaation alayksikkö", "Julkaisun tekijät", "Julkaisun tekijöiden lukumäärä", "Kansainvälinen yhteisjulkaisu", "Yhteisjulkaisu yrityksen kanssa", "Julkaisun nimi", "Julkaisuvuosi", "Volyymi", "Numero", "Sivut", "Artikkelinumero", "Julkaisun kieli", "Lehden/sarjan nimi", "ISSN", "ISBN", "Emojulkaisun nimi", "Emojulkaisun toimittajat", "Kustantaja", "Julkaisun kustannuspaikka", "Julkaisumaa", "Julkaisun kansainvälisyys", "DOI-tunniste", "Pysyvä verkko-osoite", "Avoin saatavuus", "Lähdetietokannan koodi", "Julkaisun julkaisukanava (JUFO-ID)", "Julkaisun organisaatiokohtainen id", "Konferenssin vakiintunut nimi", "Avainsanat", "Julkaisu rinnakkaistallennettu", "Rinnakkaistallennetun version verkko-osoite", "ORCID", "Julkaisun tila" };

            // Jos on tallennettu suoraan Excelistä niin sarakenimet ovat todennäköisesti "väärin" eli ä => "Ã¤", ö => "Ã¶" jne
            String[] CSVSarakkeet2015_utf_vaarin_amk = { "Ammattikorkeakoulu", "Vuosi", "Julkaisutyyppi", "I Tieteenala", "II Tieteenala", "III Tieteenala", "IV Tieteenala", "V Tieteenala", "VI Tieteenala", "I Koulutusala", "II Koulutusala", "III Koulutusala", "IV Koulutusala", "V Koulutusala", "VI Koulutusala", "Organisaation tekijÃ¤t", "I Organisaation alayksikkÃ¶", "II Organisaation alayksikkÃ¶", "III Organisaation alayksikkÃ¶", "IV Organisaation alayksikkÃ¶", "V Organisaation alayksikkÃ¶", "VI Organisaation alayksikkÃ¶", "VII Organisaation alayksikkÃ¶", "VIII Organisaation alayksikkÃ¶", "IX Organisaation alayksikkÃ¶", "X Organisaation alayksikkÃ¶", "XI Organisaation alayksikkÃ¶", "XII Organisaation alayksikkÃ¶", "XIII Organisaation alayksikkÃ¶", "XIV Organisaation alayksikkÃ¶", "XV Organisaation alayksikkÃ¶", "XVI Organisaation alayksikkÃ¶", "XVII Organisaation alayksikkÃ¶", "XVIII Organisaation alayksikkÃ¶", "XIX Organisaation alayksikkÃ¶", "XX Organisaation alayksikkÃ¶", "Julkaisun tekijÃ¤t", "Julkaisun tekijÃ¶iden lukumÃ¤Ã¤rÃ¤", "KansainvÃ¤linen yhteisjulkaisu", "Yliopistollinen sairaanhoitopiiri, yhteisjulkaisu", "Valtion sektoritutkimuslaitos, yhteisjulkaisu", "Muu kotimainen tutkimusorganisaatio, yhteisjulkaisu", "Julkaisun nimi", "Julkaisuvuosi", "Volyymi", "Numero", "Sivut", "Artikkelinumero", "Julkaisun kieli", "Lehden/sarjan nimi", "ISSN", "ISBN", "Emojulkaisun nimi", "Emojulkaisun toimittajat", "Kustantaja", "Julkaisun kustannuspaikka", "Julkaisumaa", "Julkaisun kansainvÃ¤lisyys", "DOI-tunniste", "PysyvÃ¤ verkko-osoite", "Avoin saatavuus", "LÃ¤hdetietokannan koodi", "Julkaisun julkaisukanava (JUFO-ID)", "Julkaisun id", "Konferenssin vakiintunut nimi", "Avainsanat" };
            String[] CSVSarakkeet2015_utf_vaarin_yo  = { "Organisaatiotunnus", "Vuosi", "Julkaisutyyppi", "I Tieteenala", "II Tieteenala", "III Tieteenala", "IV Tieteenala", "V Tieteenala", "VI Tieteenala", "I Koulutusala", "II Koulutusala", "III Koulutusala", "IV Koulutusala", "V Koulutusala", "VI Koulutusala", "Organisaation tekijÃ¤t", "I Organisaation alayksikkÃ¶", "II Organisaation alayksikkÃ¶", "III Organisaation alayksikkÃ¶", "IV Organisaation alayksikkÃ¶", "V Organisaation alayksikkÃ¶", "VI Organisaation alayksikkÃ¶", "VII Organisaation alayksikkÃ¶", "VIII Organisaation alayksikkÃ¶", "IX Organisaation alayksikkÃ¶", "X Organisaation alayksikkÃ¶", "XI Organisaation alayksikkÃ¶", "XII Organisaation alayksikkÃ¶", "XIII Organisaation alayksikkÃ¶", "XIV Organisaation alayksikkÃ¶", "XV Organisaation alayksikkÃ¶", "XVI Organisaation alayksikkÃ¶", "XVII Organisaation alayksikkÃ¶", "XVIII Organisaation alayksikkÃ¶", "XIX Organisaation alayksikkÃ¶", "XX Organisaation alayksikkÃ¶", "Julkaisun tekijÃ¤t", "Julkaisun tekijÃ¶iden lukumÃ¤Ã¤rÃ¤", "KansainvÃ¤linen yhteisjulkaisu", "Yliopistollinen sairaanhoitopiiri, yhteisjulkaisu", "Valtion sektoritutkimuslaitos, yhteisjulkaisu", "Muu kotimainen tutkimusorganisaatio, yhteisjulkaisu", "Julkaisun nimi", "Julkaisuvuosi", "Volyymi", "Numero", "Sivut", "Artikkelinumero", "Julkaisun kieli", "Lehden/sarjan nimi", "ISSN", "ISBN", "Emojulkaisun nimi", "Emojulkaisun toimittajat", "Kustantaja", "Julkaisun kustannuspaikka", "Julkaisumaa", "Julkaisun kansainvÃ¤lisyys", "DOI-tunniste", "PysyvÃ¤ verkko-osoite", "Avoin saatavuus", "LÃ¤hdetietokannan koodi", "Julkaisun julkaisukanava (JUFO-ID)", "Julkaisun organisaatiokohtainen id", "Konferenssin vakiintunut nimi", "Avainsanat" };

            // Jos on tallennettu suoraan Excelistä niin sarakenimet ovat todennäköisesti "väärin" eli ä => "Ã¤", ö => "Ã¶" jne
            String[] CSVSarakkeet2016_utf_vaarin_amk = { "Ammattikorkeakoulu", "Vuosi", "Julkaisutyyppi", "I Tieteenala", "II Tieteenala", "III Tieteenala", "IV Tieteenala", "V Tieteenala", "VI Tieteenala", "Organisaation tekijÃ¤t", "I Organisaation alayksikkÃ¶", "II Organisaation alayksikkÃ¶", "III Organisaation alayksikkÃ¶", "IV Organisaation alayksikkÃ¶", "V Organisaation alayksikkÃ¶", "VI Organisaation alayksikkÃ¶", "VII Organisaation alayksikkÃ¶", "VIII Organisaation alayksikkÃ¶", "IX Organisaation alayksikkÃ¶", "X Organisaation alayksikkÃ¶", "XI Organisaation alayksikkÃ¶", "XII Organisaation alayksikkÃ¶", "XIII Organisaation alayksikkÃ¶", "XIV Organisaation alayksikkÃ¶", "XV Organisaation alayksikkÃ¶", "XVI Organisaation alayksikkÃ¶", "XVII Organisaation alayksikkÃ¶", "XVIII Organisaation alayksikkÃ¶", "XIX Organisaation alayksikkÃ¶", "XX Organisaation alayksikkÃ¶", "Julkaisun tekijÃ¤t", "Julkaisun tekijÃ¶iden lukumÃ¤Ã¤rÃ¤", "KansainvÃ¤linen yhteisjulkaisu", "Yhteisjulkaisu yrityksen kanssa", "Julkaisun nimi", "Julkaisuvuosi", "Volyymi", "Numero", "Sivut", "Artikkelinumero", "Julkaisun kieli", "Lehden/sarjan nimi", "ISSN", "ISBN", "Emojulkaisun nimi", "Emojulkaisun toimittajat", "Kustantaja", "Julkaisun kustannuspaikka", "Julkaisumaa", "Julkaisun kansainvÃ¤lisyys", "DOI-tunniste", "PysyvÃ¤ verkko-osoite", "Avoin saatavuus", "LÃ¤hdetietokannan koodi", "Julkaisun julkaisukanava (JUFO-ID)", "Julkaisun organisaatiokohtainen id", "Konferenssin vakiintunut nimi", "Avainsanat", "Julkaisu rinnakkaistallennettu", "Rinnakkaistallennetun version verkko-osoite", "ORCID", "Julkaisun tila" };
            String[] CSVSarakkeet2016_utf_vaarin_yo = { "Organisaatiotunnus", "Vuosi", "Julkaisutyyppi", "I Tieteenala", "II Tieteenala", "III Tieteenala", "IV Tieteenala", "V Tieteenala", "VI Tieteenala","Organisaation tekijÃ¤t", "I Organisaation alayksikkÃ¶", "II Organisaation alayksikkÃ¶", "III Organisaation alayksikkÃ¶", "IV Organisaation alayksikkÃ¶", "V Organisaation alayksikkÃ¶", "VI Organisaation alayksikkÃ¶", "VII Organisaation alayksikkÃ¶", "VIII Organisaation alayksikkÃ¶", "IX Organisaation alayksikkÃ¶", "X Organisaation alayksikkÃ¶", "XI Organisaation alayksikkÃ¶", "XII Organisaation alayksikkÃ¶", "XIII Organisaation alayksikkÃ¶", "XIV Organisaation alayksikkÃ¶", "XV Organisaation alayksikkÃ¶", "XVI Organisaation alayksikkÃ¶", "XVII Organisaation alayksikkÃ¶", "XVIII Organisaation alayksikkÃ¶", "XIX Organisaation alayksikkÃ¶", "XX Organisaation alayksikkÃ¶", "Julkaisun tekijÃ¤t", "Julkaisun tekijÃ¶iden lukumÃ¤Ã¤rÃ¤", "KansainvÃ¤linen yhteisjulkaisu", "Yhteisjulkaisu yrityksen kanssa", "Julkaisun nimi", "Julkaisuvuosi", "Volyymi", "Numero", "Sivut", "Artikkelinumero", "Julkaisun kieli", "Lehden/sarjan nimi", "ISSN", "ISBN", "Emojulkaisun nimi", "Emojulkaisun toimittajat", "Kustantaja", "Julkaisun kustannuspaikka", "Julkaisumaa", "Julkaisun kansainvÃ¤lisyys", "DOI-tunniste", "PysyvÃ¤ verkko-osoite", "Avoin saatavuus", "LÃ¤hdetietokannan koodi", "Julkaisun julkaisukanava (JUFO-ID)", "Julkaisun organisaatiokohtainen id", "Konferenssin vakiintunut nimi", "Avainsanat", "Julkaisu rinnakkaistallennettu", "Rinnakkaistallennetun version verkko-osoite", "ORCID", "Julkaisun tila" };



            // Ensimmäisen sarakkeen nimi (organisaation koodi) saa olla joku näistä kolmesta
            String[] OrganisaatioSarake = {"Ammattikorkeakoulu","Organisaatiotunnus","Yliopisto"};

            String[] CSVSarakkeet = null;
            String[] CSVSarakkeet_utf_vaarin = null;


            if (getVuosi() == 2016 || getVuosi() == 2017 || getVuosi() == 2018)
            {   
                // Jos amk
                if ( OnkoAMK(getOrganisaatioKoodi()) ) {

                    CSVSarakkeet = CSVSarakkeet2016_amk;
                    CSVSarakkeet_utf_vaarin = CSVSarakkeet2016_utf_vaarin_amk;
                }
                // Jos YO tai Sairaala - Tutkimuslaitos
                if (OnkoSairaalaTutkimus(getOrganisaatioKoodi()) || OnkoYO(getOrganisaatioKoodi()))
	            {
                    CSVSarakkeet = CSVSarakkeet2016;
                    CSVSarakkeet_utf_vaarin = CSVSarakkeet2016_utf_vaarin_yo;
                }    
                
            }
            else
            {   
                // Jos amk             
                if ( OnkoAMK(getOrganisaatioKoodi()) ) {

                    CSVSarakkeet = CSVSarakkeet2015_amk;
                    CSVSarakkeet_utf_vaarin = CSVSarakkeet2015_utf_vaarin_amk;
                }
                // Jos YO tai Sairaala - Tutkimuslaitos
                if (OnkoSairaalaTutkimus(getOrganisaatioKoodi()) || OnkoYO(getOrganisaatioKoodi()))
                {
                    CSVSarakkeet = CSVSarakkeet2015;
                    CSVSarakkeet_utf_vaarin = CSVSarakkeet2015_utf_vaarin_yo;
                }
            }

			string virheelliset = "";  // Virheilmoitustekstiin
            string virheSarakkeet = "";

            string virheelliset_utf8 = "";
            string virheelliset_utf8Sarakkeet = "";

			int indeksi = 0;  // sarakeindeksi

            if (CSVSarakkeet != null)	{
            
                    foreach (String sarake in sarakkeetCSV)  {
                            
					        // Jos sarake löytyy mutta väärässä paikassa, tai muutettu nimeä?  
                            if ( !CSVSarakkeet[indeksi].Trim().Equals(sarake.Trim()) )  
                            {                               
                                if ((indeksi == 0) && ( OrganisaatioSarake[0].Trim().Equals(sarake.Trim()) || OrganisaatioSarake[1].Trim().Equals(sarake.Trim()) || OrganisaatioSarake[2].Trim().Equals(sarake.Trim())   ))
                                {
                                    virheelliset = virheelliset + sarake + ", ";
                                    virheSarakkeet = virheSarakkeet + "väärä:" + sarake + " " + "oikea arvo [" + indeksi + "]: " + OrganisaatioSarake[0] + ", " + OrganisaatioSarake[1] + " tai  " + OrganisaatioSarake[2] + ", \r\n";

                                    MessageBox.Show(sarake.ToString() + " indeksi: " + indeksi.ToString(), "csv sarake väärässä kohdassa!");
                                }
                                else
                                {                     
                                    virheelliset = virheelliset + sarake + ", ";
                                    virheSarakkeet = virheSarakkeet + "väärä:" + sarake + " " + "oikea["+indeksi +"]: " + CSVSarakkeet[indeksi] +",\r\n";                                    
                                }
                            }


                            // Jos taas löytyy oikeasta sarakkeesta mutta on väärän merkistön mukainen sarakenimi
                            if (CSVSarakkeet_utf_vaarin[indeksi].Trim().Equals(sarake.Trim()))
                            {
                                if ((indeksi >= 15 && indeksi <= 37))
                                {                                   
                                    virheelliset_utf8 = virheelliset_utf8 + sarake + ", ";
                                    virheelliset_utf8Sarakkeet = virheelliset_utf8Sarakkeet + " " + CSVSarakkeet_utf_vaarin[indeksi] + ",\r\n";                                    
                                }
                            }	

					        indeksi++;                    
                    }
            }
            else
            {
                virheelliset =  virheelliset + "Väärä organisaatiokoodi";

                virheSarakkeet = "Väärä organisaatiokoodi: " + getOrganisaatioKoodi();

                string lyhyt_str = "";

                if ( OrganisaatiokoodinPituus(getOrganisaatioKoodi()) < 5  )
                {
                    lyhyt_str = " etunolla puuttuu?";
                }

                virheSarakkeet = virheSarakkeet + lyhyt_str;

                MessageBox.Show(virheSarakkeet, "Väärä organisaatiokoodi!");

                if (LokiOlemassa(masterloki))
                {
                    KirjoitaLokiin(virheSarakkeet, masterloki);
                }

            }

            if (virheelliset.Length > 0)
            {
                virheIlmoitus2();

                errorTextBox.AppendText(virheSarakkeet);

                if (virheelliset_utf8Sarakkeet.Length > 0)
                {
                    virheSarakkeet = virheSarakkeet + " Tarkista sarakenimet jotka eivät ole UTF8 merkistöä:\r\n\r\n" + virheelliset_utf8Sarakkeet;

                }

                if (LokiOlemassa(masterloki))
                {
                    KirjoitaLokiin(virheSarakkeet, masterloki);
                }



                return false;

            }
            else
            {
                return true;			

            }    
            

        }


        // Käännetty merkkijono ISBN tarkistusta varten
        private static string ReverseStr(string s)
        {
            string temp_str = s.Trim();

            char[] jono = temp_str.ToCharArray();
            Array.Reverse(jono);

            string jono_output = new string(jono);

            return jono_output;
        }

        // ISBN tarkistemerkki merkkijonon mukaan
        private static string ISBN_tarkistemerkki(string isbn_str)
        {
            string isbn = isbn_str.Trim();
            string tarkistusmerkki = "";           

            int checknum = 100;

            if (isbn.Length > 0)
            {
                string pattern = @"/[- ]|^(ISBN(?:-1[03])?:?)";  // poistaa isbn(10/13) veke  

                isbn = Regex.Replace(isbn, pattern, "");
                isbn = Regex.Replace(isbn, " ", ""); // Ei välilyöntejä
                isbn = Regex.Replace(isbn.Trim(), "-", ""); // Eikä "-" merkkejä

                int summa = 0;
                int i = 0;               

                // ISBN 13
                if (isbn.Length == 13)
                {
                    while (i < 12)
                    {
                        summa = summa + ((i % 2) * 2 + 1) * (Int32.Parse(isbn.Substring(i, 1)));
                        i++;
                    }
                    checknum = (10 - (summa % 10));
                    tarkistusmerkki = checknum.ToString();

                    if (checknum == 10)
                    {
                        tarkistusmerkki = "0";
                    }

                }

                // ISBN 10 
                if (isbn.Length == 10)
                {
                    isbn = isbn.Substring(0, 9);
                    string rev_str = ReverseStr(isbn);

                    i = 0;
                    summa = 0;

                    while (i < 9)
                    {
                        summa = summa + (i + 2) * Int32.Parse(rev_str.Substring(i, 1));
                        i++;
                    }

                    checknum = (11 - (summa % 11));
                    tarkistusmerkki = checknum.ToString();

                    if (checknum == 10)
                    {
                        tarkistusmerkki = "X";
                    }

                    if (checknum == 11)
                    {
                        tarkistusmerkki = "0";
                    }
                }		
                // .... ISBN 10 

            }


            return tarkistusmerkki;

        }


        // Tarkistetaan ISBN
        private bool TarkistaISBN(string isbn_str)
        {
            string isbn = isbn_str.Trim();

            string pattern = @"^(?:ISBN(?:-1[03])?:? )?(?=[0-9X]{10}$|(?=(?:[0-9]+[- ]){3})[- 0-9X]{13}$|97[89][0-9]{10}$|(?=(?:[0-9]+[- ]){4})[- 0-9]{17}$)(?:97[89][- ]?)?[0-9]{1,5}[- ]?[0-9]+[- ]?[0-9]+[- ]?[0-9X]$";
            
            bool val = false;

            if ( isbn.Length > 0) {

                // Teknisesti validi ISBN
                val = Regex.IsMatch(isbn, pattern,RegexOptions.IgnoreCase);

                if (val)
                {
                    // Tarkistetaan vielä viimeinen, eli tarkistusmerkki...

                    pattern = @"/[- ]|^(ISBN(?:-1[03])?:?)";  // poista isbn(10/13) veke  

                    isbn = Regex.Replace( isbn, pattern,"");
                    isbn = Regex.Replace(isbn, " ", ""); // Ei välilyöntejä
                    isbn = Regex.Replace(isbn.Trim(), "-", ""); // Ei "-" merkkejä


                    int summa = 0;
                    int i = 0;
                    int tarkistus;

                    // char tarkistusmerkki;
                    string check_num;

                    check_num = isbn.Substring(isbn.Length - 1); // viimeinen merkki

                    int check_numero = Int32.Parse(check_num);

                    // ISBN 13
                    if ( isbn.Length == 13 )
                    {
                        summa = 0;
                        i = 0;

                        while ( i < 12)
                        {
                            summa = summa + ((i%2)*2 + 1) * (Int32.Parse(isbn.Substring(i, 1)));
                            i++;
                        }
                        
                        int mod = (summa % 10);

                        tarkistus = 10 - (summa%10);

                        if (tarkistus == 10)
                        {
                            tarkistus = 0;
                        }

                        if ( tarkistus == check_numero )
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    } // ISBN 13 loppu


                    // ISBN 10 
                    if ( isbn.Length == 10 )
                    {
                        isbn = isbn.Substring(0, 9);

                        string rev_str = ReverseStr(isbn);

                        i = 0;
                        summa = 0;

                        while (i < 9)
                        {
                            summa = summa + (i + 2) * Int32.Parse(rev_str.Substring(i, 1));
                            i++;
                        }

                        tarkistus = (11 - (summa % 11));

                        string check = tarkistus.ToString();

                        if (tarkistus == 10)
                        {
                            check = "X";
                        }

                        if (tarkistus == 11)
                        {
                            check = "0";
                        }


                        // Tarkistusmerkki ok ?
                        if (check.Equals(check_num))
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }		
                    }  // ISBN 10 loppu

                    
                    return false;
                }
                else
                {
                    return false;

                }

            }
            
            
            return false;
            
        }


        // Tarkistetaan ISSN
        private static string ISSN_tarkistemerkki(string issn)
        {
            string issn_str = issn.Trim();
            string tarkistusmerkki = "";

            if (issn_str.Length > 0)
            {

                issn_str = Regex.Replace(issn_str, " ", ""); // Ei välilyöntejä
                issn_str = Regex.Replace(issn_str.Trim(), "-", ""); // Ei "-" merkkejä

                int summa = 0;
                int i = 0;
                int tarkistus = 0;
                string check_num = null;

                check_num = issn_str.Substring(issn_str.Length - 1); // viimeinen merkki

                if (issn_str.Length == 8)
                {
                    issn_str = issn_str.Substring(0, 7);

                    string rev_str = ReverseStr(issn_str);

                    while (i < 7)
                    {
                        summa = summa + (i + 2) * Int32.Parse(rev_str.Substring(i, 1));
                        i++;
                    }

                    tarkistus = (11 - (summa % 11));

                    tarkistusmerkki = tarkistus.ToString();

                    if (tarkistus == 10)
                    {
                        tarkistusmerkki = "X";
                    }

                    if (tarkistus == 11)
                    {
                        tarkistusmerkki = "0";
                    }
                }
            }

            return tarkistusmerkki;
        }


         // Tarkistetaan ISSN
        private bool TarkistaISSN(string issn)
        {
	        string pattern = @"^\d{4}-\d{3}[\dxX]$";

	        bool val = false;

	        string issn_str = issn.Trim();
	
	        if ( issn_str.Length > 0) {

		        // Teknisesti validi ISSN
		        val = Regex.IsMatch(issn_str, pattern, RegexOptions.IgnoreCase);

		        if (val)
		        {
			        // Tarkistetaan viimeinen eli tarkistusmerkki....
			
			        issn_str = Regex.Replace(issn_str, " ", ""); // Ei välilyöntejä
			        issn_str = Regex.Replace(issn_str.Trim(), "-", ""); // Ei "-" merkkejä

			        int summa = 0;
			        int i = 0;
                    int tarkistus = 0;
                    string tarkistusmerkki = null;
                    string check_num = null;

			        check_num = issn_str.Substring(issn_str.Length - 1); // viimeinen merkki
		
			        if ( issn_str.Length == 8 )
			        {
				        issn_str = issn_str.Substring(0,7);

				        string rev_str = ReverseStr(issn_str);

				        while ( i < 7)
				        {
					        summa = summa + (i+2) * Int32.Parse(rev_str.Substring(i, 1));
					        i++;
				        }

				        tarkistus = (11 - (summa%11));
                       
                        tarkistusmerkki = tarkistus.ToString();

                        if ( tarkistus == 10 )
				        {
					        tarkistusmerkki = "X";					
				        }
				
				        if ( tarkistus == 11 )
				        {
					        tarkistusmerkki = "0";
				        }
                        

				        // Tarkistusmerkki ok ?
                        if (check_num.Equals(tarkistusmerkki))
                        {					
					        return true;						
				        }
				        else	{
				
					        return false;
				        }			
			        }
			
			        return false;
		        }
		        else
		        {
			        return false;
		        }
	        }		
	        return false;
	
        }


        // Tarkistetaan ORCID
        private bool TarkistaORCID(string orcid_str)
        {
            string orcid = orcid_str.Trim();

            string pattern = @"^(\d{4})-(\d{4})-(\d{4})-(\d{3}[0-9X])$";

            bool val = false;

            if (orcid.Length == 0) return true;

            if (orcid.Length > 0)
            {
                // Teknisesti validi ORCID
                val = Regex.IsMatch(orcid, pattern, RegexOptions.IgnoreCase);

                if (val)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            return false;
        }


        // Poista viimeiset puolipisteet kentän loppupäästä
        private string PoistaViimeisetPuolipisteet(string arvo)
        {
            int poistaLen = 0;

            for (int i = arvo.Length - 1; i >= 0; i--)
            {
                char merkki = arvo[i];

                // Jos on puolipisteitä
                if ( merkki == ';' )    {
                        poistaLen++;
                }
                else                    {
                        break;
                }
            }
            

            if (poistaLen > 0)   {
                
                return arvo.Substring(0, arvo.Length - poistaLen);
            }

            return arvo;
        }



        // Tarkistetaan Orcidit
        private bool TarkistaORCIDArvot(string orcid_str_arvo)
        {
            string[] orcidit = orcid_str_arvo.Split(';');  // Tulee lista ; eroteltuja koodiarvoja

            bool arvot = false;
            string virheelliset = "";

            foreach (string orcid in orcidit)
            {
                if (TarkistaORCID(orcid))
                {
                    arvot = true;
                }
                else
                {
                    arvot = false;
                    virheelliset = virheelliset + orcid + " virheellinen. ";
                }
            }

            if (arvot)
            {
                return true;
            }
            else
            {
                MessageBox.Show(virheelliset,"ORCID virheellinen");
                return false;
            }
        }

        // Tarkistetaan DOI
        private bool TarkistaDOI(string doi_str)
        {
            string doi = doi_str.Trim();
		
            // Testattu eri säännöllisiä lausekkeita

			// string pattern = @"\b(10[.][0-9]{4,}(?:[.][0-9]+)*/(?:(?!['&\'<>])\S)+)\b";

            // string pattern = @"\b(10[.][0-9]{4,}(?:[.][0-9]+)*/(?:(?![%'#? ])\S)+)\b";
            
            string pattern = @"\b10\.(\d+\.*)+[\/](([^\s\.])+\.*)+\b";

			
            bool val = false;

            if (doi.Length > 0)
            {
                // Teknisesti validi DOI
                val = Regex.IsMatch(doi, pattern, RegexOptions.IgnoreCase);

                if (val)
                {
                    return true;
                }
                else
                {
                    return false;
                }               
            }

            return false;
          }

        // Tarkistetaan Pysyväosoite
        private bool TarkistaOsoite(string url_str)
        {
            string url = url_str.Trim();            
            
            string pattern_url = @"^(http(s)?://)?([\w-]+\.)+[\w-]+(/[\w- ;,./?%&=]*)?";

            bool val_http = false;

            // Tarpeeksi pitkä on x merkkiä?
            if (url.Length > 0 )
            {
                // Teknisesti validi http osoite tai urn
                val_http = Regex.IsMatch(url, pattern_url, RegexOptions.IgnoreCase);

                if (   val_http   )
                {
                    return true;
                }
                else
                {
                    return false;
                }                
            }
            return false;
        }
					

        // CSV ainestosta luetaan mikä ilmoitusvuosi riviltä 2 (indeksi 1), sarake 2 (indeksi 1), sekä organisaatiokoodi
        private void setVuosiFromCSV(List<string[]> parsedData)
        {
            LahdeDataGridView.Rows.Clear();
            LahdeDataGridView.Columns.Clear();
            LahdeDataGridView.Refresh();
            LahdeDataGridView.ColumnCount = 100;            

            for (int i = 0; i < 5; i++)
            {
                var sb = new StringBuilder(parsedData[0][i]);
                LahdeDataGridView.Columns[i].Name = sb.ToString();
            }                            

            foreach (string[] row in parsedData)
            {
                LahdeDataGridView.Rows.Add(row);
            }

            int temp_vuosiMuuttuja = 0;
            string s1;

            s1 = LahdeDataGridView.Rows[1].Cells[1].Value.ToString();  // Mikä vuosi?
            bool parsed = Int32.TryParse(s1, out temp_vuosiMuuttuja);
            if (parsed)
            {
                setVuosi(temp_vuosiMuuttuja);
            }
            else
            {
                setVuosi(2016); // default vuosi
            }

            string s2 = LahdeDataGridView.Rows[1].Cells[0].Value.ToString();  // Mikä organisaatiokoodi

            if ( s2.Trim().Length > 0)
            {
                setOrganisaatioKoodi(s2);
            }

            LahdeDataGridView.Rows.Clear();
            LahdeDataGridView.Columns.Clear();
            LahdeDataGridView.Refresh();

        }

        // XML elementtien mukaiset sarakkeet skeeman mukaisessa järjestyksessä. Pitää olla sama datagrid indeksointi ja elementtien järjestys XML muodostaessa
        private void DataGridXMLSarakkeet(DataGridView dgv)
        {
            DataGridViewTextBoxColumn col = null;

            // Ei autogeneroituja sarakenimiä
            dgv.AutoGenerateColumns = false;

            string[] elements = { "OrganisaatioTunnus", "IlmoitusVuosi", "JulkaisunTunnus", "JulkaisunTilaKoodi", "JulkaisunOrgTunnus", "YksikkoKoodi", "JulkaisuVuosi", "JulkaisunNimi", "TekijatiedotTeksti", "TekijoidenLkm", "SivunumeroTeksti", "Artikkelinumero", "AvainsanaTeksti", "ISBN", "JufoTunnus", "JufoLuokkaKoodi", "JulkaisumaaKoodi", "LehdenNimi", "ISSN", "VolyymiTeksti", "LehdenNumeroTeksti", "KonferenssinNimi", "KustantajanNimi", "KustannuspaikkaTeksti", "EmojulkaisunNimi", "EmojulkaisunToimittajatTeksti", "JulkaisutyyppiKoodi", "TieteenalaKoodi", "KoulutusalaKoodi", "YhteisjulkaisuKVKytkin", "YhteisjulkaisuSHPKytkin", "YhteisjulkaisuTutkimuslaitosKytkin", "YhteisjulkaisuMuuKytkin", "JulkaisunKansainvalisyysKytkin", "JulkaisunKieliKoodi", "AvoinSaatavuusKoodi", "EVOjulkaisuKytkin", "DOI", "PysyvaOsoiteTeksti", "LahdetietokannanTunnus", "Sukunimi", "Etunimet", "YksikkoKoodi2", "ORCID", "HankenumeroTeksti", "RahoittajaOrgTunnus" };

            string[] elements2015 = { "OrganisaatioTunnus", "IlmoitusVuosi", "JulkaisunTunnus", "JulkaisunTilaKoodi", "JulkaisunOrgTunnus", "YksikkoKoodi", "JulkaisuVuosi", "JulkaisunNimi", "TekijatiedotTeksti", "TekijoidenLkm", "SivunumeroTeksti", "Artikkelinumero", "AvainsanaTeksti", "ISBN", "JufoTunnus", "JufoLuokkaKoodi", "JulkaisumaaKoodi", "LehdenNimi", "ISSN", "VolyymiTeksti", "LehdenNumeroTeksti", "KonferenssinNimi", "KustantajanNimi", "KustannuspaikkaTeksti", "EmojulkaisunNimi", "EmojulkaisunToimittajatTeksti", "JulkaisutyyppiKoodi", "TieteenalaKoodi", "KoulutusalaKoodi", "YhteisjulkaisuKVKytkin", "YhteisjulkaisuSHPKytkin", "YhteisjulkaisuTutkimuslaitosKytkin", "YhteisjulkaisuMuuKytkin", "JulkaisunKansainvalisyysKytkin", "JulkaisunKieliKoodi", "AvoinSaatavuusKoodi", "EVOjulkaisuKytkin", "DOI", "PysyvaOsoiteTeksti", "LahdetietokannanTunnus", "Sukunimi", "Etunimet", "YksikkoKoodi2", "ORCID", "HankenumeroTeksti", "RahoittajaOrgTunnus" };

            string[] elements2016 = { "OrganisaatioTunnus", "IlmoitusVuosi", "JulkaisunTunnus", "JulkaisunTilaKoodi", "JulkaisunOrgTunnus", "YksikkoKoodi", "JulkaisuVuosi", "JulkaisunNimi", "TekijatiedotTeksti", "TekijoidenLkm", "SivunumeroTeksti", "Artikkelinumero", "AvainsanaTeksti", "ISBN", "JufoTunnus", "JufoLuokkaKoodi", "JulkaisumaaKoodi", "LehdenNimi", "ISSN", "VolyymiTeksti", "LehdenNumeroTeksti", "KonferenssinNimi", "KustantajanNimi", "KustannuspaikkaTeksti", "EmojulkaisunNimi", "EmojulkaisunToimittajatTeksti", "JulkaisutyyppiKoodi", "TieteenalaKoodi", "KoulutusalaKoodi", "YhteisjulkaisuKVKytkin", "JulkaisunKansainvalisyysKytkin", "JulkaisunKieliKoodi", "AvoinSaatavuusKoodi", "YhteisjulkaisuYritysKytkin", "RinnakkaistallennettuKytkin", "RinnakkaistallennusOsoiteTeksti", "DOI", "PysyvaOsoiteTeksti", "LahdetietokannanTunnus", "Sukunimi", "Etunimet", "YksikkoKoodi2", "ORCID" };

            if (getVuosi() == 2015)
            {
                elements = elements2015;
            }

            if (getVuosi() == 2016 || getVuosi() == 2017 || getVuosi() == 2018)
            {
                elements = elements2016;
            }

            try
            { 
                foreach (string arvo in elements)   {
            
                    col = new DataGridViewTextBoxColumn();
                    col.DataPropertyName = arvo;
                    col.Name = arvo;
                    col.HeaderText = arvo;
                    col.Visible = true;
                    dgv.Columns.Add(col);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Virhe datagrid taulukossa: " + ex.Message);
            }
        }


        // Onko sallittu Julkaisutyyppi koodi 
        private Boolean CheckJulkaisuTyyppiKoodi(string arvo)
        {
            string[] JulkaisuTyyppiKoodit = { "A1", "A2", "A3", "A4", "B1", "B2", "B3", "C1", "C2", "D1", "D2", "D3", "D4", "D5", "D6", "E1", "E2", "E3", "G4", "G5" };
            
            if ( Array.Exists(JulkaisuTyyppiKoodit, element => element == arvo) ) 
            {
                return true;
            }
            return false;
        }

        // Onko sallittu Tieteenala koodi 
        private Boolean CheckTieteenAlaKoodi(string koodit)
        {
            string[] TieteenAlaKoodit = {   "111","112","113","114",
                                            "115","116","1171","1172",
                                            "1181","1182","1183","1184",
                                            "119","211","212","213",
                                            "214","215","216","217",
                                            "218","219","220","221",
                                            "222","3111","3112","3121",
                                            "3122","3123","3124","3125",
                                            "3126","313","3141","3142",
                                            "315","316","317","318",
                                            "319","4111","4112","412",
                                            "413","414","415","511",
                                            "512","513","5141","5142",
                                            "515","516","517","518",
                                            "519","520","611","6121",
                                            "6122","6131","6132","614",
                                            "615","616","" };


            string[] tieteenalat = koodit.Split(';');

            string virheelliset = "";

            foreach (string tieteenala in tieteenalat)
            {               

                if ( !Array.Exists(TieteenAlaKoodit, element => element == tieteenala.Trim() ))
                {
                    virheelliset = virheelliset + tieteenala + ", ";                                                                               
                }               
            }

            if ( virheelliset.Trim().Length > 0 )
            {                
                return false;
            }
            else
            {
                return true;
            }      
                   
        }

         // Jos on liian lyhyt puuttuukohan ehkä etunolla?
        private int OrganisaatiokoodinPituus(string organisaatiokoodi)
        {   
            return organisaatiokoodi.Trim().Length;
        }

        // Onko Ammattikorkeakoulu?
        private Boolean OnkoAMK( string organisaatiokoodi )
        {
            string[] AMKt = { "02535", "02536", "02623", "10056", "02467", "02631", "02504", "02473", "02469", "02608", "02470", "10108", "02629", "02506", "10065", "10066", "02471", "02609", "02507", "02537", "02472", "02630", "02509", "02627", "02557", "10118" };

            if (Array.Exists(AMKt, element => element == organisaatiokoodi))
            {                
                return true;
            }
            return false;
        }

        // Onko Yliopisto?
        private Boolean OnkoYO( string organisaatiokoodi )
        {
            string[] YOt = { "01903", "10076", "01901", "10088", "01906", "01918", "01914", "01904", "01910", "10103", "01905", "01915", "10089", "01913", "02358" };

            if ( Array.Exists(YOt, element => element == organisaatiokoodi) )
            {                
                return true;
            }
            return false;
        }

        // Onko Sairaala tai Tutkimuslaitos - tuntematon 99999
        private Boolean OnkoSairaalaTutkimus(string organisaatiokoodi)
        {
            string[] organisaatiot = {	"4080015","5040011","4940015","4100010","4020217","7020017",
								        "5550012","26473754","5610017","02022481"
                                        ,"02202669"
                                        ,"1120017"
                                        ,"3060016"
                                        ,"02022481"
                                        ,"558005"
                                        ,"411001"
                                        ,"404001"
                                        ,"15675350"
                                        ,"08282559"
                                        ,"06794809"
                                        ,"08265978"
                                        ,"01714953"
                                        ,"99999"};

            if (Array.Exists(organisaatiot, element => element == organisaatiokoodi))
            {	
                return true;
            }
            return false;
        }
        

        // Onko sallittu Koulutuskoodi AMK ja YO eri koodisto, ei pakollinen, voi olla tyhjä
        private Boolean CheckKoulutuskoodi( string koodit, string organisaatio )
        {           
            string[] KoulutusKooditAMK = { "0", "1", "2", "3", "4", "5", "6", "7", "8","" };
            string[] KoulutusKooditYO =  { "75", "76", "77", "78", "79", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "90", "91", "92", "93", "94","" };

            string[] koulutusalat = koodit.Trim().Split(';');

            string virheelliset = "";

            if ( OnkoAMK(organisaatio) )
            {
                foreach (string koulutuskoodi in koulutusalat) 
                { 
                    if (!Array.Exists(KoulutusKooditAMK, element => element == koulutuskoodi))
                    {
                        virheelliset = virheelliset + koulutuskoodi + ", ";                       
                    }
                }
            }

            if ( OnkoYO(organisaatio) )
            {
                foreach (string koulutuskoodi in koulutusalat) 
                { 
                    if ( !Array.Exists(KoulutusKooditYO, element => element == koulutuskoodi) )
                    {
                        virheelliset = virheelliset + koulutuskoodi + ", ";                                          
                    }
                }
            }


            if ( OnkoSairaalaTutkimus(organisaatio) )
            {           
                foreach (string koulutuskoodi in koulutusalat)
                {
                    if (!Array.Exists(KoulutusKooditYO, element => element == koulutuskoodi))
                    {
                        virheelliset = virheelliset + "Koulutuskoodi ei ole pakollinen tieto sairaaloille tai tutktimuslaitoksille! ";
                        virheelliset = virheelliset + koulutuskoodi + ", ";
                    }
                }
            }



            if (virheelliset.Trim().Length > 0)
            {
                if ( !OnkoSairaalaTutkimus(organisaatio) ) { 
                    
                    MessageBox.Show(virheelliset.ToString(), "Koulutusalakoodit virheelliset!");
                }

                return false;
            }
            else
            {
                return true;
            }     		 
            
        }


        // Onko sallittu maa koodi 
        private Boolean CheckMaaKoodi(string arvo)
        {
            string[] MaaKoodit = {	    "004", 
					                    "008", 
					                    "010", 
					                    "012", 
					                    "016", 
					                    "020", 
					                    "024", 
					                    "028", 
					                    "031", 
					                    "032", 
					                    "036", 
					                    "040", 
					                    "044", 
					                    "048", 
					                    "050", 
					                    "051", 
					                    "052", 
					                    "056", 
					                    "060", 
					                    "064", 
					                    "068", 
					                    "070", 
					                    "072", 
					                    "074", 
					                    "076", 
					                    "084", 
					                    "086", 
					                    "090", 
					                    "092", 
					                    "096", 
					                    "100", 
					                    "104", 
					                    "108", 
					                    "112", 
					                    "116", 
					                    "120", 
					                    "124", 
					                    "132", 
					                    "136", 
					                    "140", 
					                    "144", 
					                    "148", 
					                    "152", 
					                    "156", 
					                    "158", 
					                    "162", 
					                    "166", 
					                    "170", 
					                    "174", 
					                    "175", 
					                    "178", 
					                    "180", 
					                    "184", 
					                    "188", 
					                    "191", 
					                    "192", 
					                    "196", 
					                    "203", 
					                    "204", 
					                    "208", 
					                    "212", 
					                    "214", 
					                    "218", 
					                    "222", 
					                    "226", 
					                    "231", 
					                    "232", 
					                    "233", 
					                    "234", 
					                    "238", 
					                    "239", 
					                    "242", 
					                    "246", 
					                    "248", 
					                    "250", 
					                    "254", 
					                    "258", 
					                    "260", 
					                    "262", 
					                    "266", 
					                    "268", 
					                    "270", 
					                    "275", 
					                    "276", 
					                    "288", 
					                    "292", 
					                    "296", 
					                    "300", 
					                    "304", 
					                    "308", 
					                    "312", 
					                    "316", 
					                    "320", 
					                    "324", 
					                    "328", 
					                    "332", 
					                    "334", 
					                    "336", 
					                    "340", 
					                    "344", 
					                    "348", 
					                    "352", 
					                    "356", 
					                    "360", 
					                    "364", 
					                    "368", 
					                    "372", 
					                    "376", 
					                    "380", 
					                    "384", 
					                    "388", 
					                    "392", 
					                    "398", 
					                    "400", 
					                    "404", 
					                    "408", 
					                    "410", 
					                    "414", 
					                    "417", 
					                    "418", 
					                    "422", 
					                    "426", 
					                    "428", 
					                    "430", 
					                    "434", 
					                    "438", 
					                    "440", 
					                    "442", 
					                    "446", 
					                    "450", 
					                    "454", 
					                    "458", 
					                    "462", 
					                    "466", 
					                    "470", 
					                    "474", 
					                    "478", 
					                    "480", 
					                    "484", 
					                    "492", 
					                    "496", 
					                    "498", 
					                    "499", 
					                    "500", 
					                    "504", 
					                    "508", 
					                    "512", 
					                    "516", 
					                    "520", 
					                    "524", 
					                    "528", 
					                    "530", 
					                    "531", 
					                    "533", 
					                    "534", 
					                    "535", 
					                    "540", 
					                    "548", 
					                    "554", 
					                    "558", 
					                    "562", 
					                    "566", 
					                    "570", 
					                    "574", 
					                    "578", 
					                    "580", 
					                    "581", 
					                    "583", 
					                    "584", 
					                    "585", 
					                    "586", 
					                    "591", 
					                    "598", 
					                    "600", 
					                    "604", 
					                    "608", 
					                    "612", 
					                    "616", 
					                    "620", 
					                    "624", 
					                    "626", 
					                    "630", 
					                    "634", 
					                    "638", 
					                    "642", 
					                    "643", 
					                    "646", 
					                    "652", 
					                    "654", 
					                    "659", 
					                    "660", 
					                    "662", 
					                    "663", 
					                    "666", 
					                    "670", 
					                    "674", 
					                    "678", 
					                    "682", 
					                    "686", 
					                    "688", 
					                    "690", 
					                    "694", 
					                    "702", 
					                    "703", 
					                    "704", 
					                    "705", 
					                    "706", 
					                    "710", 
					                    "716", 
					                    "724", 
					                    "728", 
					                    "729", 
					                    "732", 
					                    "736", 
					                    "740", 
					                    "744", 
					                    "748", 
					                    "752", 
					                    "756", 
					                    "760", 
					                    "762", 
					                    "764", 
					                    "768", 
					                    "772", 
					                    "776", 
					                    "780", 
					                    "784", 
					                    "788", 
					                    "792", 
					                    "795", 
					                    "796", 
					                    "798", 
					                    "800", 
					                    "804", 
					                    "807", 
					                    "810", 
					                    "818", 
					                    "826", 
					                    "831", 
					                    "832", 
					                    "833", 
					                    "834", 
					                    "840", 
					                    "850", 
					                    "854", 
					                    "858", 
					                    "860", 
					                    "862", 
					                    "876", 
					                    "882", 
					                    "887", 
					                    "891", 
					                    "894", 
					                    "999",
					                    ""	};


            if (Array.Exists(MaaKoodit, element => element == arvo))
            {
                return true;
            }
            
            return false;
        }
        

        // Onko sallittu kielikoodi 
        private Boolean CheckKieliKoodi(string arvo)
        {
            string[] KieliKoodit =	{   "20", 
							            "98", 
							            "aa", 
							            "ab", 
							            "ae", 
							            "af", 
							            "ak", 
							            "am", 
							            "an", 
							            "ar", 
							            "as", 
							            "av", 
							            "ay", 
							            "az", 
							            "ba", 
							            "be", 
							            "bg", 
							            "bh", 
							            "bi", 
							            "bm", 
							            "bn", 
							            "bo", 
							            "br", 
							            "bs", 
							            "ca", 
							            "ce", 
							            "ch", 
							            "co", 
							            "cr", 
							            "cs", 
							            "cu", 
							            "cv", 
							            "cy", 
							            "da", 
							            "de", 
							            "dv", 
							            "dz", 
							            "ee", 
							            "el", 
							            "en", 
							            "eo", 
							            "es", 
							            "et", 
							            "eu", 
							            "fa", 
							            "ff", 
							            "fi", 
							            "fj", 
							            "fo", 
							            "fr", 
							            "fy", 
							            "ga", 
							            "gd", 
							            "gl", 
							            "gn", 
							            "gu", 
							            "gv", 
							            "ha", 
							            "he", 
							            "hi", 
							            "ho", 
							            "hr", 
							            "ht", 
							            "hu", 
							            "hy", 
							            "hz", 
							            "ia", 
							            "id", 
							            "ie", 
							            "ig", 
							            "ii", 
							            "ik", 
							            "io", 
							            "is", 
							            "it", 
							            "iu", 
							            "ja", 
							            "jv", 
							            "ka", 
							            "kg", 
							            "ki", 
							            "kj", 
							            "kk", 
							            "kl", 
							            "km", 
							            "kn", 
							            "ko", 
							            "kr", 
							            "ks", 
							            "ku", 
							            "kv", 
							            "kw", 
							            "ky", 
							            "la", 
							            "lb", 
							            "lg", 
							            "li", 
							            "ln", 
							            "lo", 
							            "lt", 
							            "lu", 
							            "lv", 
							            "mg", 
							            "mh", 
							            "mi", 
							            "mk", 
							            "ml", 
							            "mn", 
							            "mo", 
							            "mr", 
							            "ms", 
							            "mt", 
							            "my", 
							            "na", 
							            "nb", 
							            "nd", 
							            "ne", 
							            "ng", 
							            "nl", 
							            "nn", 
							            "no", 
							            "nr", 
							            "nv", 
							            "ny", 
							            "oc", 
							            "oj", 
							            "om", 
							            "or", 
							            "os", 
							            "pa", 
							            "pi", 
							            "pl", 
							            "ps", 
							            "pt", 
							            "qu", 
							            "rm", 
							            "rn", 
							            "ro", 
							            "ru", 
							            "rw", 
							            "sa", 
							            "sc", 
							            "sd", 
							            "se", 
							            "sg", 
							            "sh", 
							            "si", 
							            "sk", 
							            "sl", 
							            "sm", 
							            "sn", 
							            "so", 
							            "sq", 
							            "sr", 
							            "ss", 
							            "st", 
							            "su", 
							            "sv", 
							            "sw", 
							            "ta", 
							            "te", 
							            "tg", 
							            "th", 
							            "ti", 
							            "tk", 
							            "tl", 
							            "tn", 
							            "to", 
							            "tr", 
							            "ts", 
							            "tt", 
							            "tw", 
							            "ty", 
							            "ug", 
							            "uk", 
							            "ur", 
							            "uz", 
							            "wa", 
							            "ve", 
							            "vi", 
							            "vo", 
							            "wo", 
							            "xh", 
							            "yi", 
							            "yo", 
							            "za", 
							            "zh", 
							            "zu" 
							             };



            if (Array.Exists(KieliKoodit, element => element == arvo))
            {       
                return true;
            }

            return false;
        }
        

        // XML elementtien mukaiset sarakkeet indeksin mukaan, skeeman mukaisessa järjestyksessä
        private string PalautaXMLElementtiNimi(int numero)
        {
            string[] elements = { "OrganisaatioTunnus", "IlmoitusVuosi", "JulkaisunTunnus", "JulkaisunTilaKoodi", "JulkaisunOrgTunnus", "YksikkoKoodi", "JulkaisuVuosi", "JulkaisunNimi", "TekijatiedotTeksti", "TekijoidenLkm", "SivunumeroTeksti", "Artikkelinumero", "AvainsanaTeksti", "ISBN", "JufoTunnus", "JufoLuokkaKoodi", "JulkaisumaaKoodi", "LehdenNimi", "ISSN", "VolyymiTeksti", "LehdenNumeroTeksti", "KonferenssinNimi", "KustantajanNimi", "KustannuspaikkaTeksti", "EmojulkaisunNimi", "EmojulkaisunToimittajatTeksti", "JulkaisutyyppiKoodi", "TieteenalaKoodi", "KoulutusalaKoodi", "YhteisjulkaisuKVKytkin", "YhteisjulkaisuSHPKytkin", "YhteisjulkaisuTutkimuslaitosKytkin", "YhteisjulkaisuMuuKytkin", "JulkaisunKansainvalisyysKytkin", "JulkaisunKieliKoodi", "AvoinSaatavuusKoodi", "EVOjulkaisuKytkin", "DOI", "PysyvaOsoiteTeksti", "LahdetietokannanTunnus", "Sukunimi", "Etunimet", "YksikkoKoodi2", "ORCID", "HankenumeroTeksti", "RahoittajaOrgTunnus" };

            string[] elements2015 = { "OrganisaatioTunnus", "IlmoitusVuosi", "JulkaisunTunnus", "JulkaisunTilaKoodi", "JulkaisunOrgTunnus", "YksikkoKoodi", "JulkaisuVuosi", "JulkaisunNimi", "TekijatiedotTeksti", "TekijoidenLkm", "SivunumeroTeksti", "Artikkelinumero", "AvainsanaTeksti", "ISBN", "JufoTunnus", "JufoLuokkaKoodi", "JulkaisumaaKoodi", "LehdenNimi", "ISSN", "VolyymiTeksti", "LehdenNumeroTeksti", "KonferenssinNimi", "KustantajanNimi", "KustannuspaikkaTeksti", "EmojulkaisunNimi", "EmojulkaisunToimittajatTeksti", "JulkaisutyyppiKoodi", "TieteenalaKoodi", "KoulutusalaKoodi", "YhteisjulkaisuKVKytkin", "YhteisjulkaisuSHPKytkin", "YhteisjulkaisuTutkimuslaitosKytkin", "YhteisjulkaisuMuuKytkin", "JulkaisunKansainvalisyysKytkin", "JulkaisunKieliKoodi", "AvoinSaatavuusKoodi", "EVOjulkaisuKytkin", "DOI", "PysyvaOsoiteTeksti", "LahdetietokannanTunnus", "Sukunimi", "Etunimet", "YksikkoKoodi2", "ORCID", "HankenumeroTeksti", "RahoittajaOrgTunnus" };

            string[] elements2016 = { "OrganisaatioTunnus", "IlmoitusVuosi", "JulkaisunTunnus", "JulkaisunTilaKoodi", "JulkaisunOrgTunnus", "YksikkoKoodi", "JulkaisuVuosi", "JulkaisunNimi", "TekijatiedotTeksti", "TekijoidenLkm", "SivunumeroTeksti", "Artikkelinumero", "AvainsanaTeksti", "ISBN", "JufoTunnus", "JufoLuokkaKoodi", "JulkaisumaaKoodi", "LehdenNimi", "ISSN", "VolyymiTeksti", "LehdenNumeroTeksti", "KonferenssinNimi", "KustantajanNimi", "KustannuspaikkaTeksti", "EmojulkaisunNimi", "EmojulkaisunToimittajatTeksti", "JulkaisutyyppiKoodi", "TieteenalaKoodi", "KoulutusalaKoodi", "YhteisjulkaisuKVKytkin", "JulkaisunKansainvalisyysKytkin", "JulkaisunKieliKoodi", "AvoinSaatavuusKoodi", "YhteisjulkaisuYritysKytkin", "RinnakkaistallennettuKytkin", "RinnakkaistallennusOsoiteTeksti", "DOI", "PysyvaOsoiteTeksti", "LahdetietokannanTunnus", "Sukunimi", "Etunimet", "YksikkoKoodi2", "ORCID" };

            string[] elements2017 = { "OrganisaatioTunnus", "IlmoitusVuosi", "JulkaisunTunnus", "JulkaisunTilaKoodi", "JulkaisunOrgTunnus", "YksikkoKoodi", "JulkaisuVuosi", "JulkaisunNimi", "TekijatiedotTeksti", "TekijoidenLkm", "SivunumeroTeksti", "Artikkelinumero", "AvainsanaTeksti", "ISBN", "JufoTunnus", "JufoLuokkaKoodi", "JulkaisumaaKoodi", "LehdenNimi", "ISSN", "VolyymiTeksti", "LehdenNumeroTeksti", "KonferenssinNimi", "KustantajanNimi", "KustannuspaikkaTeksti", "EmojulkaisunNimi", "EmojulkaisunToimittajatTeksti", "JulkaisutyyppiKoodi", "TieteenalaKoodi", "KoulutusalaKoodi", "YhteisjulkaisuKVKytkin", "JulkaisunKansainvalisyysKytkin", "JulkaisunKieliKoodi", "AvoinSaatavuusKoodi", "YhteisjulkaisuYritysKytkin", "RinnakkaistallennettuKytkin", "RinnakkaistallennusOsoiteTeksti", "DOI", "PysyvaOsoiteTeksti", "LahdetietokannanTunnus", "Sukunimi", "Etunimet", "YksikkoKoodi2", "ORCID" };

            string[] elements2018 = { "OrganisaatioTunnus", "IlmoitusVuosi", "JulkaisunTunnus", "JulkaisunTilaKoodi", "JulkaisunOrgTunnus", "YksikkoKoodi", "JulkaisuVuosi", "JulkaisunNimi", "TekijatiedotTeksti", "TekijoidenLkm", "SivunumeroTeksti", "Artikkelinumero", "AvainsanaTeksti", "ISBN", "JufoTunnus", "JufoLuokkaKoodi", "JulkaisumaaKoodi", "LehdenNimi", "ISSN", "VolyymiTeksti", "LehdenNumeroTeksti", "KonferenssinNimi", "KustantajanNimi", "KustannuspaikkaTeksti", "EmojulkaisunNimi", "EmojulkaisunToimittajatTeksti", "JulkaisutyyppiKoodi", "TieteenalaKoodi", "KoulutusalaKoodi", "YhteisjulkaisuKVKytkin", "JulkaisunKansainvalisyysKytkin", "JulkaisunKieliKoodi", "AvoinSaatavuusKoodi", "YhteisjulkaisuYritysKytkin", "RinnakkaistallennettuKytkin", "RinnakkaistallennusOsoiteTeksti", "DOI", "PysyvaOsoiteTeksti", "LahdetietokannanTunnus", "Sukunimi", "Etunimet", "YksikkoKoodi2", "ORCID" };

            if (getVuosi() == 2015)
            {
                elements = elements2015;
            }

            if (getVuosi() == 2016 )
            {
                elements = elements2016;
            }

            if (getVuosi() == 2017 )
            {
                elements = elements2017;
            }

            if (getVuosi() == 2018)
            {
                elements = elements2018;
            } 

            return elements[numero];
        }

        // Valitaan ja kopioidaan yhdestä Datagridistä DataTableen toiseen Datagridiin ja Tableen 
        private DataTable SwapDataFromGrid(DataGridView dgv1, DataGridView dgv2, int vuosiV)
        {
            var dt = new DataTable("Alkuper");
            var julkaisuTable = new DataTable("Julkaisu");
            
            dt.Columns.Clear();
            julkaisuTable.Columns.Clear();

            try
            { 
                    // Datagrid sarakkeet
                    foreach (DataGridViewColumn column in dgv1.Columns)
                    {
                        string sarakeNimi = column.Name;

                        if (column.Visible)
                        {
                            dt.Columns.Add(sarakeNimi, typeof(String));  // Sarakkeet csv/xls
                        }
                    }
                       
                    // Kuinka monta saraketta... viedään datatableen
                    foreach (DataGridViewColumn column in dgv2.Columns)
                    {
                        string sarakeNimi = column.Name;

                        if (column.Visible)
                        {
                            julkaisuTable.Columns.Add(sarakeNimi, typeof(String));  // Sarakkeet
                        }
                    }

                    object[] cellValues1 = new object[dgv1.Columns.Count];  // CSV lähdetiedoston sarakkeet järjestyksessä gridissä
                    object[] cellValues2 = new object[dgv2.Columns.Count];  // XML elementit järjestyksessä gridissä

                    // Datagrid rivit
                    foreach (DataGridViewRow row in dgv1.Rows)
                    {
                        for (int i = 0; i < row.Cells.Count; i++)
                        {
                            cellValues1[i] = row.Cells[i].Value;

                            // Vanhempaa ainestoa varten
                            if (vuosiV == 2013)
                            {
                                // Mitä tapahtuu jos sarakkeet tulee eri järjestyksessä, indeksointi menee väärin
                                // Pitää siis ensin validoida csv-tiedosto (että sarakkeet oikeassa järjestyksessä) 

                                cellValues2[0] = cellValues1[0]; // korkeakoulu
                                cellValues2[1] = cellValues1[1]; // Vuosi
                                cellValues2[2] = cellValues1[45]; // Julkaisun id
                                cellValues2[6] = cellValues1[24]; // Julkaisuvuosi
                                cellValues2[7] = cellValues1[23]; // Julkaisun nimi
                                cellValues2[8] = cellValues1[17]; // Julkaisun tekijät
                                cellValues2[9] = cellValues1[18]; // Julkaisun tekijöiden lukumäärä
                                cellValues2[10] = cellValues1[27]; // Sivut
                                cellValues2[11] = cellValues1[28]; // Artikkelinumero
                                cellValues2[13] = cellValues1[32]; // ISBN
                                cellValues2[15] = cellValues1[44]; // Julkaisun julkaisufoorumiluokitus
                                cellValues2[16] = cellValues1[37]; // Julkaisumaa
                                cellValues2[17] = cellValues1[30]; // Lehden/sarjan nimi
                                cellValues2[18] = cellValues1[31]; // ISSN
                                cellValues2[19] = cellValues1[25]; // Volyymi
                                cellValues2[20] = cellValues1[26]; // Numero
                                cellValues2[22] = cellValues1[35]; // Kustantaja
                                cellValues2[23] = cellValues1[36]; // Julkaisun kustannuspaikka
                                cellValues2[24] = cellValues1[33]; // Emojulkaisun nimi
                                cellValues2[25] = cellValues1[34]; // Emojulkaisun toimittajat
                                cellValues2[26] = cellValues1[2]; // Julkaisutyyppi 
                                cellValues2[27] = cellValues1[3]; // I Tieteenala
                                cellValues2[28] = cellValues1[9]; // I Koulutusala
                                cellValues2[29] = cellValues1[19]; // Kansainvälinen yhteisjulkaisu
                                cellValues2[30] = cellValues1[20]; // Yliopistollinen sairaanhoitopiiri, yhteisjulkaisu
                                cellValues2[31] = cellValues1[21]; // Valtion sektoritutkimuslaitos, yhteisjulkaisu
                                cellValues2[32] = cellValues1[22]; // Muu kotimainen tutkimusorganisaatio, yhteisjulkaisu
                                cellValues2[33] = cellValues1[38]; // Julkaisun kansainvälisyys
                                cellValues2[34] = cellValues1[29]; // Julkaisun kieli
                                cellValues2[35] = cellValues1[41]; // Avoin saatavuus
                                cellValues2[36] = cellValues1[43]; // EVO-julkaisu
                                cellValues2[37] = cellValues1[39]; // DOI-tunniste
                                cellValues2[38] = cellValues1[40]; // Pysyvä verkko-osoite
                                cellValues2[39] = cellValues1[42]; // Lähdetietokannan koodi
                                cellValues2[40] = cellValues1[15]; // Organisaation tekijät
                                cellValues2[42] = cellValues1[16]; // Organisaation alayksikkö

                            }

                            // Vanhempaa csc-ainestoa on haastavaa saada validia XML:ää 2015-16 skeman mukaan...
                            if (vuosiV == 2014)
                            {
                                cellValues2[0] = cellValues1[0]; // korkeakoulu Organisaatio
                                cellValues2[1] = cellValues1[1]; // Ilmoitusvuosi
                                cellValues2[2] = ""; // cellValues1[64]; // Julkaisun id tunnus

                                cellValues2[4] = cellValues1[64]; // Julkaisun Organisaatiotunnus

                                cellValues2[5] = cellValues1[16]; // Organisaation alayksikkö

                                for (int k = 1; k < 20; k++)
                                {

                                    if (cellValues1[16 + k] != null)
                                    {
                                        cellValues2[5] = cellValues2[5] + ";" + cellValues1[16 + k];
                                    }

                                }

                                cellValues2[6] = cellValues1[43]; // Julkaisuvuosi
                                cellValues2[7] = cellValues1[42]; // Julkaisun nimi
                                cellValues2[8] = cellValues1[36]; // Julkaisun tekijät
                                cellValues2[9] = cellValues1[37]; // Julkaisun tekijöiden lukumäärä
                                cellValues2[10] = cellValues1[46]; // Sivut
                                cellValues2[11] = cellValues1[47]; // Artikkelinumero

                                cellValues2[13] = cellValues1[51]; // ISBN
                                cellValues2[14] = cellValues1[63]; // JufoTunnus - Julkaisun julkaisukanava JUFO ID
                                //cellValues2[15] = cellValues1[64]; // Julkaisun julkaisufoorumiluokitus ID!!
                                cellValues2[16] = cellValues1[56]; // Julkaisumaa
                                cellValues2[17] = cellValues1[49]; // Lehden/sarjan nimi
                                cellValues2[18] = cellValues1[50]; // ISSN
                                cellValues2[19] = cellValues1[44]; // Volyymi
                                cellValues2[20] = cellValues1[45]; // Numero

                                cellValues2[21] = cellValues1[65];  // Konferenssin vakiintunut nimi

                                cellValues2[22] = cellValues1[54]; // Kustantaja
                                cellValues2[23] = cellValues1[55]; // Julkaisun kustannuspaikka
                                cellValues2[24] = cellValues1[52]; // Emojulkaisun nimi
                                cellValues2[25] = cellValues1[53]; // Emojulkaisun toimittajat
                                cellValues2[26] = cellValues1[2]; // Julkaisutyyppi 
                        
                                cellValues2[27] = cellValues1[3]; // I -VI Tieteenala

                                for (int k = 1; k < 6; k++)
                                {
                                    if (cellValues1[3 + k] != null)
                                    {
                                        cellValues2[27] = cellValues2[27] + ";" + cellValues1[3 + k];
                                    }
                                }

                                cellValues2[28] = cellValues1[9]; // I - VI Koulutusala
                        
                                for (int k = 1; k < 6; k++)
                                {
                                    if (cellValues1[9 + k] != null)
                                    {
                                        cellValues2[28] = cellValues2[28] + ";" + cellValues1[9 + k];
                                    }
                                }                        
                        
                                cellValues2[29] = cellValues1[38]; // Kansainvälinen yhteisjulkaisu
                                cellValues2[30] = cellValues1[39]; // Yliopistollinen sairaanhoitopiiri, yhteisjulkaisu
                                cellValues2[31] = cellValues1[40]; // Valtion sektoritutkimuslaitos, yhteisjulkaisu
                                cellValues2[32] = cellValues1[41]; // Muu kotimainen tutkimusorganisaatio, yhteisjulkaisu
                                cellValues2[33] = cellValues1[57]; // Julkaisun kansainvälisyys
                                cellValues2[34] = cellValues1[48]; // Julkaisun kieli
                                cellValues2[35] = cellValues1[60]; // Avoin saatavuus
                                cellValues2[36] = cellValues1[62]; // EVO-julkaisu
                                cellValues2[37] = cellValues1[58]; // DOI-tunniste
                                cellValues2[38] = cellValues1[59]; // Pysyvä verkko-osoite
                                cellValues2[39] = cellValues1[61]; // Lähdetietokannan koodi
                                cellValues2[40] = cellValues1[15]; // Organisaation tekijät

                                cellValues2[42] = cellValues1[16]; // Organisaation alayksikkö
                        
                                for (int k = 1; k < 20; k++) {

                                    if ( cellValues1[16 + k] != null )
                                    { 
                                        cellValues2[42] = cellValues2[42] + ";" + cellValues1[16+k];
                                    }    

                                }

                                cellValues2[43] = cellValues1[65];  // Konferenssin vakiintunut nimi

                            }

                            // Ilmoitusvuosi 2015. Tärkeää että nämä mäppäykset menevät oikein!
                            if (vuosiV == 2015)
                            {

                                cellValues2[0] = cellValues1[0]; // korkeakoulu Organisaatio
                                cellValues2[1] = cellValues1[1]; // Ilmoitusvuosi
                                cellValues2[2] = ""; // cellValues1[63]; // Julkaisun id tunnus

                                cellValues2[3] = 2; // JulkaisunTilaKoodi, vakioarvo 2 2015 vuoden csv-mallitiedostoa käytettäessä

                                cellValues2[4] = cellValues1[63]; // Julkaisun Organisaatiotunnus

                                cellValues2[5] = cellValues1[16]; // Organisaation alayksikkö

                                for (int k = 1; k < 20; k++)
                                {
                                    if (cellValues1[16 + k] != null)
                                    {
                                        cellValues2[5] = cellValues2[5] + ";" + cellValues1[16 + k];
                                    }
                                }

                                cellValues2[6] = cellValues1[43]; // Julkaisuvuosi
                                cellValues2[7] = cellValues1[42]; // Julkaisun nimi
                                cellValues2[8] = cellValues1[36]; // Julkaisun tekijät
                                cellValues2[9] = cellValues1[37]; // Julkaisun tekijöiden lukumäärä
                                cellValues2[10] = cellValues1[46]; // Sivut
                                cellValues2[11] = cellValues1[47]; // Artikkelinumero

                                cellValues2[13] = cellValues1[51]; // ISBN
                                cellValues2[14] = cellValues1[62]; // JufoTunnus - Julkaisun julkaisukanava JUFO ID
                                //cellValues2[15] = cellValues1[64]; // Julkaisun julkaisufoorumiluokitus ID!!
                        
                                cellValues2[16] = cellValues1[56]; // Julkaisumaa
                                cellValues2[17] = cellValues1[49]; // Lehden/sarjan nimi
                                cellValues2[18] = cellValues1[50]; // ISSN
                                cellValues2[19] = cellValues1[44]; // Volyymi
                                cellValues2[20] = cellValues1[45]; // Numero

                                cellValues2[22] = cellValues1[54]; // Kustantaja
                                cellValues2[23] = cellValues1[55]; // Julkaisun kustannuspaikka
                                cellValues2[24] = cellValues1[52]; // Emojulkaisun nimi
                                cellValues2[25] = cellValues1[53]; // Emojulkaisun toimittajat
                                cellValues2[26] = cellValues1[2]; // Julkaisutyyppi 

                                cellValues2[27] = cellValues1[3]; // I -VI Tieteenala

                                for (int k = 1; k < 6; k++)
                                {
                                    if (cellValues1[3 + k] != null)
                                    {
                                        cellValues2[27] = cellValues2[27] + ";" + cellValues1[3 + k];
                                    }
                                }

                                cellValues2[28] = cellValues1[9]; // I - VI Koulutusala

                                for (int k = 1; k < 6; k++)
                                {
                                    if (cellValues1[9 + k] != null)
                                    {
                                        cellValues2[28] = cellValues2[28] + ";" + cellValues1[9 + k];
                                    }
                                }

                                cellValues2[29] = cellValues1[38]; // Kansainvälinen yhteisjulkaisu
                                cellValues2[30] = cellValues1[39]; // Yliopistollinen sairaanhoitopiiri, yhteisjulkaisu
                                cellValues2[31] = cellValues1[40]; // Valtion sektoritutkimuslaitos, yhteisjulkaisu
                                cellValues2[32] = cellValues1[41]; // Muu kotimainen tutkimusorganisaatio, yhteisjulkaisu
                                cellValues2[33] = cellValues1[57]; // Julkaisun kansainvälisyys
                                cellValues2[34] = cellValues1[48]; // Julkaisun kieli
                                cellValues2[35] = cellValues1[60]; // Avoin saatavuus
                                //cellValues2[36] = cellValues1[62]; // EVO-julkaisu   Poistettu käytöstä 2015
                                cellValues2[37] = cellValues1[58]; // DOI-tunniste
                                cellValues2[38] = cellValues1[59]; // Pysyvä verkko-osoite
                                cellValues2[39] = cellValues1[61]; // Lähdetietokannan koodi
                                cellValues2[40] = cellValues1[15]; // Organisaation tekijät

                                cellValues2[42] = cellValues1[16]; // Organisaation alayksikkö

                                for (int k = 1; k < 20; k++)
                                {
                                    if (cellValues1[16 + k] != null)
                                    {
                                        cellValues2[42] = cellValues2[42] + ";" + cellValues1[16 + k];
                                    }
                                }

                                cellValues2[12] = cellValues1[65]; // Avain sanat eli AvainsanaTeksti
                                cellValues2[21] = cellValues1[64];  // Konferenssin vakiintunut nimi                      

                            }



                            // 2017 raportointivuosi 2018 
                            // 2016 raportointivuosi 2017
                            // - 6 koska koulutus ala on poistettu
                            if (vuosiV == 2016 || vuosiV == 2017 || vuosiV == 2018)
                            {
                                cellValues2[0] = cellValues1[0];    // korkeakoulu Organisaatio
                                cellValues2[1] = cellValues1[1];    // Ilmoitusvuosi
                                cellValues2[2] = ""; // cellValues1[55];   // Julkaisun id tunnus

                                if (cellValues1[61] == null || cellValues1[61].ToString().Trim().Length == 0)
                                {
                                    cellValues2[3] = 2;                 // JulkaisunTilaKoodi, vakioarvo 2 2016 csv-mallitiedostoa käytettäessä
                                }
                                else
                                {
                                    cellValues2[3] = cellValues1[61];
                                }


                                cellValues2[4] = cellValues1[55];   // Julkaisun Organisaatiotunnus
                                cellValues2[5] = cellValues1[10];   // Organisaation alayksikkö

                                for (int k = 1; k < 20; k++)
                                {
                                    if (cellValues1[10 + k] != null)
                                    {
                                        cellValues2[5] = cellValues2[5] + ";" + cellValues1[10 + k];
                                    }
                                }

                                cellValues2[6] = cellValues1[35]; // Julkaisuvuosi
                                cellValues2[7] = cellValues1[34]; // Julkaisun nimi
                                cellValues2[8] = cellValues1[30]; // Julkaisun tekijät
                                cellValues2[9] = cellValues1[31]; // Julkaisun tekijöiden lukumäärä
                                cellValues2[10] = cellValues1[38]; // Sivut
                                cellValues2[11] = cellValues1[39]; // Artikkelinumero
                                cellValues2[12] = cellValues1[57]; // Avain sanat
                                cellValues2[13] = cellValues1[43]; // ISBN
                                cellValues2[14] = cellValues1[54]; // JufoTunnus - Julkaisun julkaisukanava JUFO ID
                                //cellValues2[15] = null; // Julkaisun julkaisufoorumiluokitus ID!!
                                cellValues2[16] = cellValues1[48]; // Julkaisumaa
                                cellValues2[17] = cellValues1[41]; // Lehden/sarjan nimi
                                cellValues2[18] = cellValues1[42]; // ISSN
                                cellValues2[19] = cellValues1[36]; // Volyymi
                                cellValues2[20] = cellValues1[37]; // Numero
                                cellValues2[21] = cellValues1[56];  // Konferenssin vakiintunut nimi
                                cellValues2[22] = cellValues1[46]; // Kustantaja
                                cellValues2[23] = cellValues1[47]; // Julkaisun kustannuspaikka
                                cellValues2[24] = cellValues1[44]; // Emojulkaisun nimi
                                cellValues2[25] = cellValues1[45]; // Emojulkaisun toimittajat
                                cellValues2[26] = cellValues1[2]; // Julkaisutyyppi 

                                cellValues2[27] = cellValues1[3]; // I -VI Tieteenala

                                for (int k = 1; k < 6; k++)
                                {
                                    if (cellValues1[3 + k] != null)
                                    {
                                        cellValues2[27] = cellValues2[27] + ";" + cellValues1[3 + k];
                                    }
                                }

                                //cellValues2[28] = cellValues1[9]; // I - VI Koulutusala

                                cellValues2[28] = null;

                                //for (int k = 1; k < 6; k++)
                                //{
                                //    if (cellValues1[9 + k] != null)
                                //    {
                                //        cellValues2[28] = cellValues2[28] + ";" + cellValues1[9 + k];
                                //    }
                                //}

                                cellValues2[29] = cellValues1[32]; // Kansainvälinen yhteisjulkaisu
                                cellValues2[30] = cellValues1[49]; // Julkaisun kansainvälisyys // ent Valtion sektoritutkimuslaitos, yhteisjulkaisu
                                cellValues2[31] = cellValues1[40]; // Julkaisun kieli // Ent Muu kotimainen tutkimusorganisaatio, yhteisjulkaisu
                                cellValues2[32] = cellValues1[52]; // Avoin saatavuus

                                cellValues2[33] = cellValues1[33]; // Yhteisjulkaisu yrityksen kanssa "YhteisjulkaisuYritysKytkin"

                                // Julkaisu rinnakkaistallennettu, Rinnakkaistallennetun version verkko-osoite
                                cellValues2[34] = cellValues1[58]; // Julkaisu rinnakkaistallennettu
                                cellValues2[35] = cellValues1[59]; // Rinnakkaistallennetun version verkko-osoite

                                cellValues2[36] = cellValues1[50]; // DOI-tunniste
                                cellValues2[37] = cellValues1[51]; // Pysyvä verkko-osoite
                                cellValues2[38] = cellValues1[53]; // Lähdetietokannan koodi

                                cellValues2[39] = cellValues1[9]; // Organisaation tekijät
                                cellValues2[40] = null;
                                cellValues2[41] = cellValues1[10]; // Organisaation alayksikkö

                                for (int k = 1; k < 20; k++)
                                {
                                    if (cellValues1[10 + k] != null)
                                    {
                                        cellValues2[41] = cellValues2[41] + ";" + cellValues1[10 + k];
                                    }
                                }
                                //  ORCID    
                                cellValues2[42] = cellValues1[60]; // ORCID                        

                            } // 2016 ja 2017 loppuu... Pysyköhän kasassa?



                        }

                        julkaisuTable.Rows.Add(cellValues2);

                    } // foreach Datagrid rivi loppu

                }
                catch (Exception ex)
	            {
		            MessageBox.Show("Virhe datagridissä: " + ex.Message);  
	            }

                return julkaisuTable; // Palautetaan käsitelty datataulu
        }

  
        // Käytetään VB Parseria CSV-datan lukemiseen koska meillä on "arvo;arvo" sarakkeita   
        private static List<string[]> LueCSVDataa(string tiedosto, char separator)
        {
            var csvData = new List<string[]>();

            try
            {
                TextFieldParser parser = new TextFieldParser(tiedosto);

                parser.HasFieldsEnclosedInQuotes = true;
                parser.SetDelimiters(";");

                string[] fields;

                while (!parser.EndOfData)
                {
                    fields = parser.ReadFields();
                    csvData.Add(fields);
                }
                parser.Close();               

            }
            catch (Exception e)
            {
                MessageBox.Show("Virhe csv-tiedostoa luettaessa: " + e.Message);              

            }

            return csvData;
        }


        // Näytetään csv-data Datagridissä
        private void DrawDataGridView(List<string[]> parsedData)
        {
            LahdeDataGridView.Rows.Clear();

            LahdeDataGridView.ColumnCount = 80; // Note to myself Dynaamiseksi kiitos!  2013: i < 46  2014 => i < 66, 2016, 2017  62 max

            int maximi = 0;

            if (getVuosi() == 2016 || getVuosi() == 2017 || getVuosi() == 2018)
            {
                maximi = 62;    // Voi että!
            }

            if ( getVuosi() == 2015 ) 
            {   
                maximi = 66;
            }

            if ( getVuosi() == 2014)
            {
                maximi = 66;
            }

            if (getVuosi() == 2013)
            {   
                maximi = 46;
            }

            try
            {
                // Tää on vain nice-to-have eli nähdä sarakkeiden nimet luetusta csv:stä
                for (int i = 0; i < maximi; i++)
                {
                    var sb = new StringBuilder(parsedData[0][i]);

                    sb.Replace("\"", "");  // Siivotaan mahdollisia "-merkkejä

                    LahdeDataGridView.Columns[i].Name = sb.ToString();
                }

                foreach (string[] row in parsedData)
                {
                    LahdeDataGridView.Rows.Add(row);
                }

                LahdeDataGridView.Rows.Remove(LahdeDataGridView.Rows[0]);

                int temp_vuosiMuuttuja = 0;
                string s1;

                s1 = LahdeDataGridView.Rows[1].Cells[1].Value.ToString();  // Mikä vuosi?

                bool parsed = Int32.TryParse(s1, out temp_vuosiMuuttuja);

                if (parsed)
                {
                    vuosiMuuttuja = temp_vuosiMuuttuja;
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("Virhe lähdeainestoa luettaessa: " + e.Message);
            }

        }

        // HaeExcelTyokirjat ei vielä käytössä, mutta aloitettu tulevaisuutta varten, seuraavaan versioon jos csv muodostaminen osoittautuu liian hankalaksi
        // Ja myös UTF8 merkistöjen kanssa haastavaa joku autoconvert
        // Metodi jolla luetaan työkirjan (worksheet) nimi -- tarvitaan jos luetaan OLEDB kautta Excelistä dataa
        public string[] HaeExcelTyokirjat(string excelTiedosto)
        {
            MessageBox.Show(excelTiedosto, "HaeExcelTyokirjat");

            //var currentDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //string tiedosto = Path.Combine(currentDirectory, Path.GetFileName(excelTiedosto));

            DataTable datataulu = null;

            OleDbConnection yhteys = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelTiedosto + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\";");

            yhteys.Open();

            datataulu = yhteys.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (datataulu == null)  { 
                return null;
            }

            yhteys.Close();

            String[] excelTyokirjat = new String[datataulu.Rows.Count];
            int i = 0;

            foreach (DataRow rivi in datataulu.Rows)
            {
                excelTyokirjat[i] = rivi["TABLE_NAME"].ToString();
                i++;
            }
            return excelTyokirjat;
        }





        // Tässä luetaan lähdetiedosto (csv) jostakin ja viedään sen tiedot datagridiin
        private void LueLahdeTiedostoButton_Click(object sender, EventArgs e)
        {
            if (LueLahdeTiedostoButton.Enabled) 
            {
                errorTextBox.Clear();

                LahdeDataGridView.Rows.Clear();
                LahdeDataGridView.Columns.Clear();
                LahdeDataGridView.Refresh();

                bool UTF8_merkisto = false;

                Stream myStream = null;
                OpenFileDialog avaaTiedosto = new OpenFileDialog();

                avaaTiedosto.InitialDirectory = Directory.GetCurrentDirectory();

                // avaaTiedosto.Filter = "Excel (*.xls)|*.xls|csv (*.csv)|*.csv"; // Excel vanhassa xls muodossa tai csv

                avaaTiedosto.Filter = "csv (*.csv)|*.csv"; // Vain csv muodossa 

                avaaTiedosto.FilterIndex = 1;
                avaaTiedosto.RestoreDirectory = true;

                if (avaaTiedosto.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if ((myStream = avaaTiedosto.OpenFile()) != null)
                        {                  
                            string extension;
                            extension = Path.GetExtension(avaaTiedosto.FileName);   // Tiedoston pääte. Onko .xls vai .csv ?               


                            // Tarkistetaan TextFileEncodingDetectori.DetectTextFileEncoding:lla jos tiedosto on UTF8  
                            try
                            {   
                                var encoding = TextFileEncodingDetector.DetectTextFileEncoding(avaaTiedosto.FileName);

                                String merkisto = encoding.ToString();

                                if (LokiOlemassa(masterloki))
                                {
                                    string error_str = merkisto;
                                    KirjoitaLokiin(avaaTiedosto.FileName + " tiedoston merkistö on: " + error_str, masterloki);
                                }                                

                                errorTextBox.AppendText(avaaTiedosto.FileName + " : " + merkisto);  
                            
                                if (merkisto.Trim().Equals("System.Text.UTF8Encoding")) {                                                                         
                                
                                    UTF8_merkisto = true;
                                }

                            }
                            catch (Exception exept)
                            {
                                MessageBox.Show("csv-lähdetiedosto ei ole UTF8!","Tiedostovirhe");

                                String loki_str = exept.ToString();
                                errorTextBox.AppendText("Tiedosto ei ole UTF8!");                               


                            }                      


                            if (UTF8_merkisto)
                            {
                                try
                                {
                                    // Jos Excel .xls tiedosto muoto
                                    if (extension.Contains("xls"))
                                    {
                                        LahdeDataGridView.DataSource = null;

                                        // Haetaan Excel worksheet names eli työkirjan/-kirjojen nimi
                                        string[] ExcelBlad = HaeExcelTyokirjat(avaaTiedosto.FileName);

                                        String Tyokirja = ExcelBlad[0];

                                        OleDbConnection yhteys = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + avaaTiedosto.FileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\";");
                                        OleDbDataAdapter dataAdapteri = new OleDbDataAdapter("select * from [" + Tyokirja + "]", yhteys);

                                        DataTable dataTaulu = new DataTable();
                                        dataAdapteri.Fill(dataTaulu);

                                        LahdeDataGridView.DataSource = dataTaulu;  // Ylin datagrid näyttää lähdetiedoston tiedot
                                        yhteys.Close();

                                        MuunnaXMLButton.Enabled = true;
                                    }


                                    // Jos csv tiedosto
                                    if (extension.Contains("csv"))
                                    {
                                        LahdeDataGridView.DataSource = null;
                                        List<string[]> csvData = null;

                                        csvData = LueCSVDataa(avaaTiedosto.FileName, ';');

                                        setVuosiFromCSV(csvData);

                                        // Tarkistetaan onko sarakkeiden määrä ok luetussa csv-tiedostossa
                                        int num = MontakoSaraketta(avaaTiedosto.FileName, ';');


                                        int rivia = MontakoRivia(avaaTiedosto.FileName);



                                        bool num_ok = false;

                                        // Muutettu  if (getVuosi() >= 2014) 1.3.2016

                                        //MessageBox.Show(vuosiIlmo.ToString(), "vuosiIlmo");

                                        if (getVuosi() == vuosiIlmo)
                                        {
                                            if (getVuosi() <= 2015 && num >= 66)
                                            {
                                                num_ok = true;
                                            }

                                            if ((getVuosi() == 2016 || getVuosi() == 2017 || getVuosi() == 2018) && num == 62)
                                            {
                                                num_ok = true;
                                            }

                                        }
                                        else
                                        {
                                            string infostr = "Väärä ilmoitusvuosi: " + getVuosi().ToString() + " Oikea ilmoitusvuosi: " + vuosiIlmo.ToString() + "! \r\n";
                                            MessageBox.Show(infostr, "Väärä ilmoitusvuosi");

                                            if (LokiOlemassa(masterloki))
                                            {                                               
                                                KirjoitaLokiin(infostr, masterloki);
                                            }


                                        }



                                        if (num_ok)
                                        {
                                            if (TarkistaSarakkeetCSVsta(csvData, num))
                                            {
                                                string infostr = "CSV-lähdetiedosto ok! \r\n\r\n" + rivia.ToString() + " riviä luettu!";

                                                MessageBox.Show(infostr, "CSV-tiedosto ok!");

                                                if (LokiOlemassa(masterloki))
                                                {
                                                    KirjoitaLokiin("CSV-lähdetiedosto: " + avaaTiedosto.FileName + " ok!", masterloki);
                                                    KirjoitaLokiin(infostr, masterloki);
                                                }



                                                DrawDataGridView(csvData);

                                                MuunnaXMLButton.Enabled = true;

                                                tiedostoNimiLabel.Text = "" + avaaTiedosto.FileName + "";
                                                setRowNumber(LahdeDataGridView);
                                            }
                                            else
                                            {
                                                MessageBox.Show("Virheellinen csv-lähdetiedosto! :( ","Virheilmoitus");

                                                if (LokiOlemassa(masterloki))
                                                {
                                                    KirjoitaLokiin("Virheellinen csv-lähdetiedosto: " + avaaTiedosto.FileName + " ei ole ok!", masterloki);
                                                }

                                            }
                                        }
                                        else
                                        {
                                            MuunnaXMLButton.Enabled = false;
                                            MessageBox.Show("Sarakkeiden määrä väärä: " + num.ToString(), "Virheellinen csv tiedosto!");

                                            errorTextBox.AppendText("Sarakemäärä väärä: " + num.ToString());

                                            if (LokiOlemassa(masterloki))
                                            {
                                                KirjoitaLokiin("Virheellinen csv-lähdetiedosto. Sarakkeiden määrä väärä: " + num.ToString(), masterloki);
                                            }


                                        }
                                    }
                                }
                                catch (Exception Args)
                                {
                                    MessageBox.Show("Harmi! Lähdetiedostoa ei voitu avata: " + Args.Message);
                                }

                            }

                            else
                            {

                                MessageBox.Show("CSV-tiedosto ei ole UTF-8!\r\nKatso Työkalut/Avaa lokitiedosto ja Ohje/Ohjeet", "Tiedostovirhe");

    
                                if (LokiOlemassa(masterloki))
                                {
                                    string error_str = avaaTiedosto.FileName + " csv-lähdetiedosto ei ole UTF8!" + "";

                                    KirjoitaLokiin(error_str, masterloki);
                                }
	


                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Virhe: Voi ei! Taas kaikki meni pieleen! Original error: " + ex.Message);
                    }
                }

            }
            else
            {

                MessageBox.Show("Uutta lähdetiedostoa ei voi avata!\n\r" + "Jos haluat aloittaa muunnoksen alusta valitse\n\rTyhjennä vanhat tiedot ja lue uusi lähdetiedosto tai paina CTRL+N", "Ilmoitus");

            }

        }

        // Viedän dataa toiseen datagridiin
        private void MuunnaXMLButton_Click(object sender, EventArgs e)
        {

            if ( MuunnaXMLButton.Enabled )
            { 
                    XMLdataGridView.Columns.Clear();

                    // Före detta populera columner Datagridview 2        
                    DataGridXMLSarakkeet(XMLdataGridView);

                    int vuosi = vuosiMuuttuja;  // Mikä vuosi, mikä valuutta...                        

                    DataTable dT = null;
                    dT = SwapDataFromGrid(LahdeDataGridView, XMLdataGridView, vuosi);

                    DataSet dS = new DataSet("Julkaisut");
                    dS.Tables.Add(dT);

                    XMLdataGridView.DataSource = dT;
                    XMLdataGridView.EnableHeadersVisualStyles = false;

                    MuunnaXMLButton.Enabled = false;

                    setRowNumber(XMLdataGridView);

                    ValidoiButton.Enabled = true;
                    LueLahdeTiedostoButton.Enabled = false;
            }
            else
            {
                MessageBox.Show("Lähdetiedon dataa ei voi tällä hetkellä muuntaa!");
            }

        }

        // Datagridistä dataTableen kaikki data josta sitten muodostetaan XML:ää
        private DataTable GetDataTableFromDGV(DataGridView dgv)
        {
            var dt = new DataTable("Julkaisu");

            dt.Columns.Clear();

            foreach (DataGridViewColumn column in dgv.Columns)
            {
                string sarakeNimi = column.Name;

                if (column.Visible)
                {
                    dt.Columns.Add(sarakeNimi, typeof(String));  // Sarakkeet
                }
            }

            object[] cellValues = new object[dgv.Columns.Count];
            foreach (DataGridViewRow row in dgv.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    cellValues[i] = row.Cells[i].Value;
                }
                dt.Rows.Add(cellValues);
            }
            return dt;
        }


        // Tässä luodaan itse xml pitkä, pitkä string...
        private string JulkaisutXML(DataTable taulukko)
        {       
                DataSet dS = new DataSet("Julkaisut");
                dS.Tables.Add(taulukko);            

                MemoryStream memStream = new MemoryStream();

                var settings = new XmlWriterSettings();
                settings.OmitXmlDeclaration = false;
                settings.Indent = true;
                // settings.NewLineOnAttributes = true;  // Ei uutta riviä jokaisen attribuutin kohdalla            
                settings.Encoding = Encoding.UTF8;

                try
                {       
                    XmlWriter writer = XmlWriter.Create(memStream, settings);            

                    writer.WriteStartDocument();

                    writer.WriteStartElement(null,"Julkaisut","urn:mace:funet.fi:julkaisut/2015/03/01"); // ok!            

                    writer.WriteAttributeString("xmlns","xsi",null, "http://www.w3.org/2001/XMLSchema-instance"); // ok!

                    writer.WriteAttributeString("xsi", "schemaLocation", null, "urn:mace:funet.fi:julkaisut/2015/03/01");

                    foreach (DataRow dataRow in taulukko.Rows)
                    {
                        writer.WriteStartElement("Julkaisu");

                        // Käydään taulukko läpi rivi riviltä per elementti
                        for (int j = 0; j < dataRow.ItemArray.Length; j++)
                        {
                            string elementti = PalautaXMLElementtiNimi(j); // Elementin nimi xml:ssä

                            Boolean validiElementti = true; // Tarvitaan jos lopussa kaikki ok tai false jos on jo käsitelty

                            if (dataRow.ItemArray[j] == DBNull.Value)
                                    dataRow.SetField(j, string.Empty);
                    
                            string arvo = dataRow.ItemArray[j].ToString().Trim(); // Elementin arvo                            
                            
                            if ( elementti.Equals("JulkaisunNimi") && arvo.Trim().Length == 0)
                            {
                                validiElementti = false;
                            }

                            // JulkaisunTilaKoodi
                            if (elementti.Equals("JulkaisunTilaKoodi") )
                            {
                                validiElementti = true;
                                // Default arvo = "2";
                            }


                            // JulkaisunTunnus
                            if (elementti.Equals("JulkaisunTunnus"))
                            {
                                validiElementti = false;
                            }


                            if (elementti.Equals("YksikkoKoodi2"))
                            {
                                validiElementti = false;       // elementti "YksikkoKoodi" tulee sisäelementtinä kahdessa kohtaa 
                            }                    


                            if (elementti.Equals("YksikkoKoodi"))
                            {
                                validiElementti = false; // Ei lopussa saa kirjoittaa uudelleen

                                writer.WriteStartElement("JulkaisunOrgYksikot"); // JulkaisunOrgYksikot alkaa                                                     

                                string[] JulkaisunOrgYksikotArray = arvo.Split(';'); // datassa ; eroteltu jos on enemmän

                                foreach (string orgyksikko in JulkaisunOrgYksikotArray)
                                {
                                    string orgYksikkoKoodi = orgyksikko.Trim();

                                    // Vain jos on jotain kirjoitettavaa
                                    if (orgYksikkoKoodi.Length > 0)
                                    {
                                        writer.WriteElementString("YksikkoKoodi", orgYksikkoKoodi);
                                    }
                                }
                                writer.WriteEndElement(); // JulkaisunOrgYksikot end          
                            }


                            if (elementti.Equals("TieteenalaKoodi") && arvo.Trim().Length > 0 )
                            {
                                validiElementti = false; // Ei lopussa saa kirjoittaa uudelleen

                                writer.WriteStartElement("TieteenalaKoodit"); // TieteenalaKoodit alkaa                                                     

                                string[] TieteenalaKooditArray = arvo.Split(';'); // datassa ; eroteltu jos on enemmän

                                int n = 1;

                                foreach (string tieteenala in TieteenalaKooditArray)
                                {
                                    string alaKoodi = tieteenala.Trim();

                                    // Vain jos on jotain kirjoitettavaa
                                    if (alaKoodi.Length > 0)
                                    {                               

                                        writer.WriteStartElement("TieteenalaKoodi");
                                        writer.WriteAttributeString("JNro", n.ToString() );
                                        writer.WriteString(alaKoodi);
                                        writer.WriteEndElement();
                                
                                        n++; // Laskuri
                                    }
                                }
                                writer.WriteEndElement(); // TieteenalaKoodit end          
                            }

                            if (elementti.Equals("KoulutusalaKoodi") && arvo.Trim().Length > 0)
                            {
                                validiElementti = false; // Ei lopussa saa kirjoittaa uudelleen

                                writer.WriteStartElement("KoulutusalaKoodit"); // KoulutusalaKoodit alkaa                                                     

                                string[] KoulutusalaKooditArray = arvo.Split(';'); // datassa ; eroteltu jos on enemmän

                                int n = 1; // Laskuri

                                foreach (string kala in KoulutusalaKooditArray)
                                {
                                    string kalaKoodi = kala.Trim();

                                    // Vain jos on jotain kirjoitettavaa
                                    if (kalaKoodi.Length > 0)
                                    {
                                        //writer.WriteElementString("KoulutusalaKoodi", kalaKoodi);

                                        writer.WriteStartElement("KoulutusalaKoodi");

                                            writer.WriteAttributeString("JNro", n.ToString());
                                            writer.WriteString(kalaKoodi);

                                        writer.WriteEndElement();

                                        n++; // Laskuri

                                    }
                                }
                                writer.WriteEndElement(); // KoulutusalaKoodit end          
                            }

                            // Ei tyhjiä avainsanoja
                            if (elementti.Equals("AvainsanaTeksti") && arvo.Trim().Length > 0 )
                            {
                                validiElementti = false; // Ei saa lopussa kirjoittaa uudelleen

                                writer.WriteStartElement("Avainsanat"); // Avainsanat alkaa                                                     

                                string[] AvainsanatArray = arvo.Split(';'); // datassa , tai ; eroteltu jos on enemmän

                                foreach (string sana in AvainsanatArray)
                                {
                                    string avainsana = sana.Trim();

                                    // Vain jos on jotain kirjoitettavaa
                                    if (avainsana.Length > 0)
                                    {
                                        writer.WriteElementString("AvainsanaTeksti", avainsana);
                                    }
                                }
                                writer.WriteEndElement(); // Avainsanat end          
                            }

                            // Voi olla kaksi ISBN-numeroa
                            if (elementti.Equals("ISBN") && arvo.Trim().Length > 0)
                            {
                                validiElementti = false; // Ei saa lopussa kirjoittaa uudelleen

                                string isbn1;
                                string isbn2;

                                if (arvo.Contains(";"))
                                {

                                    isbn1 = arvo.Split(';')[0].Trim();
                                    isbn2 = arvo.Split(';')[1].Trim();

                                    if (isbn1.Length > 0) 
                                    { 
                                        writer.WriteElementString("ISBN", isbn1);
                                    }

                                    if (isbn2.Length > 0)
                                    {
                                        writer.WriteElementString("ISBN", isbn2);
                                    }
                                }
                                else
                                {
                                    isbn1 = arvo.Trim();
                                    writer.WriteElementString("ISBN", isbn1);
                                    // Vain yksi ISBN		
                                }

                            }


                            // Voi olla kaksi ISSN-numeroa
                            if (elementti.Equals("ISSN") && arvo.Trim().Length > 0)
                            {
                                validiElementti = false; // Ei saa lopussa kirjoittaa uudelleen

                                string issn1;
                                string issn2;

                                if (arvo.Contains(";"))
                                {
                                    issn1 = arvo.Split(';')[0].Trim();
                                    issn2 = arvo.Split(';')[1].Trim();

                                    if (issn1.Length > 0)
                                    {
                                        writer.WriteElementString("ISSN", issn1);
                                    }

                                    if (issn2.Length > 0)
                                    {
                                        writer.WriteElementString("ISSN", issn2);
                                    }
                                    
                                }
                                else
                                {
                                    issn1 = arvo.Trim();
                                    writer.WriteElementString("ISSN", issn1);
                                    // Vain yksi ISSN		
                                }
                            }


                            if (elementti.Equals("RinnakkaistallennettuKytkin"))
                            {
                                validiElementti = false; // Ei saa lopussa kirjoittaa uudelleen                               

                                string kytkystr = arvo.Trim();

                                writer.WriteElementString("RinnakkaistallennettuKytkin", kytkystr);                               

                            }


                            if (elementti.Equals("RinnakkaistallennusOsoiteTeksti"))
                            {
                                validiElementti = false; // Ei saa lopussa kirjoittaa uudelleen

                                if (arvo.Trim().Length > 0)
                                {
                                    writer.WriteStartElement("Rinnakkaistallennettu"); // Rinnakkaistallennettu alkaa                                                    

                                        string osoite = arvo.Trim();
                                        writer.WriteElementString("RinnakkaistallennusOsoiteTeksti", osoite);

                                    writer.WriteEndElement(); // Rinnakkaistallennettu end          
                                }
                            }
                       

                            // Tulee myöhemmin <tekija>:n sisällä
                            if (elementti.Equals("ORCID"))
                            {
                                validiElementti = false;                        
                            }

                            //string etunimetString = "Etunimet";

                            // Tulee myöhemmin <tekija>:n sisällä
                            if (elementti.Equals("Etunimet"))
                            { 
                                validiElementti = false;
                            }

                            // Kaikki nimet on aluksi Sukunimi kentässä
                            string sukunimiString = "Sukunimi";

                            if (elementti.Equals(sukunimiString) && arvo.Trim().Length > 0 )
                            {
                                    validiElementti = false;

                                    writer.WriteStartElement("Tekijat"); // Tekijät alkaa                                                     

                                    string[] sukunimetArray = arvo.Split(';');

                                    int n = 0;
                                    int m = 0;

                                    int max_elements = sukunimetArray.Length;
                                    int nimia_num = 0;

                                    int maksimi_orcid = max_elements;

                                    foreach (string snimi in sukunimetArray)
                                    {
                                        string nimi = snimi.Trim();
                                        string suku = "";
                                        string etu = "";

                                        nimia_num++;

                                        // Jos on todella paljon porukkaa ilmoitetaan asiasta
                                        if ( nimia_num > 20 )
                                        {                                            
                                            errorTextBox.AppendText("Nimiä > " + nimia_num.ToString() + " yht :" + max_elements.ToString() + "\n\r\n\r");                                            
                                        }
                                        
                                        writer.WriteStartElement("Tekija"); // Tekijä alkaa

                                        // Vain jos on jotain
                                        if ( nimi.Length > 0) {

                                            
                                            // Jos on sukunimi, etunimi
                                            if (nimi.Contains(","))
                                            {
                                                suku = nimi.Split(',')[0].Trim();
                                                etu = nimi.Split(',')[1].Trim();

                                                writer.WriteElementString("Sukunimi", suku);
                                                writer.WriteElementString("Etunimet", etu);
                                            }
                                            else                                          
                                            {
                                                suku = nimi.Trim();
                                                writer.WriteElementString("Sukunimi", suku);
                                                // Ei ole etunimeä			
                                            }                                           


                                            // Yksikkökoodia
                                            string sisaElementti_yksikko = PalautaXMLElementtiNimi(j + 2);

                                            Boolean validisisaElementti = true;

                                            if (dataRow.ItemArray[j + 2] == DBNull.Value)
                                                dataRow.SetField(j + 2, string.Empty);

                                            string sisaArvo = dataRow.ItemArray[j + 2].ToString();
                                            string yksikkokoodi2String = "YksikkoKoodi2";

                                            string[] yksikotArray = sisaArvo.Split(';'); // datassa ; eroteltu jos on enemmän
                                                                                         // Jos on useampi, pitää olla järjestyksessä   

                                            if (sisaElementti_yksikko.Equals(yksikkokoodi2String))
                                            {
                                                validisisaElementti = true;
                                                sisaElementti_yksikko = "YksikkoKoodi";
                                            }


                                            if (validisisaElementti)
                                            {
                                                // Yksikköjä saa olla max 20
                                                if (n < 20) { 
                                                
                                                    // Vain jos on yksikkö
                                                        if (sisaArvo.Trim().Length > 0 && yksikotArray[n].Trim().Length > 0 )
                                                        {
                                                            writer.WriteStartElement("Yksikot"); // Yksikot alkaa
                                                        }        

                                                        // Jos on useampi, enemmän kuin yksi
                                                        if (yksikotArray.Length > 1)
                                                        {
                                                            if (yksikotArray[n].Trim().Length > 0)
                                                            { 
                                                                writer.WriteElementString(sisaElementti_yksikko, yksikotArray[n]); // Yksikkökoodi
                                                            }                                            
                                                        }
                                                        else
                                                        {
                                                            if (sisaArvo.Trim().Length > 0)
                                                            {
                                                                writer.WriteElementString(sisaElementti_yksikko, sisaArvo); // Yksikkökoodi
                                                            }
                                                        }

                                                        // Vain jos on yksikkö muistetaan sulkea elementti
                                                        if (sisaArvo.Trim().Length > 0 && yksikotArray[n].Trim().Length > 0 )
                                                        {
                                                            writer.WriteEndElement(); // Yksikot end
                                                        }
                                                }

                                            }  

                                            // ORCID
                                            string sisaElementti_orcid = PalautaXMLElementtiNimi(j + 3);

                                            validisisaElementti = false;

                                            if (dataRow.ItemArray[j + 3] == DBNull.Value)
                                                dataRow.SetField(j + 3, string.Empty);

                                            sisaArvo = dataRow.ItemArray[j + 3].ToString();

                                            string[] orcidArray = sisaArvo.Split(';'); // datassa ; eroteltu jos on enemmän

                                            int max_orcids = orcidArray.Length;

                                            if (sisaElementti_orcid.Equals("ORCID"))
                                            {
                                                validisisaElementti = true;
                                                sisaElementti_orcid = "ORCID";
                                            }

                                            if (validisisaElementti)
                                            {
                                                if (m < maksimi_orcid && m < max_orcids)
                                                {
                                                    if (orcidArray.Length > 1)
                                                    {
                                                        if (orcidArray[m].Trim().Length > 0)
                                                        {
                                                            writer.WriteElementString(sisaElementti_orcid, orcidArray[m]); // orcid
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (sisaArvo.Trim().Length > 0)
                                                        {
                                                            writer.WriteElementString(sisaElementti_orcid, sisaArvo); // orcid
                                                        }
                                                    }
                                                }
                                              }  
                                        }                      
                            
                                        writer.WriteEndElement(); // Tekijä end                                

                                        n++; // indeksi
                                        m++;
                                    }
                                    writer.WriteEndElement(); // Tekijät end                                               

                                }
                                



                               // Muu validi elementti? 
                               if (validiElementti) {

                                    if ( arvo.Trim().Length > 0) { 
                                        writer.WriteElementString(elementti, arvo); // eli joku muu elementti
                                    }

                                }

                               
                        }

                        writer.WriteEndElement();
                    }

                    writer.WriteEndElement();         
                    writer.WriteEndDocument();
                    writer.Flush();
                    writer.Close();          

            }
			catch (Exception ex)
			{
				MessageBox.Show("Virhe XML-writerissa: " + ex.Message); 
			}
            
            string xmlString = Encoding.UTF8.GetString(memStream.ToArray());

            memStream.Close();
            memStream.Dispose();

            return xmlString.ToString();
        }

        // Tallennetaan XML tiedosto
        private void TallennaXMLButton_Click(object sender, EventArgs e)
        {
            if ( TallennaXMLButton.Enabled )
            { 
                    DataTable dT = GetDataTableFromDGV(XMLdataGridView);  // Haetaan data XMLdataGridista        

                    string xmlString = JulkaisutXML(dT); // Viedään data ja luodaan xml, palautetaan string                     

                    // Tallennetaan tiedostoon
            
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.InitialDirectory = Directory.GetCurrentDirectory();

                    saveFileDialog1.Filter = "xml (*.xml)|*.xml";
                    saveFileDialog1.FilterIndex = 2;
                    saveFileDialog1.RestoreDirectory = true;            

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                                StreamWriter file = new StreamWriter(saveFileDialog1.FileName, false); 
                                
                                file.WriteLine(xmlString);                    
                                file.Close();                    

                                statusLabel.Text = "Ok!";
                                Statuslabel2.Text = "";

                                errorTextBox.Clear(); // Vanhaa pois

                                Poista_XML_virheLokiTiedosto("xmlvirheloki.txt");



                                string UusiXMLTiedosto = "XML-tiedosto " + saveFileDialog1.FileName + " luotu!";

                                AvaaXMLButton.Enabled = true;

                                setXMLTiedostoNimi(saveFileDialog1.FileName);
                                errorTextBox.Clear();
                                errorTextBox.AppendText(UusiXMLTiedosto);

                                ValidoiXMLButton.Enabled = true;

                                // "Pakkovalidointi"
                                if (ValidoiXMLButton.Enabled) 
                                    ValidoiXMLButton_Click(sender,e);


                                if (LokiOlemassa(masterloki))
                                {
                                    KirjoitaLokiin(UusiXMLTiedosto, masterloki);
                                }

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Virhe: Voi ei! Taas meni kaikki pieleen! Original error: " + ex.Message);
                        }
                    }           
            
                    MuunnaXMLButton.Enabled = false;

            }
            else
            {
                MessageBox.Show("Uutta XML-tiedostoa ei voi vielä tallentaa!");

                errorTextBox.Clear();


                if (LokiOlemassa(masterloki))
                {
                    KirjoitaLokiin("Uutta XML-tiedostoa ei voi vielä tallentaa!", masterloki);
                }

            }

        }

        // Suljetaan hela hoito ja lähdetään kotiin... :)
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (LokiOlemassa(masterloki))
            {
                KirjoitaLokiin("Ohjelma suljettiin.", masterloki);
            }


            this.Close();
        }

        // Rivinumerot gridiin
        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }

        // Nollaa Gridheader värit
        private void nollaGridVarit()
        {
                DataGridViewCellStyle rowStyle;

                int numCol = XMLdataGridView.ColumnCount;                

		        int num = 0;

                foreach (DataGridViewRow rivi in XMLdataGridView.Rows)
                {
                    try
                    {
                        for (int i = 0; i < numCol; i++) {
                            
                                XMLdataGridView.Rows[num].Cells[i].Style.BackColor = Color.Empty;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Empty;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;                            
                        }                       

                        num++;  // Rivinro laskuri
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Virhe: Voi ei! Error: " + ex.Message);
                    }
                }
        }
		

        // Valitse ISBN tai ISSN
        private void DataGridVCellClick(object sender, DataGridViewCellEventArgs e)
        {
	        try
	        {
                string info_str = "";

                if (e.ColumnIndex == 13) // ISBN
                {
                    if (XMLdataGridView.Rows[e.RowIndex].Cells[13] != null && XMLdataGridView.Rows[e.RowIndex].Cells[13].Value.ToString().Length > 0)
                    {
                        info_str = XMLdataGridView.Rows[e.RowIndex].Cells[13].Value.ToString().Trim();

                        string isbn1 = "";
                        string isbn2 = "";

                        // Jos on kaksi ISBN numeroa, kaksoispisteellä erotettuna
                        if (info_str.Contains(";"))
                        {
                            isbn1 = info_str.Split(';')[0].Trim();
                            isbn2 = info_str.Split(';')[1].Trim();
                        }
                        else
                        {
                            isbn1 = info_str.Trim();
                        }

                        bool val1 = TarkistaISBN(isbn1);
                        bool val2 = true;

                        if (isbn2.Length > 0)
                        {
                            val2 = TarkistaISBN(isbn2);
                        }
                        
                        if (val1)
                        {
                            string tarkiste = ISBN_tarkistemerkki(isbn1);

                            MessageBox.Show(isbn1 + " on validi ISBN!" + " Tarkiste: " + tarkiste, "ISBN tarkiste");
                        }
                        else
                        {
                            if (isbn1.Length > 0)
                            {
                                // Annetaan oikea tarkistusmerkki tiedoksi
                                string tarkiste = ISBN_tarkistemerkki(isbn1);

                                MessageBox.Show(isbn1 + " Ei ole validi ISBN!" + " Tarkiste: " + tarkiste, "ISBN tarkiste");
                            }   
                        }

                        if (val2 && isbn2.Length > 0)
                        {
                            string tarkiste = ISBN_tarkistemerkki(isbn2);

                            MessageBox.Show(isbn2 + " on validi ISBN!" + " Tarkiste: " + tarkiste, "ISBN tarkiste");
                        }
                        else
                        {

                            if (isbn2.Length > 0) { 
                                    // Annetaan oikea tarkistusmerkki tiedoksi
                                    string tarkiste = ISBN_tarkistemerkki(isbn2);

                                    MessageBox.Show(isbn2 + " Ei ole validi ISBN!" + " Tarkiste: " + tarkiste, "ISBN tarkiste");
                            }

                        }
                    }
                }


                if (e.ColumnIndex == 18) // ISSN
                {
                    if (XMLdataGridView.Rows[e.RowIndex].Cells[18] != null && XMLdataGridView.Rows[e.RowIndex].Cells[18].Value.ToString().Length > 0)
                    {
                        info_str = XMLdataGridView.Rows[e.RowIndex].Cells[18].Value.ToString();

                        string issn1 = "";
                        string issn2 = "";

                        // Jos on kaksi ISSN numeroa, kaksoispisteellä erotettuna
                        if (info_str.Contains(";"))
                        {
                            issn1 = info_str.Split(';')[0].Trim();
                            issn2 = info_str.Split(';')[1].Trim();
                        }
                        else
                        {
                            issn1 = info_str.Trim();
                        }

                        bool val1 = TarkistaISSN(issn1);
                        bool val2 = true;

                        if (issn2.Length > 0)
                        {
                            val2 = TarkistaISSN(issn2);
                        }

                        if (val1)
                        {
                            string tarkiste = ISSN_tarkistemerkki(issn1);

                            MessageBox.Show(issn1 + " on validi ISSN!" + " Tarkiste: " + tarkiste, "ISSN tarkiste");
                        }
                        else
                        {
                            if (issn1.Length > 0)
                            {
                                // Anna oikea tarkistusmerkki tiedoksi
                                string tarkiste = ISSN_tarkistemerkki(issn1);

                                MessageBox.Show(issn1 + " Ei ole validi ISSN!" + " Tarkiste: " + tarkiste, "ISSN tarkiste");
                            }
                        }

                        if (val2 && issn2.Length > 0)
                        {
                            string tarkiste = ISSN_tarkistemerkki(issn2);
                            MessageBox.Show(issn2 + " on validi ISSN!" + " Tarkiste: " + tarkiste, "ISSN tarkiste");
                        }
                        else
                        {
                            if (issn2.Length > 0) {
                                // Anna oikea tarkistusmerkki tiedoksi

                                string tarkiste = ISSN_tarkistemerkki(issn2);
                                MessageBox.Show(issn2 + " Ei ole validi ISSN!" + " Tarkiste: " + tarkiste, "ISSN tarkiste");
                            }
                        }
                    }
                }


                // CheckJulkaisuTyyppiKoodi
                if (e.ColumnIndex == 26) // Julkaisutyyppi
                {
                    if (XMLdataGridView.Rows[e.RowIndex].Cells[26] != null && XMLdataGridView.Rows[e.RowIndex].Cells[26].Value.ToString().Length > 0)
                    {
                        info_str = XMLdataGridView.Rows[e.RowIndex].Cells[26].Value.ToString();
                        Boolean val = CheckJulkaisuTyyppiKoodi(info_str);

                        if (val)
                        {
                            MessageBox.Show(info_str + " on sallittu Julkaisutyyppi koodi!");
                        }
                        else
                        {
                            MessageBox.Show(info_str + " ei ole sallittu Julkaisutyyppi koodi!");
                        }
                    }
                }

                // CheckTieteenaalakoodi
                if (e.ColumnIndex == 27) // Tieteenala
                {
                    if (XMLdataGridView.Rows[e.RowIndex].Cells[27] != null && XMLdataGridView.Rows[e.RowIndex].Cells[27].Value.ToString().Length > 0)
                    {
                        info_str = XMLdataGridView.Rows[e.RowIndex].Cells[27].Value.ToString();
                        Boolean val = CheckTieteenAlaKoodi(info_str);

                        if (val)
                        {
                            MessageBox.Show(info_str + " sisältää oikeita tieteenalakoodeja!");
                        }
                        else
                        {
                            MessageBox.Show(info_str + " sisältää virheellisiä tieteenalakoodeja!");
                        }
                    }
                }

	        }
	        catch (Exception ex)
	        {
                MessageBox.Show("Virhe: Voi ei! Error: " + ex.Message);
	        }
        }

        // "Yleinen" virhe herja
        private void virheIlmoitus2()
        {
            statusLabel.Text = "Aineisto sisältää virheitä!";
            Statuslabel2.Text = "XML-tiedostoa ei voi luoda!";
            statusLabel.Font = new Font(statusLabel.Font, FontStyle.Bold);
            statusLabel.ForeColor = Color.Red;
        }


        // Onko validi Julkaisutyyppi Koodi arvo Huom! 2015 alkaen
        private Boolean TarkistaISBNPakollisuus(string arvo)
        {
            string[] koodit = arvo.Split(';');  // Tulee lista ; eroteltuja koodiarvoja

            // A3 A4 B2 B3 C1  
            // Poistettu D2  D5  E2   13.2.2016
            string[] ISBNPakollisuusJosJulkaisuTyyppiKoodi = { "A3", "A4", "B2", "B3", "C1" };

            foreach (string koodi in koodit)
            {
                if (Array.Exists(ISBNPakollisuusJosJulkaisuTyyppiKoodi, element => element == koodi))
                {
                    return true;  // Eli ISBN on pakollinen				
                }
            }

            return false; // ISBN ei pakollinen kun ei ole pakollisten julkaisutyyppien joukossa
        }


        // Onko validi Julkaisutyyppi Koodi arvo Huom! 2015 alkaen
        private Boolean TarkistaISSNPakollisuus(string arvo)
        {
            string[] koodit = arvo.Split(';');  // Jos tulee lista ; eroteltuja koodiarvoja

            // A1 A2 A4 B1 B2 B3  
            // Poistettu D1 13.2.2016
            string[] ISSNPakollisuusJosJulkaisuTyyppiKoodi = { "A1", "A2", "A4", "B1", "B2", "B3" };

            foreach (string koodi in koodit)
            {
                if (Array.Exists(ISSNPakollisuusJosJulkaisuTyyppiKoodi, element => element == koodi))
                {
                    return true;  // Eli ISSN on pakollinen				
                }
            }

            return false; // ISSN ei pakollinen kun ei ole pakollisten julkaisutyyppien joukossa
        }


        // 2017 Lehden nimi pakollinen jos julkaisutyyppi D1
        private Boolean TarkistaLehdenNimiPakollisuus(string arvo, int vuosi)
        {
            string[] koodit = arvo.Split(';');  // Tulee lista ; eroteltuja koodiarvoja


            string[] LehdenNimiPakollisuusJosJulkaisuTyyppiKoodi = { "D1", "E1" };

            foreach (string koodi in koodit)
            {
                if (Array.Exists(LehdenNimiPakollisuusJosJulkaisuTyyppiKoodi, element => element == koodi) && (vuosi == 2017 || vuosi == 2018))
                {
                    return true;  // Eli Lehden nimi on pakollinen				
                }
            }

            return false; // Lehden nimi  ei pakollinen kun ei ole pakollisten julkaisutyyppien joukossa
        }






        // Onko validi Julkaisutyyppi Koodi B3 tai A4 ISBN tai ISSN pakollinen 
        private Boolean JosJulkaisuTyyppille_ISBN_tai_ISSN_pakollinen(string julkaisutyyppit, string isbn, string issn)
        {
            string[] koodit = julkaisutyyppit.Split(';');  // Tulee lista ; eroteltuja koodiarvoja

            // B3, A4
            string[] ISSN_tai_ISBN_pakollinen = { "B3","A4", "C2"  };

            foreach (string koodi in koodit)
            {
                if (Array.Exists(ISSN_tai_ISBN_pakollinen, element => element == koodi))
                {
                        if ((isbn.Trim().Length == 0) && (issn.Trim().Length == 0))
                        {
                            return true;  // eli ISSN tai ISBN on pakollinen julkaisutyypille A4, B3 ja C2											
                        }
                        
                }
            }

            return false; // ISSN tai ISBN ei pakollinen
        }



        // Onko validi Julkaisutyyppi Koodi E1 lehdennimi tai emojulkaisunnimi pakollinen 
        private Boolean JosJulkaisuTyyppille_LehdenNimi_tai_EmojulkaisunNimi_pakollinen(string julkaisutyyppit, string lehdennimi, string emojulkaisunnimi)
        {
            string[] koodit = julkaisutyyppit.Split(';');  // Tulee lista ; eroteltuja koodiarvoja

            // E1
            string[] LehdenNimi_tai_EmojulkaisunNimi_pakollinen = { "E1" };

            foreach (string koodi in koodit)
            {
                if (Array.Exists(LehdenNimi_tai_EmojulkaisunNimi_pakollinen, element => element == koodi))
                {
                    if ((lehdennimi.Trim().Length == 0) && (emojulkaisunnimi.Trim().Length == 0))
                    {
                        return true;  // eli LehdenNimi tai EmojulkaisunNimi pakollinen julkaisutyypille E1										
                    }

                }
            }

            return false; // LehdenNimi tai EmojulkaisunNimi ei pakollinen
        }







        // <KustantajanNimi>	2016 Pakollinen julkaisutyypeille A3, B2, C1  
        //                      2017 Pakollinen julkaisutyypeille A3, B2, C1, C2, D2, D4, D6, E2, E3   
        private Boolean TarkistaKustantajaPakollisuus(string arvo, int vuosi)
        {
            string[] koodit = arvo.Split(';');  // Jos tulee lista ; eroteltuja koodiarvoja

            // KustantajanNimi pakollinen julkaisutyypeille A3, B2, C1
            // C2 Poistettu 9.3.2016 
            string[] KustantajaJulkaisuTyyppiKoodi2016 = {"A3", "B2", "C1"};
            string[] KustantajaJulkaisuTyyppiKoodi2017 = { "A3", "B2", "C1", "C2", "D2", "D4", "D5", "D6", "E2", "E3" }; 
		
            foreach (string koodi in koodit)
            {
                if (Array.Exists(KustantajaJulkaisuTyyppiKoodi2016, element => element == koodi) && vuosi == 2016)
                {
                    return true;  // Eli Kustantaja nimi on pakollinen				
                }

                if (Array.Exists(KustantajaJulkaisuTyyppiKoodi2017, element => element == koodi) && (vuosi == 2017 || vuosi == 2018))
                {
                    return true;  // Eli Kustantaja nimi on pakollinen				
                }
            }

            return false; // Kustantaja ei pakollinen kun ei ole pakollisten julkaisutyyppien joukossa
        }



        // Onko KonferenssinNimi pakollinen 2016 A4, B3
        //      KonferenssinNimi pakollinen 2017 A4, B3, D3
        private Boolean TarkistaKonferenssinNimiPakollisuus(string julkaisutyyppit, int vuosi)
        {
            string[] koodit = julkaisutyyppit.Split(';');  // Tulee lista ; eroteltuja koodiarvoja

            // B3, A4
            string[] Onpakollinen2016 = { "B3", "A4" };
            string[] Onpakollinen2017 = { "B3", "A4", "D3" };

            foreach (string koodi in koodit)
            {
                if (Array.Exists(Onpakollinen2016, element => element == koodi) && vuosi == 2016)
                {
                    return true;  // eli on pakollinen									
                }

                if (Array.Exists(Onpakollinen2017, element => element == koodi) && (vuosi == 2017 || vuosi == 2018))
                {
                    return true;  // eli on pakollinen									
                }

            }

            return false; // ei pakollinen kun ei ole B3 tai A4
        }


        // Tarkistetaan Tekijatiedot
        private bool TarkistaTekijatiedotTeksti(string tekijat_str)
        {
            string tekijat = tekijat_str.Trim();           

            if (tekijat.Length > 0)
            {
                tekijat = Regex.Replace(tekijat.Trim(), " ", ""); // Ei välilyöntejä
                tekijat = Regex.Replace(tekijat.Trim(), ",", ""); // Ei "," merkkejä
                tekijat = Regex.Replace(tekijat.Trim(), ";", ""); // Ei ";" merkkejä

                if (tekijat.Trim().Length > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            return false;
        }


        // Onko sallittu arvo sallitaan tyhjä jos Sairaala tai tutkimuslaitos
        private Boolean CheckKytkinkoodi(string koodiarvot, string organisaatio)
        {
            string[] Koodit = { "0", "1", "" };

            string[] koodiarvo = koodiarvot.Trim().Split(';');

            string virheelliset = "";

            if (OnkoSairaalaTutkimus(organisaatio))
            {
                foreach (string koodi in koodiarvo)
                {
                    if (!Array.Exists(Koodit, element => element == koodi))
                    {
                        virheelliset = virheelliset + koodi + "  ";
                    }
                }
            }

            if (virheelliset.Trim().Length > 0)
            {
                // MessageBox.Show(virheelliset.ToString(), "Tieto vapaaehtoinen mutta annettu koodi: " + koodiarvot + " virheellinen!");
                return false;
            }
            else
            {
                return true;
            }

        }

        // Toinen yritys ISBN
        private string ISBN_ilmanviivoja(string isbn_str)
        {
            string isbn = isbn_str.Trim();

            if (isbn.Length > 0)
            {

                string pattern = @"/[- ]|^(ISBN(?:-1[03])?:?)";  // poista isbn(10/13) veke  

                isbn = Regex.Replace(isbn, pattern, "");
                isbn = Regex.Replace(isbn, " ", ""); // Ei välilyöntejä
                isbn = Regex.Replace(isbn.Trim(), "-", ""); // Ei "-" merkkejä	

                // ISBN 13
                if (isbn.Length == 13)
                {
                    return isbn;
                }
                else
                {
                    return "";

                }
            }
            else
            {
                return "";
            }                

        }
        

        // ---- TARKISTUKSET alkaa aina kun  Validoi -----
        // Pääkohtaiset virheet kerätään ja käsitellään
        // Validoidaan dataa ennen XML muodostusta
        // Tarkistukset indeksoituna Grid indeksin mukaan
        private void ValidoiButton_Click(object sender, EventArgs e)
        {          
            int num = 0; // Rivinro laskuri
            string virheita = "";

            string tiedotettavaa = "";

            string virhejulkaisuorgtunnus = "";

            statusLabel.Text = "";

            errorTextBox.Clear(); // Vanhat virheilmoitukset pois

            TallennaXMLButton.Enabled = false;
            LueLahdeTiedostoButton.Enabled = false; // Ei lueta uutta tietolähdettä

            setRowNumber(XMLdataGridView); // indeksoi rivit uudestaan
            nollaGridVarit(); // Nollaa väritykset
            
            errorTextBox.Clear(); // Tyhjennä errorboxi

            virheRivi.Clear();
            virheSarake.Clear();

            numericUpDownVirhe.Enabled = true;

            int maxVirheita = XMLdataGridView.RowCount;

            numericUpDownVirhe.Maximum = maxVirheita;
           

            //DataSet virheSetti = new DataSet("VirheSetti");
            //virheSetti.Tables.Add(virheTaulukko);

            string sarakeNimi;

            virheTaulukko.Columns.Clear();

            sarakeNimi = "VIRHE";
            virheTaulukko.Columns.Add(sarakeNimi, typeof(String));

            sarakeNimi = "julkaisuorgtunnus";
            virheTaulukko.Columns.Add(sarakeNimi, typeof(String));

            sarakeNimi = "rivi";
            virheTaulukko.Columns.Add(sarakeNimi, typeof(String));

            sarakeNimi = "viesti";
            virheTaulukko.Columns.Add(sarakeNimi, typeof(String));

            sarakeNimi = "tieto";
            virheTaulukko.Columns.Add(sarakeNimi, typeof(String));

            // Virhelistaa tarkistuscsv:tä varten
            object[] virheLista = new object[5];           

            DataGridViewCellStyle rowStyle;           
            
            int sarkkeita = XMLdataGridView.ColumnCount;

            DataGridViewColumn[] dgvc = new DataGridViewColumn[sarkkeita];

            for (int i = 0; i < sarkkeita; i++)
            {
	            dgvc[i] = XMLdataGridView.Columns[i];        
	            dgvc[i].HeaderCell.Style.BackColor = Color.Empty;            
                XMLdataGridView.Rows[num].Cells[i].Style.BackColor = Color.Empty;	
            }

            // indeksi  kenttä  (vuoden 2015 osalta)
            // 	0	 OrganisaatioTunnus
            // 	1	 IlmoitusVuosi
            // 	2	 JulkaisunTunnus
            // 	3	 JulkaisunTilaKoodi
            // 	4	 JulkaisunOrgTunnus
            // 	5	 YksikkoKoodi
            // 	6	 JulkaisuVuosi
            // 	7	 JulkaisunNimi
            // 	8	 TekijatiedotTeksti
            // 	9	 TekijoidenLkm
            // 	10	 SivunumeroTeksti
            // 	11	 Artikkelinumero
            // 	12	 AvainsanaTeksti
            // 	13	 ISBN
            // 	14	 JufoTunnus
            // 	15	 JufoLuokkaKoodi
            // 	16	 JulkaisumaaKoodi
            // 	17	 LehdenNimi
            // 	18	 ISSN
            // 	19	 VolyymiTeksti
            // 	20	 LehdenNumeroTeksti
            // 	21	 KonferenssinNimi
            // 	22	 KustantajanNimi
            // 	23	 KustannuspaikkaTeksti
            // 	24	 EmojulkaisunNimi
            // 	25	 EmojulkaisunToimittajatTeksti
            // 	26	 JulkaisutyyppiKoodi
            // 	27	 TieteenalaKoodi
            // 	28	 KoulutusalaKoodi
            // 	29	 YhteisjulkaisuKVKytkin
            // 	30	 YhteisjulkaisuSHPKytkin
            // 	31	 YhteisjulkaisuTutkimuslaitosKytkin
            // 	32	 YhteisjulkaisuMuuKytkin            
            // 	33	 JulkaisunKansainvalisyysKytkin            
            // 	34	 JulkaisunKieliKoodi
            // 	35	 AvoinSaatavuusKoodi
            // 	36	 EVOjulkaisuKytkin (Ei käytössä 2015!)
            // 	37	 DOI
            // 	38	 PysyvaOsoiteTeksti
            // 	39	 LahdetietokannanTunnus
            // 	40	 Sukunimi
            // 	41	 Etunimet
            // 	42	 YksikkoKoodi2
            // 	43	 ORCID
            // 	44	 HankenumeroTeksti
            // 	45	 RahoittajaOrgTunnus


            // 2016 ja 2017
            // indeksi  kenttä  (vuoden 2016 ja 2017 osalta)
            // 	0	 OrganisaatioTunnus
            // 	1	 IlmoitusVuosi
            // 	2	 JulkaisunTunnus
            // 	3	 JulkaisunTilaKoodi
            // 	4	 JulkaisunOrgTunnus
            // 	5	 YksikkoKoodi
            // 	6	 JulkaisuVuosi
            // 	7	 JulkaisunNimi
            // 	8	 TekijatiedotTeksti
            // 	9	 TekijoidenLkm
            // 	10	 SivunumeroTeksti
            // 	11	 Artikkelinumero
            // 	12	 AvainsanaTeksti
            // 	13	 ISBN
            // 	14	 JufoTunnus
            // 	15	 JufoLuokkaKoodi
            // 	16	 JulkaisumaaKoodi
            // 	17	 LehdenNimi
            // 	18	 ISSN
            // 	19	 VolyymiTeksti
            // 	20	 LehdenNumeroTeksti
            // 	21	 KonferenssinNimi
            // 	22	 KustantajanNimi
            // 	23	 KustannuspaikkaTeksti
            // 	24	 EmojulkaisunNimi
            // 	25	 EmojulkaisunToimittajatTeksti
            // 	26	 JulkaisutyyppiKoodi
            // 	27	 TieteenalaKoodi
            // 	28	 KoulutusalaKoodi
            // 	29	 YhteisjulkaisuKVKytkin

            // 	33 => 30 !! 2016:	 JulkaisunKansainvalisyysKytkin          Entinen YhteisjulkaisuTutkimuslaitosKytkin Ei käytössä 2016   
            // 	34 => 31 !! 2016:	 JulkaisunKieliKoodi            Entinen YhteisjulkaisuMuuKytkin  Ei käytössä 2016
            // 	35 => 32 !! 2016:	 AvoinSaatavuusKoodi

            //  33    !! 2016 YhteisjulkaisuYritysKytkin  Yhteisjulkaisu yrityksen kanssa Entinen YhteisjulkaisuSHPKytkin 2015

            // 	43	=> 34 !! 2016: Julkaisu rinnakkaistallennettu
            // 	44	=> 35 !! 2016: Rinnakkaistallennetun version verkko-osoite
            
            // 	36	 DOI
            // 	37	 PysyvaOsoiteTeksti
            // 	38	 LahdetietokannanTunnus
            // 	39	 Sukunimi
            // 	40	 Etunimet
            // 	41	 YksikkoKoodi2
            // 	42	 ORCID         
                        

            // Tässä nyt mennään eikä meinata, koko taulukko läpi, rivi riviltä, sarake sarakkeelta ...
            foreach (DataGridViewRow rivi in XMLdataGridView.Rows)
            {
                try
                {   
                    // Mitkä ovat pakolliset kentät ja mikä indeksi millekin tiedolle

                    // Muutettu että jos ei ole olemassa csv:ssä niin generoidaan organisaatiolle julkaisuntunnus
                    // organisaation nro + . 2015 ,  2016 tai 2017 + . juokseva numero + .csv
                    // JulkaisunOrgTunnus (julkaisun organisaatiokohtainen ID) on pakollinen
                    if ((rivi.Cells[4] == null) || (rivi.Cells[4].Value.ToString().Trim().Length == 0))
                    {

                        if (GeneroiJulkaisunOrganisaatioTunnus)
                        {
                            String julkaisunorganisaatiotunnus = "";

                            julkaisunorganisaatiotunnus = rivi.Cells[0].Value.ToString().Trim() + "." + vuosiIlmo.ToString() + "." + num.ToString() + ".csv";

                            rivi.Cells[4].Value = julkaisunorganisaatiotunnus;

                            RekisteroiVirheet("Generoitu:", julkaisunorganisaatiotunnus, (num + 1).ToString(), " tyhjän julkaisu id-arvon tilalle", "Korjaa jos on olemassa oikea organisaatiokohtainen julkaisun tunnus");

                            tiedotettavaa += "Generoitu julkaisunorganisaatiotunnus: " + julkaisunorganisaatiotunnus  + " rivi " + (num + 1) + " tyhjän arvon tilalle.\n\r\n\r";

                        }
                        else
                        {
                            XMLdataGridView.Rows[num].Cells[4].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Pakollinen tieto. Julkaisun organisaatio tunnus ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(4);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Pakollinen tieto. Julkaisun organisaatio tunnus ei voi olla tyhjä ", rivi.Cells[4].Value.ToString()); 

                            dgvc[4].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }



                    virhejulkaisuorgtunnus = " jultunnus: " + rivi.Cells[4].Value.ToString().Trim();

                    // OrganisaatioTunnus, JulkaisunOrgTunnus, JulkaisuVuosi, JulkaisunNimi, TekijatiedotTeksti                    
                    // OrganisaatioTunnus pakollinen
                    if ((rivi.Cells[0] == null) || (rivi.Cells[0].Value.ToString().Trim().Length == 0))
                    {                        
                        XMLdataGridView.Rows[num].Cells[0].Style.BackColor = Color.Red;
                        TallennaXMLButton.Enabled = false;                                               

                        virheita += "Pakollinen tieto. Organisaation tunnus ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                        virheIlmoitus2();

                        virheRivi.Add(num + 1);
                        virheSarake.Add(0);

                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Pakollinen tieto. Organisaation tunnus ei voi olla tyhjä ", rivi.Cells[0].Value.ToString()); 

                        dgvc[0].HeaderCell.Style.BackColor = Color.Red;
                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                        rowStyle.BackColor = Color.Red;
                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                    }


                    // Organisaatio koodin on oltava AMK, YO, Sairaala tai Tutkimuslaitos
                    if ((rivi.Cells[0] != null) && (rivi.Cells[0].Value.ToString().Trim().Length > 0))
                    {
                        if (!(OnkoAMK(rivi.Cells[0].Value.ToString().Trim()) || OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()) || OnkoYO(rivi.Cells[0].Value.ToString().Trim())))
                        {
                            XMLdataGridView.Rows[num].Cells[0].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;                                                   

                            string etunolla_str = ".";

                            if (OrganisaatiokoodinPituus( rivi.Cells[0].Value.ToString().Trim() ) < 5)
                            {
                                etunolla_str = " etunolla (0) puuttuu?";
                            }

                            virheita += "Organisaatiokoodi " + rivi.Cells[0].Value.ToString() + " virheellinen rivillä " + (num + 1) + " " + etunolla_str + virhejulkaisuorgtunnus + "\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(0);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Organisaatiokoodi virheellinen ", rivi.Cells[0].Value.ToString()); 

                            dgvc[0].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }

                    // Ilmoitusvuosi pakollinen tieto
                    if ((rivi.Cells[1] == null) || (rivi.Cells[1].Value.ToString().Length == 0))
                    {                        
                        XMLdataGridView.Rows[num].Cells[1].Style.BackColor = Color.Red;
                        TallennaXMLButton.Enabled = false;
                                                
                        virheita += "Pakollinen tieto. Ilmoitusvuosi ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                        virheIlmoitus2();

                        virheRivi.Add(num + 1);
                        virheSarake.Add(1);

                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Pakollinen tieto. Ilmoitusvuosi ei voi olla tyhjä ", rivi.Cells[1].Value.ToString()); 

                        dgvc[1].HeaderCell.Style.BackColor = Color.Red;
                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                        rowStyle.BackColor = Color.Red;
                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                    }

                    // Ilmoitusvuosi on oltava numeerinen ja oikea vuosi  
                    // Konfiguroitavana asetuksena mikä vuosi
                    if ((rivi.Cells[1] != null) || (rivi.Cells[1].Value.ToString().Length > 0))
                    {
                        int vuosiluku;                                                

                        if (!Int32.TryParse(rivi.Cells[1].Value.ToString().Trim(), out vuosiluku))
                        {
                            XMLdataGridView.Rows[num].Cells[1].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Ilmoitusvuoden pitää olla vuosiluku (numeerinen), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(1);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Ilmoitusvuoden pitää olla vuosiluku (numeerinen), virheellinen ", rivi.Cells[1].Value.ToString()); 

                            dgvc[1].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                        else if (vuosiluku != vuosiIlmo)
                        {
                            XMLdataGridView.Rows[num].Cells[1].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Ilmoitusvuoden pitää olla " + vuosiIlmo + ", rivillä " + (num + 1) + " vuosiluku: " + vuosiluku + " on virheellinen " + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(1);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Ilmoitusvuoden pitää olla " + vuosiIlmo.ToString() + " virheellinen " , rivi.Cells[1].Value.ToString()); 

                            dgvc[1].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }

                    // JulkaisunTilaKoodi vakio arvo 2 jos ei ole csv:ssä tätä tietoa, tietoa voi muuttaa sallittujen arvojen mukaan datagridissä
                    if ((rivi.Cells[3] != null) || (rivi.Cells[3].Value.ToString().Length > 0))
                    {
                        int tilakoodi;                        

                        if (!Int32.TryParse(rivi.Cells[3].Value.ToString().Trim(), out tilakoodi))
                        {
                            XMLdataGridView.Rows[num].Cells[3].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Julkaisun tilakoodi voi olla vain -1, 0, 1 tai 2 (numeerinen), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(3);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisun tilakoodi pitää olla 0, 1, 2 tai sitten -1", rivi.Cells[3].Value.ToString()); 

                            dgvc[3].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                        else if (tilakoodi <  -1 || tilakoodi > 2)
                        {
                            XMLdataGridView.Rows[num].Cells[3].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Julkaisun tilakoodi voi olla vain -1, 0, 1 tai 2, rivillä " + (num + 1) + " koodi: " + tilakoodi + " on virheellinen " + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(3);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisun tilakoodi pitää olla 0, 1, 2 tai sitten -1", rivi.Cells[3].Value.ToString()); 

                            dgvc[3].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                        }
                    }

                    // JulkaisunTilaKoodi ei kuitenkaan saa olla tyhjä
                    if ((rivi.Cells[3] == null) || (rivi.Cells[3].Value.ToString().Length == 0))
                    {
                        XMLdataGridView.Rows[num].Cells[3].Style.BackColor = Color.Red;
                        TallennaXMLButton.Enabled = false;                                               

                        virheita += "Pakollinen tieto. Julkaisun tilakoodi ei saa olla tyhjä (arvon pitää olla 0, 1, 2 tai sitten -1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                        virheIlmoitus2();

                        virheRivi.Add(num + 1);
                        virheSarake.Add(3);

                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisun tilakoodi ei saa olla tyhjä (arvon pitää olla 0, 1, 2 tai sitten -1)", rivi.Cells[3].Value.ToString()); 

                        dgvc[3].HeaderCell.Style.BackColor = Color.Red;
                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                        rowStyle.BackColor = Color.Red;
                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                    }
                    
   
                    // JulkaisuVuosi on pakollinen
                    if ((rivi.Cells[6] == null) || (rivi.Cells[6].Value.ToString().Length == 0))
                    {                        
                        XMLdataGridView.Rows[num].Cells[6].Style.BackColor = Color.Red;
                        TallennaXMLButton.Enabled = false;           

                        virheita += "Pakollinen tieto. Julkaisun vuosi ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                        virheIlmoitus2();

                        virheRivi.Add(num + 1);
                        virheSarake.Add(6);

                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisuvuosi ei saa olla tyhjä ", rivi.Cells[6].Value.ToString()); 

                        dgvc[6].HeaderCell.Style.BackColor = Color.Red;
                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                        rowStyle.BackColor = Color.Red;
                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                    } 


                    // Julkaisuvuoden oltava numeerinen ja 1970 - 2030 väliltä
                    if ((rivi.Cells[6] != null) || (rivi.Cells[6].Value.ToString().Length > 0))
                    {                        
                        int vuosiluku;              

                        if (!Int32.TryParse(rivi.Cells[6].Value.ToString().Trim(), out vuosiluku))
                        {
                            XMLdataGridView.Rows[num].Cells[6].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Julkaisu vuoden pitää olla vuosiluku (numeerinen), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(6);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisuvuoden pitää olla vuosiluku 1970 - 2030 ", rivi.Cells[6].Value.ToString()); 

                            dgvc[6].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                        else if ( vuosiluku <= 1970 || vuosiluku >= 2030 )
                        {
                            XMLdataGridView.Rows[num].Cells[6].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Julkaisu vuoden pitää olla 1970 - 2030, rivillä " + (num + 1) + " vuosiluku: " + vuosiluku + " on virheellinen " + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(6);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisuvuoden pitää olla 1970 - 2030 ", rivi.Cells[6].Value.ToString()); 

                            dgvc[6].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                        }
                    }


                    // Julkaisun nimi on pakollinen
                    if ((rivi.Cells[7] == null) || (rivi.Cells[7].Value.ToString().Length == 0))
                    {                        
                        XMLdataGridView.Rows[num].Cells[7].Style.BackColor = Color.Red;
                        TallennaXMLButton.Enabled = false;                                               

                        virheita += "Pakollinen tieto. Julkaisun nimi ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                        virheIlmoitus2();

                        virheRivi.Add(num + 1);
                        virheSarake.Add(7);

                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Pakollinen tieto. Julkaisun nimi ei voi olla tyhjä ", rivi.Cells[7].Value.ToString()); 

                        dgvc[7].HeaderCell.Style.BackColor = Color.Red;
                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                        rowStyle.BackColor = Color.Red;
                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                    }

                    // TekijatiedotTeksti on pakollinen
                    if ((rivi.Cells[8] == null) || (rivi.Cells[8].Value.ToString().Trim().Length == 0))
                    {                        
                        XMLdataGridView.Rows[num].Cells[8].Style.BackColor = Color.Red;
                        TallennaXMLButton.Enabled = false;                                               

                        virheita += "Pakollinen tieto. Tekijatiedot teksti ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                        virheIlmoitus2();

                        virheRivi.Add(num + 1);
                        virheSarake.Add(8);

                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Pakollinen tieto. Tekijatiedot teksti ei voi olla tyhjä ", rivi.Cells[8].Value.ToString()); 

                        dgvc[8].HeaderCell.Style.BackColor = Color.Red;
                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                        rowStyle.BackColor = Color.Red;
                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                    }                    

                    // TekijatiedotTeksti on pakollinen ja pitää sisältää muuta kuin ; merkkejä
                    if ((rivi.Cells[8] != null) || (rivi.Cells[8].Value.ToString().Trim().Length != 0))
                    {
                        bool val = TarkistaTekijatiedotTeksti(rivi.Cells[8].Value.ToString());

                        if (!val)
                        {
                            XMLdataGridView.Rows[num].Cells[8].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;                                                       

                            virheita += "Pakollinen tieto. Tekijatiedot teksti ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(8);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Pakollinen tieto. Tekijatiedot teksti ei voi olla tyhjä ", rivi.Cells[8].Value.ToString()); 

                            dgvc[8].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }



                    // Artikkelinumero ei ole pakollinen mutta pitää olla alle 50 merkkiä
                    if ((rivi.Cells[11] != null) || (rivi.Cells[8].Value.ToString().Trim().Length != 0))
                    {
                        int pituus = rivi.Cells[11].Value.ToString().Length;

                        if (pituus > 50)
                        {
                            XMLdataGridView.Rows[num].Cells[11].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Artikkelinumero ei ole pakollinen tieto, mutta oltava alle 50 merkkiä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(11);

                            RekisteroiVirheet("VIRHE", rivi.Cells[11].Value.ToString().Trim(), (num + 1).ToString(), " on oltava alle 50 merkkiä. Nyt se on: " + pituus.ToString() + " merkkiä", rivi.Cells[11].Value.ToString());

                            dgvc[11].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }

                    }


                    // Tarkista ISBN  indeksi 13 
                    if ((rivi.Cells[13] != null) && (rivi.Cells[13].Value.ToString().Length > 0))
                    {
                        string isbn1 = "";
                        string isbn2 = "";

                        // Jos on kaksi ISBN numeroa, kaksoispisteellä erotettuna
                        if (rivi.Cells[13].Value.ToString().Contains(";"))
                        {

                            isbn1 = rivi.Cells[13].Value.ToString().Split(';')[0].Trim();
                            isbn2 = rivi.Cells[13].Value.ToString().Split(';')[1].Trim();
                        }
                        else
                        {
                            isbn1 = rivi.Cells[13].Value.ToString().Trim();
                        }


                        bool val1 = TarkistaISBN(isbn1);

                        bool val2 = true;

                        if ( isbn2.Length > 0)
                        {

                            val2 = TarkistaISBN(isbn2);
                        }

                        if (!val1)
                        {
                            bool toinen_tsaanssi = false;

                            string isbn_ilman_mitaan = ISBN_ilmanviivoja(isbn1);

                            if (isbn_ilman_mitaan.Length == 13)
                            {
                                toinen_tsaanssi = TarkistaISBN(isbn_ilman_mitaan);
                            }

                            if (!toinen_tsaanssi)
                            {

                                XMLdataGridView.Rows[num].Cells[13].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "ISBN " + isbn1 + " virheellinen rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "ISBN virheellinen", isbn1);

                                virheRivi.Add(num + 1);
                                virheSarake.Add(13);

                                dgvc[13].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                            else
                            {   // Ei virhe vaan ilmoitus korjausliikkeestä
                                errorTextBox.AppendText("Korjataan ISBN:" + isbn1 + " => " + isbn_ilman_mitaan + " rivillä " + (num + 1) + "\n\r\n\r");

                                if (LokiOlemassa(masterloki))
                                {
                                    KirjoitaLokiin("Korjataan ISBN:" + isbn1 + " => " + isbn_ilman_mitaan + " rivillä " + (num + 1) + "\n\r\n\r", masterloki);
                                }

                                rivi.Cells[13].Value = isbn_ilman_mitaan;
                                XMLdataGridView.Rows[num].Cells[13].Style.BackColor = Color.Green;

                            }

                        }


                        if (!val2 && isbn2.Length > 0)
                        {
                            bool toinen_tsaanssi = false;

                            string isbn_ilman_mitaan = ISBN_ilmanviivoja(isbn2);

                            if (isbn_ilman_mitaan.Length == 13)
                            {
                                toinen_tsaanssi = TarkistaISBN(isbn_ilman_mitaan);
                            }

                            if (!toinen_tsaanssi)
                            {

                                XMLdataGridView.Rows[num].Cells[13].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "ISBN " + isbn2 + " virheellinen rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "ISBN virheellinen", isbn2);

                                virheRivi.Add(num + 1);
                                virheSarake.Add(13);

                                dgvc[13].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                            else
                            {   // Ei virhe vaan ilmoitus korjausliikkeestä
                                errorTextBox.AppendText("Korjataan ISBN:" + isbn2 + " => " + isbn_ilman_mitaan + " rivillä " + (num + 1) + "\n\r\n\r");

                                if (LokiOlemassa(masterloki))
                                {
                                    KirjoitaLokiin("Korjataan ISBN:" + isbn2 + " => " + isbn_ilman_mitaan + " rivillä " + (num + 1) + "\n\r\n\r", masterloki);
                                }

                                rivi.Cells[13].Value = isbn_ilman_mitaan;
                                XMLdataGridView.Rows[num].Cells[13].Style.BackColor = Color.Green;

                            }
                        }
                    }
                    

                    // Jos ISBN indeksi 13 puuttuu tarkista onko pakollinen
                    if ( ( rivi.Cells[13] == null ) || ( rivi.Cells[13].Value.ToString().Trim().Length == 0 ) )
                    {
                        Boolean val = TarkistaISBNPakollisuus(rivi.Cells[26].Value.ToString().Trim()); // Julkaisutyyppi

                        if ( val )
                        {                                                  
                            // Myös ISSN kenttä tyhjä?
                            if (rivi.Cells[18].Value.ToString().Trim().Length == 0)
                            {
                                XMLdataGridView.Rows[num].Cells[13].Style.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].Cells[18].Style.BackColor = Color.Red;
                                
                                XMLdataGridView.Rows[num].Cells[26].Style.BackColor = Color.Red;

                                TallennaXMLButton.Enabled = false;
                                                                
                                virheita += "ISBN tai ISSN on pakollinen jos julkaisutyyppi on " + rivi.Cells[26].Value.ToString() + " rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                                
                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(13);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "ISSN tai ISBN on pakollinen jos julkaisutyyppi ", rivi.Cells[26].Value.ToString()); 

                                dgvc[13].HeaderCell.Style.BackColor = Color.Red;  // ISBN sarake                              
                                dgvc[26].HeaderCell.Style.BackColor = Color.Red;  // Julkaisutyyppi sarake 

                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle; 
                           
                            }
                           
                        }
                    }


                    // Jos Lehden nimi indeksi 17 puuttuu tarkista onko pakollinen
                    if ((rivi.Cells[17] == null) || (rivi.Cells[17].Value.ToString().Trim().Length == 0))
                    {
                        Boolean val = TarkistaLehdenNimiPakollisuus(rivi.Cells[26].Value.ToString().Trim(), vuosiIlmo); // Julkaisutyyppi

                        if (val)
                        {
                            XMLdataGridView.Rows[num].Cells[17].Style.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].Cells[26].Style.BackColor = Color.Red;

                            TallennaXMLButton.Enabled = false;

                            virheita += "Lehden nimi on pakollinen jos julkaisutyyppi on " + rivi.Cells[26].Value.ToString() + " rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(17);

                            RekisteroiVirheet("VIRHE", rivi.Cells[17].Value.ToString().Trim(), (num + 1).ToString(), "Lehden nimi on pakollinen jos julkaisutyyppi ", rivi.Cells[26].Value.ToString());

                            dgvc[17].HeaderCell.Style.BackColor = Color.Red;  // Lehden nimi sarake                              
                            dgvc[26].HeaderCell.Style.BackColor = Color.Red;  // Julkaisutyyppi sarake 

                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }



                    // Tarkista ISSN indeksi 18            
                    // Voi olla 2 ISSN numeroa puolipilkulla erotettu
                    // Siivotaan välilyöntejä
                    if ((rivi.Cells[18] != null) && (rivi.Cells[18].Value.ToString().Length > 0))
                    {
                        string issn1 = "";
                        string issn2 = "";

                        // Jos on kaksi ISBN numeroa, puolipilkulla erotettuna
                        if (rivi.Cells[18].Value.ToString().Contains(";"))
                        {
                            issn1 = rivi.Cells[18].Value.ToString().Split(';')[0].Trim().ToUpper();
                            issn2 = rivi.Cells[18].Value.ToString().Split(';')[1].Trim().ToUpper();
                        }
                        else
                        {
                            issn1 = rivi.Cells[18].Value.ToString().Trim().ToUpper();
                        }

                        // Siivousta
                        string temp_str_issn1 = issn1;

                        if (temp_str_issn1.Length > 9)
                        {
                            string issn_korjattu = Regex.Replace(temp_str_issn1, " ", ""); // Ei välilyöntejä sisällä	

                            errorTextBox.AppendText("Korjataan ISSN poistetaan välilyöntejä: " + issn1 + " => " + issn_korjattu + " rivillä " + (num + 1) + "\n\r\n\r");

                            if (LokiOlemassa(masterloki))
                            {
                                KirjoitaLokiin("Korjataan ISSN poistetaan välilyöntejä: " + issn1 + " => " + issn_korjattu + " rivillä " + (num + 1) + "\n\r\n\r", masterloki);
                            }

                            issn1 = issn_korjattu;
                            XMLdataGridView.Rows[num].Cells[18].Style.BackColor = Color.Green;
                        }

                        string temp_str_issn2 = issn2.Trim();

                        if (temp_str_issn2.Length > 9)
                        {
                            string issn_korjattu = Regex.Replace(temp_str_issn2, " ", ""); // Ei välilyöntejä sisällä	

                            errorTextBox.AppendText("Korjataan ISSN poistetaan välilyöntejä: " + issn2 + " => " + issn_korjattu + " rivillä " + (num + 1) + "\n\r\n\r");

                            if (LokiOlemassa(masterloki))
                            {
                                KirjoitaLokiin("Korjataan ISSN poistetaan välilyöntejä: " + issn2 + " => " + issn_korjattu + " rivillä " + (num + 1) + "\n\r\n\r", masterloki);
                            }

                            issn2 = issn_korjattu;
                            XMLdataGridView.Rows[num].Cells[18].Style.BackColor = Color.Green;
                        }
                        

                        // Tarkistetaan vielä onko tarkiste ynm ok myös
                        bool val1 = TarkistaISSN(issn1);

                        bool val2 = true;

                        if (issn2.Length > 0)
                        {
                            val2 = TarkistaISSN(issn2);
                        }



                        if (!val1)
                        {
                            XMLdataGridView.Rows[num].Cells[18].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "ISSN " + issn1 + " virheellinen rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "ISSN virheellinen", issn1);

                            virheRivi.Add(num + 1);
                            virheSarake.Add(18);

                            dgvc[18].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                        

                        if (!val2 && issn2.Length > 0)
                        {
                            XMLdataGridView.Rows[num].Cells[18].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "ISSN " + issn2 + " virheellinen rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "ISSN virheellinen", issn2);

                            virheRivi.Add(num + 1);
                            virheSarake.Add(18);

                            dgvc[18].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                        
                        // Korjattu arvo
                        if (issn2.Length > 0)
                        {
                            rivi.Cells[18].Value = issn1 + ";" + issn2;
                        }
                        else
                        {
                            rivi.Cells[18].Value = issn1;
                        }
                    }



                    // Jos ISSN indeksi 18  puuttuu tarkista onko pakollinen
                    if ((rivi.Cells[18] == null) || (rivi.Cells[18].Value.ToString().Trim().Length == 0))
                    {
                        Boolean val = TarkistaISSNPakollisuus(rivi.Cells[26].Value.ToString().Trim()); // Julkaisutyyppi?

                        if (val)
                        {                                                   
                            // Myös ISBN tyhjä
                            if (rivi.Cells[13].Value.ToString().Trim().Length == 0)
                            {
                                XMLdataGridView.Rows[num].Cells[18].Style.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].Cells[13].Style.BackColor = Color.Red;

                                XMLdataGridView.Rows[num].Cells[26].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;                                                              

                                virheita += "ISSN tai ISBN on pakollinen jos julkaisutyyppi on " + rivi.Cells[26].Value.ToString() + " rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                               
                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(18);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "ISSN tai ISBN on pakollinen jos julkaisutyyppi ", rivi.Cells[26].Value.ToString()); 

                                dgvc[18].HeaderCell.Style.BackColor = Color.Red;  // ISSN sarake                              
                                dgvc[26].HeaderCell.Style.BackColor = Color.Red;  // Julkaisutyyppi sarake 

                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }                      
   
                        }
                    }


                    // EmojulkaisunToimittajatTeksti ei ole pakollinen mutta pitää olla alle 500 merkkiä
                    if ((rivi.Cells[25] != null))
                    {
                        int pituus = rivi.Cells[25].Value.ToString().Length;

                        if (pituus > 500)
                        {
                            XMLdataGridView.Rows[num].Cells[25].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "EmojulkaisunToimittajatTeksti ei ole pakollinen tieto, mutta oltava alle 500 merkkiä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(25);

                            RekisteroiVirheet("VIRHE", rivi.Cells[25].Value.ToString().Trim(), (num + 1).ToString(), " on oltava alle 500 merkkiä. Nyt se on: " + pituus.ToString() + " merkkiä", rivi.Cells[25].Value.ToString());

                            dgvc[25].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }

                    }		
                    
                                        
                    //  Julkaisun tyyppikoodi on pakollinen
                    if ((rivi.Cells[26] == null) || (rivi.Cells[26].Value.ToString().Length == 0))
                    {
                        XMLdataGridView.Rows[num].Cells[26].Style.BackColor = Color.Red;
                        TallennaXMLButton.Enabled = false;
                       
                        virheita += "Pakollinen tieto.  Julkaisun tyyppikoodi ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                        virheIlmoitus2();

                        virheRivi.Add(num + 1);
                        virheSarake.Add(26);

                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Pakollinen tieto.  Julkaisun tyyppikoodi ei voi olla tyhjä", rivi.Cells[26].Value.ToString()); 

                        dgvc[26].HeaderCell.Style.BackColor = Color.Red;
                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                        rowStyle.BackColor = Color.Red;
                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                    }

                    // Tarkista Julkaisun tyyppikoodi indeksi 26
                    if ((rivi.Cells[26] != null) && (rivi.Cells[26].Value.ToString().Length > 0))
                    {
                        Boolean val = CheckJulkaisuTyyppiKoodi(rivi.Cells[26].Value.ToString());

                        if (!val)                        
                        {
                            XMLdataGridView.Rows[num].Cells[26].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;                                                       

                            virheita += "Julkaisutyyppikoodi " + rivi.Cells[26].Value.ToString() + " virheellinen rivillä " + (num + 1) + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(26);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisutyyppikoodi virheellinen ", rivi.Cells[26].Value.ToString()); 

                            dgvc[26].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }                        
                    }



                    // Tarkista ISSN_tai_ISBN_pakollinen jos julkaisun tyyppi C2, B3, A4 . Julkaisun tyyppi koodi indeksi 26  ISSN_tai_ISBN_pakollinen = { "B3","A4","C2" }  
                    // ISBN indeksi 13,  ISSN indeksi 18

                    if ((rivi.Cells[26] != null) && (rivi.Cells[26].Value.ToString().Length > 0))
                    {
                        Boolean val = JosJulkaisuTyyppille_ISBN_tai_ISSN_pakollinen(rivi.Cells[26].Value.ToString().Trim(), rivi.Cells[13].Value.ToString().Trim(), rivi.Cells[18].Value.ToString().Trim());

                        if (val)
                        {      
                            string virhe_teksti = "";                            

                            virhe_teksti = virhe_teksti + " ISBN tai ISSN ";
                            XMLdataGridView.Rows[num].Cells[26].Style.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].Cells[18].Style.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].Cells[13].Style.BackColor = Color.Red;

                            TallennaXMLButton.Enabled = false;

                            virheita += "Julkaisutyyppikoodilla " + rivi.Cells[26].Value.ToString() + virhe_teksti + " on pakollinen rivillä: " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(13);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisutyyppikoodilla " + rivi.Cells[26].Value.ToString() + " ISBN tai ISSN ", rivi.Cells[26].Value.ToString()); 

                            dgvc[26].HeaderCell.Style.BackColor = Color.Red;
                            dgvc[13].HeaderCell.Style.BackColor = Color.Red;
                            dgvc[18].HeaderCell.Style.BackColor = Color.Red;

                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }

                    }



                    // Tarkista Tieteenalakoodit indeksi 27
                    if ((rivi.Cells[27] != null) && (rivi.Cells[27].Value.ToString().Length > 0))
                    {
                        Boolean val = CheckTieteenAlaKoodi(rivi.Cells[27].Value.ToString());

                        if (!val)
                        {
                            XMLdataGridView.Rows[num].Cells[27].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Tieteenalakoodi " + rivi.Cells[27].Value.ToString() + " virheellinen rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(27);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Tieteenalakoodi virheellinen ", rivi.Cells[27].Value.ToString()); 

                            dgvc[27].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }


                    // Tarkista Koulutusalakoodit indeksi 28
                    if ((rivi.Cells[28] != null) && (rivi.Cells[28].Value.ToString().Length > 0))
                    {
                        // Indeksi 0 on organisaation koodi
                        Boolean val = CheckKoulutuskoodi(rivi.Cells[28].Value.ToString(), rivi.Cells[0].Value.ToString());

                        if (!val)
                        {
                            XMLdataGridView.Rows[num].Cells[28].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Koulutusalakoodi " + rivi.Cells[28].Value.ToString() + " virheellinen rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(28);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Koulutusalakoodi virheellinen ", rivi.Cells[28].Value.ToString()); 

                            dgvc[28].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }


                    // Julkaisutyyppi koodilla "E1" Lehden nimi tai Emojulkaisun nimi pakollinen vuodesta 2017 alkaen
                    if ((rivi.Cells[26] != null) && (rivi.Cells[26].Value.ToString().Trim() == "E1") && (vuosiIlmo == 2017))
                    {
                        Boolean val = JosJulkaisuTyyppille_LehdenNimi_tai_EmojulkaisunNimi_pakollinen(rivi.Cells[26].Value.ToString().Trim(), rivi.Cells[17].Value.ToString().Trim(), rivi.Cells[24].Value.ToString().Trim());

                        if (val)
                        {
                            string virhe_teksti = "";

                            virhe_teksti = virhe_teksti + " Lehden nimi tai Emojulkaisun nimi ";

                            XMLdataGridView.Rows[num].Cells[26].Style.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].Cells[17].Style.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].Cells[24].Style.BackColor = Color.Red;

                            TallennaXMLButton.Enabled = false;

                            virheita += "Julkaisutyyppikoodilla " + rivi.Cells[26].Value.ToString() + virhe_teksti + " on pakollinen rivillä: " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(17);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisutyyppikoodilla " + rivi.Cells[26].Value.ToString() + " vuodesta 2017 alkaen Lehden nimi tai Emojulkaisun nimi ", rivi.Cells[26].Value.ToString());

                            dgvc[26].HeaderCell.Style.BackColor = Color.Red;
                            dgvc[17].HeaderCell.Style.BackColor = Color.Red;
                            dgvc[24].HeaderCell.Style.BackColor = Color.Red;

                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }

                    }		










                    // Tekijän sukunimi 2015 index 40, 2016 index 39
                    if (vuosiIlmo == 2015)
                    {

                        // Tekijän sukunimi on pakollinen
                        if ((rivi.Cells[40] == null) || (rivi.Cells[40].Value.ToString().Length == 0))
                        {
                            XMLdataGridView.Rows[num].Cells[40].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Pakollinen tieto.  Julkaisun tekijän nimi ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(40);


                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Pakollinen tieto.  Julkaisun tekijän nimi ei voi olla tyhjä", rivi.Cells[40].Value.ToString()); 

                            dgvc[40].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }


                        // Tekijän sukunimi siistimistä
                        if ((rivi.Cells[40] != null) || (rivi.Cells[40].Value.ToString().Trim().Length > 0))
                        {

                            string str = PoistaViimeisetPuolipisteet(rivi.Cells[40].Value.ToString().Trim());

                            rivi.Cells[40].Value = str;
                        }
                    }
                    

                    if (vuosiIlmo == 2016 || vuosiIlmo == 2017 || vuosiIlmo == 2018)
                    {

                        // Tekijän sukunimi on pakollinen
                        if ((rivi.Cells[39] == null) || (rivi.Cells[39].Value.ToString().Length == 0))
                        {
                            XMLdataGridView.Rows[num].Cells[39].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Pakollinen tieto.  Julkaisun tekijän nimi ei voi olla tyhjä, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(39);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Pakollinen tieto. Julkaisun tekijän nimi ei voi olla tyhjä", rivi.Cells[39].Value.ToString()); 

                            dgvc[39].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }


                        // Tekijän sukunimi siistimistä
                        if ((rivi.Cells[39] != null) || (rivi.Cells[39].Value.ToString().Trim().Length > 0))
                        {

                            string str = PoistaViimeisetPuolipisteet(rivi.Cells[39].Value.ToString().Trim());

                            rivi.Cells[39].Value = str;
                        }
                    }

                    

                    // Jos Kustantajan nimi indeksi 22 puuttuu. Tarkista onko pakollinen...
                    if ((rivi.Cells[22] == null) || (rivi.Cells[22].Value.ToString().Trim().Length == 0))
                    {
                        bool val = TarkistaKustantajaPakollisuus(rivi.Cells[26].Value.ToString(), vuosiIlmo);

                        if (val)
                        {
                            XMLdataGridView.Rows[num].Cells[22].Style.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].Cells[26].Style.BackColor = Color.Red;

                            TallennaXMLButton.Enabled = false;

                            virheita += "Kustantajan nimi on pakollinen jos julkaisutyyppi " + rivi.Cells[26].Value.ToString() + " rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(22);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Kustantajan nimi on pakollinen jos julkaisutyyppi", rivi.Cells[26].Value.ToString()); 

                            dgvc[22].HeaderCell.Style.BackColor = Color.Red;  // Kustantajan nimi sarake                              
                            dgvc[26].HeaderCell.Style.BackColor = Color.Red;  // Julkaisutyyppi sarake 

                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }


                    // Jos KonferenssinNimi indeksi 21 puuttuu. Tarkista onko pakollinen...
                    if ((rivi.Cells[21] == null) || (rivi.Cells[21].Value.ToString().Trim().Length == 0))
                    {
                        bool val = TarkistaKonferenssinNimiPakollisuus(rivi.Cells[26].Value.ToString(), vuosiIlmo);

                        if (val)
                        {
                            XMLdataGridView.Rows[num].Cells[21].Style.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].Cells[26].Style.BackColor = Color.Red;

                            TallennaXMLButton.Enabled = false;

                            virheita += "Konferenssin nimi on pakollinen jos julkaisutyyppi " + rivi.Cells[26].Value.ToString() + " rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(21);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Konferenssin nimi on pakollinen jos julkaisutyyppi", rivi.Cells[26].Value.ToString()); 

                            dgvc[21].HeaderCell.Style.BackColor = Color.Red;  // Kustantajan nimi sarake                              
                            dgvc[26].HeaderCell.Style.BackColor = Color.Red;  // Julkaisutyyppi sarake 

                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }

                    // Tekijöiden lukumäärä numeerinen ja oltava positiivinen kokonaisluku
                    if ((rivi.Cells[9] != null) || (rivi.Cells[9].Value.ToString().Length > 0))
                    {
                        int lkm = 0;

                        if (!Int32.TryParse(rivi.Cells[9].Value.ToString().Trim(), out lkm) || lkm <= 0 )
                        {
                            XMLdataGridView.Rows[num].Cells[9].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Tekijöiden lukumäärä pitää olla numeerinen positiivinen kokonaisluku, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(9);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Tekijöiden lukumäärä pitää olla numeerinen positiivinen kokonaisluku", rivi.Cells[9].Value.ToString()); 

                            dgvc[9].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }


                    // Julkaisumaa koodin oikeellisuus
                    // Korjataan koodia jos puutuu etunollia (0)

                    if ((rivi.Cells[16] != null) && (rivi.Cells[16].Value.ToString().Length > 0))
                    {

                        if (rivi.Cells[16].Value.ToString().Trim().Length < 3)
                        {
                            string maakoodi = "0" + rivi.Cells[16].Value.ToString().Trim(); // Lisätään etunolla

                            // Vieläkin puuttuu 0?
                            if (maakoodi.Trim().Length < 3)
                            {
                                maakoodi = "0" + maakoodi;
                            }

                            errorTextBox.AppendText("Korjataan maakoodi lisätään eteen etunolla (0+): " + rivi.Cells[16].Value.ToString() + " => " + maakoodi + " rivillä " + (num + 1) + "\n\r\n\r");

                            if (LokiOlemassa(masterloki))
                            {
                                KirjoitaLokiin("Korjataan maakoodi  maakoodi lisätään eteen etunolla (0+): " + rivi.Cells[16].Value.ToString() + " => " + maakoodi + " rivillä " + (num + 1) + "\n\r\n\r", masterloki);
                            }

                            rivi.Cells[16].Value = maakoodi;

                            XMLdataGridView.Rows[num].Cells[16].Style.BackColor = Color.Green;
                        }	
						

                        Boolean val = CheckMaaKoodi(rivi.Cells[16].Value.ToString());

                        if (!val)
                        {
                            XMLdataGridView.Rows[num].Cells[16].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "Julkaisumaa koodi " + rivi.Cells[16].Value.ToString() + " virheellinen rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(16);


                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisumaa koodi virheellinen", rivi.Cells[16].Value.ToString()); 

                            dgvc[16].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                    }

                    // JulkaisunKieliKoodi 2015
                    if (vuosiIlmo == 2015)
                    {
                        // Kieli koodin oikeellisuus
                        if ((rivi.Cells[34] != null) && (rivi.Cells[34].Value.ToString().Length > 0))
                        {
                            Boolean val = CheckKieliKoodi(rivi.Cells[34].Value.ToString());

                            if (!val)
                            {
                                XMLdataGridView.Rows[num].Cells[34].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "Julkaisun kielikoodi " + rivi.Cells[34].Value.ToString() + " virheellinen rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(34);



                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisun kielikoodi virheellinen", rivi.Cells[34].Value.ToString()); 

                                dgvc[34].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }

                    }

                    // JulkaisunKieliKoodi 2016  indeksi 34 => 32
                    if (vuosiIlmo == 2016 || vuosiIlmo == 2017 || vuosiIlmo == 2018)
                    {
                        // Kieli koodin oikeellisuus
                        if ((rivi.Cells[31] != null) && (rivi.Cells[31].Value.ToString().Length > 0))
                        {
                            Boolean val = CheckKieliKoodi(rivi.Cells[31].Value.ToString());

                            if (!val)
                            {
                                XMLdataGridView.Rows[num].Cells[31].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "Julkaisun kielikoodi " + rivi.Cells[31].Value.ToString() + " virheellinen rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(31);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Julkaisun kielikoodi virheellinen", rivi.Cells[31].Value.ToString()); 

                                dgvc[31].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }
                    }
                   

                    // YhteisjulkaisuKVKytkin oltava numeerinen 0 tai 1
                    if ((rivi.Cells[29] != null) || (rivi.Cells[29].Value.ToString().Length > 0))
                    {                        
	                    int nollaYx;

	                    if (!Int32.TryParse(rivi.Cells[29].Value.ToString().Trim(), out nollaYx))
	                    {
		                    XMLdataGridView.Rows[num].Cells[29].Style.BackColor = Color.Red;
		                    TallennaXMLButton.Enabled = false;

                            virheita += "YhteisjulkaisuKVKytkin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

		                    virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(29);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuKVKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[29].Value.ToString()); 

		                    dgvc[29].HeaderCell.Style.BackColor = Color.Red;
		                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
		                    rowStyle.BackColor = Color.Red;
		                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
	                    }
	                    else if ( nollaYx < 0 || nollaYx > 1 )
	                    {
		                    XMLdataGridView.Rows[num].Cells[29].Style.BackColor = Color.Red;
		                    TallennaXMLButton.Enabled = false;

                            virheita += "YhteisjulkaisuKVKytkin pitää olla 0 (ei) tai 1 (kyllä), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

		                    virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(29);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuKVKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[29].Value.ToString()); 

		                    dgvc[29].HeaderCell.Style.BackColor = Color.Red;
		                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
		                    rowStyle.BackColor = Color.Red;
		                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

	                    }
                    }

                    if (vuosiIlmo == 2015)
                    {
                        // YhteisjulkaisuSHPKytkin oltava numeerinen 0 tai 1
                        if ((rivi.Cells[30] != null) || (rivi.Cells[30].Value.ToString().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[30].Value.ToString().Trim(), out nollaYx))
                            {
                                if (!OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    XMLdataGridView.Rows[num].Cells[30].Style.BackColor = Color.Red;
                                    TallennaXMLButton.Enabled = false;

                                    virheita += "YhteisjulkaisuSHPKytkin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                    virheIlmoitus2();

                                    virheRivi.Add(num + 1);
                                    virheSarake.Add(30);

                                    RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuSHPKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[30].Value.ToString()); 


                                    dgvc[30].HeaderCell.Style.BackColor = Color.Red;
                                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                    rowStyle.BackColor = Color.Red;
                                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                }

                                if (OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    Boolean val = CheckKytkinkoodi(rivi.Cells[30].Value.ToString(), rivi.Cells[0].Value.ToString().Trim());

                                    if (!val)
                                    {
                                        XMLdataGridView.Rows[num].Cells[30].Style.BackColor = Color.Yellow;
                                        TallennaXMLButton.Enabled = false;

                                        virheita += "YhteisjulkaisuSHPKytkin vapaaehtoinen tieto sairaaloille ja tutkimuslaitoksille. Jos muu kuin tyhjä, niin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + " Virheellinen: " + rivi.Cells[30].Value.ToString() + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                        virheIlmoitus2();

                                        virheRivi.Add(num + 1);
                                        virheSarake.Add(30);

                                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuSHPKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[30].Value.ToString()); 

                                        dgvc[30].HeaderCell.Style.BackColor = Color.Yellow;
                                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                        rowStyle.BackColor = Color.Yellow;
                                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                    }
                                }


                            }
                            else if (nollaYx < 0 || nollaYx > 1)
                            {
                                XMLdataGridView.Rows[num].Cells[30].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "YhteisjulkaisuSHPKytkin pitää olla 0 (ei) tai 1 (kyllä), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(30);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuSHPKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[30].Value.ToString()); 

                                dgvc[30].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                        }
                    }


                    if (vuosiIlmo == 2015)
                    {
                        // YhteisjulkaisuTutkimuslaitosKytkin oltava numeerinen 0 tai 1
                        if ((rivi.Cells[31] != null) || (rivi.Cells[31].Value.ToString().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[31].Value.ToString().Trim(), out nollaYx))
                            {
                                if (!OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    XMLdataGridView.Rows[num].Cells[31].Style.BackColor = Color.Red;
                                    TallennaXMLButton.Enabled = false;

                                    virheita += "YhteisjulkaisuTutkimuslaitosKytkin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                    virheIlmoitus2();

                                    virheRivi.Add(num + 1);
                                    virheSarake.Add(31);


                                    RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuTutkimuslaitosKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[31].Value.ToString()); 

                                    dgvc[31].HeaderCell.Style.BackColor = Color.Red;
                                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                    rowStyle.BackColor = Color.Red;
                                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                }

                                if (OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    Boolean val = CheckKytkinkoodi(rivi.Cells[31].Value.ToString(), rivi.Cells[0].Value.ToString().Trim());

                                    if (!val)
                                    {
                                        XMLdataGridView.Rows[num].Cells[31].Style.BackColor = Color.Yellow;
                                        TallennaXMLButton.Enabled = false;

                                        virheita += "YhteisjulkaisuTutkimuslaitosKytkin vapaaehtoinen tieto sairaaloille ja tutkimuslaitoksille. Jos muu kuin tyhjä, niin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + " Virheellinen: " + rivi.Cells[31].Value.ToString() + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                        virheIlmoitus2();

                                        virheRivi.Add(num + 1);
                                        virheSarake.Add(31);

                                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuTutkimuslaitosKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[31].Value.ToString()); 

                                        dgvc[31].HeaderCell.Style.BackColor = Color.Yellow;
                                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                        rowStyle.BackColor = Color.Yellow;
                                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                    }

                                }



                            }
                            else if (nollaYx < 0 || nollaYx > 1)
                            {
                                XMLdataGridView.Rows[num].Cells[31].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "YhteisjulkaisuTutkimuslaitosKytkin pitää olla 0 (ei) tai 1 (kyllä), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(31);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuTutkimuslaitosKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[31].Value.ToString()); 

                                dgvc[31].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                        }
                    }


                    if (vuosiIlmo == 2015)
                    {
                        // YhteisjulkaisuMuuKytkin oltava numeerinen 0 tai 1
                        if ((rivi.Cells[32] != null) || (rivi.Cells[32].Value.ToString().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[32].Value.ToString().Trim(), out nollaYx))
                            {
                                if (!OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    XMLdataGridView.Rows[num].Cells[32].Style.BackColor = Color.Red;
                                    TallennaXMLButton.Enabled = false;

                                    virheita += "YhteisjulkaisuMuuKytkin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                    virheIlmoitus2();

                                    virheRivi.Add(num + 1);
                                    virheSarake.Add(32);

                                    RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuMuuKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[32].Value.ToString()); 

                                    dgvc[32].HeaderCell.Style.BackColor = Color.Red;
                                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                    rowStyle.BackColor = Color.Red;
                                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                }

                                if (OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    Boolean val = CheckKytkinkoodi(rivi.Cells[32].Value.ToString(), rivi.Cells[0].Value.ToString().Trim());

                                    if (!val)
                                    {
                                        XMLdataGridView.Rows[num].Cells[32].Style.BackColor = Color.Yellow;
                                        TallennaXMLButton.Enabled = false;

                                        virheita += "YhteisjulkaisuMuuKytkin vapaaehtoinen tieto sairaaloille ja tutkimuslaitoksille. Jos muu kuin tyhjä, niin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + " Virheellinen: " + rivi.Cells[32].Value.ToString() + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                        virheIlmoitus2();

                                        virheRivi.Add(num + 1);
                                        virheSarake.Add(32);
                                        
                                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuMuuKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[32].Value.ToString()); 

                                        dgvc[32].HeaderCell.Style.BackColor = Color.Yellow;
                                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                        rowStyle.BackColor = Color.Yellow;
                                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                    }
                                }

                            }
                            else if (nollaYx < 0 || nollaYx > 1)
                            {
                                XMLdataGridView.Rows[num].Cells[32].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "YhteisjulkaisuMuuKytkin pitää olla 0 (ei) tai 1 (kyllä), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(32);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuMuuKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[32].Value.ToString()); 

                                dgvc[32].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                        }
                    }


                    if (vuosiIlmo == 2015)
                    {
                        // JulkaisunKansainvalisyysKytkin oltava numeerinen 0 tai 1
                        if ((rivi.Cells[33] != null) || (rivi.Cells[33].Value.ToString().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[33].Value.ToString().Trim(), out nollaYx))
                            {

                                if (!OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    XMLdataGridView.Rows[num].Cells[33].Style.BackColor = Color.Red;
                                    TallennaXMLButton.Enabled = false;

                                    virheita += "JulkaisunKansainvalisyysKytkin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                    virheIlmoitus2();

                                    virheRivi.Add(num + 1);
                                    virheSarake.Add(33);

                                    RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "JulkaisunKansainvalisyysKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[33].Value.ToString()); 

                                    dgvc[33].HeaderCell.Style.BackColor = Color.Red;
                                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                    rowStyle.BackColor = Color.Red;
                                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                }

                                if (OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    Boolean val = CheckKytkinkoodi(rivi.Cells[33].Value.ToString(), rivi.Cells[0].Value.ToString().Trim());

                                    if (!val)
                                    {
                                        XMLdataGridView.Rows[num].Cells[33].Style.BackColor = Color.Yellow;
                                        TallennaXMLButton.Enabled = false;

                                        virheita += "JulkaisunKansainvalisyysKytkin vapaaehtoinen tieto sairaaloille ja tutkimuslaitoksille. Jos muu kuin tyhjä, niin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + " Virheellinen: " + rivi.Cells[33].Value.ToString() + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                        virheIlmoitus2();

                                        virheRivi.Add(num + 1);
                                        virheSarake.Add(33);

                                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "JulkaisunKansainvalisyysKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[33].Value.ToString()); 

                                        dgvc[33].HeaderCell.Style.BackColor = Color.Yellow;
                                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                        rowStyle.BackColor = Color.Yellow;
                                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                    }

                                }


                            }
                            else if (nollaYx < 0 || nollaYx > 1)
                            {
                                XMLdataGridView.Rows[num].Cells[33].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "JulkaisunKansainvalisyysKytkin pitää olla 0 (ei) tai 1 (kyllä), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(33);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "JulkaisunKansainvalisyysKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[33].Value.ToString()); 

                                dgvc[33].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                        }

                    }

                    // JulkaisunKansainvalisyysKytkin Indeksi 33 => 30 2016
                    if (vuosiIlmo == 2016 || vuosiIlmo == 2017 || vuosiIlmo == 2018)
                    {
                        // JulkaisunKansainvalisyysKytkin oltava numeerinen 0 tai 1
                        if ((rivi.Cells[30] != null) || (rivi.Cells[30].Value.ToString().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[30].Value.ToString().Trim(), out nollaYx))
                            {

                                if (!OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    XMLdataGridView.Rows[num].Cells[30].Style.BackColor = Color.Red;
                                    TallennaXMLButton.Enabled = false;

                                    virheita += "JulkaisunKansainvalisyysKytkin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                    virheIlmoitus2();

                                    virheRivi.Add(num + 1);
                                    virheSarake.Add(30);

                                    RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "JulkaisunKansainvalisyysKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[30].Value.ToString()); 

                                    dgvc[30].HeaderCell.Style.BackColor = Color.Red;
                                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                    rowStyle.BackColor = Color.Red;
                                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                }

                                if (OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    Boolean val = CheckKytkinkoodi(rivi.Cells[30].Value.ToString(), rivi.Cells[0].Value.ToString().Trim());

                                    if (!val)
                                    {
                                        XMLdataGridView.Rows[num].Cells[30].Style.BackColor = Color.Yellow;
                                        TallennaXMLButton.Enabled = false;

                                        virheita += "JulkaisunKansainvalisyysKytkin vapaaehtoinen tieto sairaaloille ja tutkimuslaitoksille. Jos muu kuin tyhjä, niin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + " Virheellinen: " + rivi.Cells[30].Value.ToString() + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                        virheIlmoitus2();

                                        virheRivi.Add(num + 1);
                                        virheSarake.Add(30);

                                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "JulkaisunKansainvalisyysKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[30].Value.ToString()); 

                                        dgvc[30].HeaderCell.Style.BackColor = Color.Yellow;
                                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                        rowStyle.BackColor = Color.Yellow;
                                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                    }

                                }


                            }
                            else if (nollaYx < 0 || nollaYx > 1)
                            {
                                XMLdataGridView.Rows[num].Cells[30].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "JulkaisunKansainvalisyysKytkin pitää olla 0 (ei) tai 1 (kyllä), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(30);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "JulkaisunKansainvalisyysKytkin pitää olla 0 (ei) tai 1 (kyllä), virheellinen", rivi.Cells[30].Value.ToString()); 

                                dgvc[30].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                        }

                    }


                    if (vuosiIlmo == 2015)
                    {
                        // AvoinSaatavuusKoodi oltava numeerinen 0 tai 1
                        if ((rivi.Cells[35] != null) || (rivi.Cells[35].Value.ToString().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[35].Value.ToString().Trim(), out nollaYx))
                            {
                                if (!OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    XMLdataGridView.Rows[num].Cells[35].Style.BackColor = Color.Red;
                                    TallennaXMLButton.Enabled = false;

                                    virheita += "AvoinSaatavuusKoodi pitää olla numeerinen arvo (0, 1, 2 tai 9), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                    virheIlmoitus2();

                                    virheRivi.Add(num + 1);
                                    virheSarake.Add(35);

                                    RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "AvoinSaatavuusKoodi pitää olla (0, 1, 2 tai 9), virheellinen", rivi.Cells[35].Value.ToString()); 

                                    dgvc[35].HeaderCell.Style.BackColor = Color.Red;
                                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                    rowStyle.BackColor = Color.Red;
                                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                }

                                if (OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    Boolean val = CheckKytkinkoodi(rivi.Cells[35].Value.ToString(), rivi.Cells[0].Value.ToString().Trim());

                                    if (!val)
                                    {
                                        XMLdataGridView.Rows[num].Cells[35].Style.BackColor = Color.Yellow;
                                        TallennaXMLButton.Enabled = false;

                                        virheita += "AvoinSaatavuusKoodi vapaaehtoinen tieto sairaaloille ja tutkimuslaitoksille. Jos muu kuin tyhjä, niin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + " Virheellinen: " + rivi.Cells[35].Value.ToString() + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                        virheIlmoitus2();

                                        virheRivi.Add(num + 1);
                                        virheSarake.Add(35);

                                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "AvoinSaatavuusKoodi pitää olla (0, 1, 2 tai 9), virheellinen", rivi.Cells[35].Value.ToString()); 

                                        dgvc[35].HeaderCell.Style.BackColor = Color.Yellow;
                                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                        rowStyle.BackColor = Color.Yellow;
                                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                    }
                                }

                            }
                            else if (nollaYx < 0 || (nollaYx > 2 && nollaYx != 9))
                            {
                                XMLdataGridView.Rows[num].Cells[35].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "AvoinSaatavuusKoodi pitää olla (0, 1, 2 tai 9), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo" + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(35);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "AvoinSaatavuusKoodi pitää olla (0, 1, 2 tai 9), virheellinen", rivi.Cells[35].Value.ToString()); 

                                dgvc[35].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                        }

                    }

                    // AvoinSaatavuusKoodi Indeksi 35 => 32 2016
                    if (vuosiIlmo == 2016 || vuosiIlmo == 2017 || vuosiIlmo == 2018)
                    {
                        // AvoinSaatavuusKoodi oltava numeerinen 0 tai 1
                        if ((rivi.Cells[32] != null) || (rivi.Cells[32].Value.ToString().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[32].Value.ToString().Trim(), out nollaYx))
                            {
                                if (!OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    XMLdataGridView.Rows[num].Cells[32].Style.BackColor = Color.Red;
                                    TallennaXMLButton.Enabled = false;

                                    virheita += "AvoinSaatavuusKoodi pitää olla numeerinen arvo (0, 1, 2 tai 9), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                    virheIlmoitus2();

                                    virheRivi.Add(num + 1);
                                    virheSarake.Add(32);

                                    RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "AvoinSaatavuusKoodi pitää olla (0, 1, 2 tai 9), virheellinen", rivi.Cells[32].Value.ToString()); 

                                    dgvc[32].HeaderCell.Style.BackColor = Color.Red;
                                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                    rowStyle.BackColor = Color.Red;
                                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                }

                                if (OnkoSairaalaTutkimus(rivi.Cells[0].Value.ToString().Trim()))
                                {
                                    Boolean val = CheckKytkinkoodi(rivi.Cells[32].Value.ToString(), rivi.Cells[0].Value.ToString().Trim());

                                    if (!val)
                                    {
                                        XMLdataGridView.Rows[num].Cells[32].Style.BackColor = Color.Yellow;
                                        TallennaXMLButton.Enabled = false;

                                        virheita += "AvoinSaatavuusKoodi vapaaehtoinen tieto sairaaloille ja tutkimuslaitoksille. Jos muu kuin tyhjä, niin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + " Virheellinen: " + rivi.Cells[32].Value.ToString() + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                        virheIlmoitus2();

                                        virheRivi.Add(num + 1);
                                        virheSarake.Add(32);

                                        RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "AvoinSaatavuusKoodi pitää olla (0, 1, 2 tai 9), virheellinen", rivi.Cells[32].Value.ToString()); 

                                        dgvc[32].HeaderCell.Style.BackColor = Color.Yellow;
                                        rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                        rowStyle.BackColor = Color.Yellow;
                                        XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                    }
                                }

                            }
                            else if (nollaYx < 0 || (nollaYx > 2 && nollaYx != 9))
                            {
                                XMLdataGridView.Rows[num].Cells[32].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "AvoinSaatavuusKoodi pitää olla (0, 1, 2 tai 9), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo" + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(32);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "AvoinSaatavuusKoodi pitää olla (0, 1, 2 tai 9), virheellinen", rivi.Cells[32].Value.ToString()); 

                                dgvc[32].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                        }
                    }


                    // JufoTunnus oltava numeerinen (4 - 5 numeroa)
                    if (((rivi.Cells[14].Value.ToString().Trim().Length > 0)))
                    {
                        int numeroX;

                        if (!Int32.TryParse(rivi.Cells[14].Value.ToString().Trim(), out numeroX))
                        {
                            XMLdataGridView.Rows[num].Cells[14].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "JufoTunnus pitää olla numeerinen arvo (4 - 5 numeroa), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(14);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "JufoTunnus pitää olla (4 - 5 numeroa) virheellinen", rivi.Cells[14].Value.ToString()); 

                            dgvc[14].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                        }
                        else if (numeroX < 1000 || numeroX > 99999 )
                        {
                            XMLdataGridView.Rows[num].Cells[14].Style.BackColor = Color.Red;
                            TallennaXMLButton.Enabled = false;

                            virheita += "JufoTunnus pitää olla (4 - 5 numeroa), rivillä " + (num + 1) + ", tunnus: " + numeroX + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                            virheIlmoitus2();

                            virheRivi.Add(num + 1);
                            virheSarake.Add(14);

                            RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "JufoTunnus pitää olla (4 - 5 numeroa) virheellinen", rivi.Cells[14].Value.ToString()); 

                            dgvc[14].HeaderCell.Style.BackColor = Color.Red;
                            rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                            rowStyle.BackColor = Color.Red;
                            XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                        }
                    }

                    if (vuosiIlmo == 2015)
                    {
                        // EVOjulkaisuKytkin oltava numeerinen  (1)
                        if ((rivi.Cells[36] != null) || (rivi.Cells[36].Value.ToString().Trim().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[36].Value.ToString().Trim(), out nollaYx) && (rivi.Cells[36].Value.ToString().Trim().Length > 0))
                            {
                                XMLdataGridView.Rows[num].Cells[36].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "EVOjulkaisuKytkin pitää olla numeerinen arvo (1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(36);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "EVOjulkaisuKytkin pitää olla 1 (kyllä) virheellinen", rivi.Cells[36].Value.ToString()); 

                                dgvc[36].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                            else if (nollaYx != 1 && (rivi.Cells[36].Value.ToString().Trim().Length > 0))
                            {
                                XMLdataGridView.Rows[num].Cells[36].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "EVOjulkaisuKytkin pitää olla 1 (kyllä), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(36);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "EVOjulkaisuKytkin pitää olla 1 (kyllä) virheellinen", rivi.Cells[36].Value.ToString()); 

                                dgvc[36].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }
                    }

                    // Tarkista DOI 
                    // Siivotaan välilyöntejä veke
                    if (vuosiIlmo == 2015)
                    {

                        if ((rivi.Cells[37].Value.ToString().Trim().Length > 0))
                        {
                            string doi = Regex.Replace(rivi.Cells[37].Value.ToString(), " ", ""); // Ei välilyöntejä	
                            rivi.Cells[37].Value = doi;

                            bool val = TarkistaDOI(rivi.Cells[37].Value.ToString());

                            if (!val)
                            {
                                XMLdataGridView.Rows[num].Cells[37].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "DOI: " + rivi.Cells[37].Value.ToString() + " virheellinen, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(37);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "DOI virheellinen", rivi.Cells[37].Value.ToString()); 

                                dgvc[37].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }

                    }


                    if (vuosiIlmo == 2016 || vuosiIlmo == 2017 || vuosiIlmo == 2018)
                    {

                        if ((rivi.Cells[36].Value.ToString().Trim().Length > 0))
                        {
                            string doi = Regex.Replace(rivi.Cells[36].Value.ToString(), " ", ""); // Ei välilyöntejä	
                            rivi.Cells[36].Value = doi;

                            bool val = TarkistaDOI(rivi.Cells[36].Value.ToString());

                            if (!val)
                            {
                                XMLdataGridView.Rows[num].Cells[36].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "DOI: " + rivi.Cells[36].Value.ToString() + " virheellinen, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(36);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "DOI virheellinen", rivi.Cells[36].Value.ToString()); 

                                dgvc[36].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }

                    }


                    // Tarkista Pysyvä osoite 
                    // Siivotaan välilyöntejä veke tuplia http ja doi/DOI
                    // Lisätään http://urn.fi jos puuttuu

                    if (vuosiIlmo == 2015)
                    {

                        if ((rivi.Cells[38].Value.ToString().Trim().Length > 0))
                        {
                            string osoite = Regex.Replace(rivi.Cells[38].Value.ToString(), " ", ""); // Ei välilyöntejä	

                            osoite = Regex.Replace(osoite, "http://http://", "http://");            // Ei tupla http://

                            osoite = Regex.Replace(osoite, "doi:", "");            // Ei doi:
                            osoite = Regex.Replace(osoite, "DOI:", "");            // Ei DOI:

                            rivi.Cells[38].Value = osoite;

                            if (osoite.Contains("URN:ISBN:") && !TarkistaOsoite(osoite))
                            {

                                string korjattu_urn_osoite = "http://urn.fi/" + osoite;

                                errorTextBox.AppendText("Korjataan pysyvä osoite: " + rivi.Cells[38].Value.ToString() + " => " + korjattu_urn_osoite + " rivillä: " + (num + 1) + "\n\r\n\r");

                                if (LokiOlemassa(masterloki))
                                {
                                    KirjoitaLokiin("Korjataan pysyvä osoite : " + rivi.Cells[38].Value.ToString() + " => " + korjattu_urn_osoite + "\n\r\n\r", masterloki);
                                }

                                rivi.Cells[38].Value = korjattu_urn_osoite;

                            }


                            bool val = TarkistaOsoite(rivi.Cells[38].Value.ToString());

                            if (!val)
                            {
                                XMLdataGridView.Rows[num].Cells[38].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "Osoite: " + rivi.Cells[38].Value.ToString() + " virheellinen, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(38);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Osoite virheellinen", rivi.Cells[38].Value.ToString()); 

                                dgvc[38].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }

                    }
                    

                    if (vuosiIlmo == 2016 || vuosiIlmo == 2017 || vuosiIlmo == 2018)
                    {
                        if ((rivi.Cells[37].Value.ToString().Trim().Length > 0))
                        {
                            string osoite = Regex.Replace(rivi.Cells[37].Value.ToString(), " ", ""); // Ei välilyöntejä	

                            osoite = Regex.Replace(osoite, "http://http://", "http://");            // Ei tupla http://

                            osoite = Regex.Replace(osoite, "doi:", "");            // Ei doi:
                            osoite = Regex.Replace(osoite, "DOI:", "");            // Ei DOI:

                            rivi.Cells[37].Value = osoite;

                            if (osoite.Contains("URN:ISBN:") && !TarkistaOsoite(osoite))
                            {

                                string korjattu_urn_osoite = "http://urn.fi/" + osoite;

                                errorTextBox.AppendText("Korjataan pysyvä osoite: " + rivi.Cells[37].Value.ToString() + " => " + korjattu_urn_osoite + " rivillä: " + (num + 1) + "\n\r\n\r");

                                if (LokiOlemassa(masterloki))
                                {
                                    KirjoitaLokiin("Korjataan pysyvä osoite : " + rivi.Cells[37].Value.ToString() + " => " + korjattu_urn_osoite + "\n\r\n\r", masterloki);
                                }

                                rivi.Cells[37].Value = korjattu_urn_osoite;
                            }

                            bool val = TarkistaOsoite(rivi.Cells[37].Value.ToString());

                            if (!val)
                            {
                                XMLdataGridView.Rows[num].Cells[37].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "Osoite: " + rivi.Cells[37].Value.ToString() + " virheellinen, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(37);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Osoite virheellinen", rivi.Cells[37].Value.ToString()); 

                                dgvc[37].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }

                    }

                    if (vuosiIlmo == 2015)
                    {
                        // Tarkista ORCID
                        if ((rivi.Cells[43].Value.ToString().Trim().Length > 0))
                        {
                            bool val = TarkistaORCIDArvot(rivi.Cells[43].Value.ToString());

                            if (!val)
                            {
                                XMLdataGridView.Rows[num].Cells[43].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "ORCID: " + rivi.Cells[43].Value.ToString() + " virheellinen, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(43);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "ORCID virheellinen", rivi.Cells[43].Value.ToString()); 

                                dgvc[43].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }
                    }
                    

                    if (vuosiIlmo == 2016 || vuosiIlmo == 2017 || vuosiIlmo == 2018)
                    {
                        // Tarkista ORCID
                        if ((rivi.Cells[42].Value.ToString().Trim().Length > 0))
                        {
                            bool val = TarkistaORCIDArvot(rivi.Cells[42].Value.ToString());

                            if (!val)
                            {
                                XMLdataGridView.Rows[num].Cells[42].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "ORCID: " + rivi.Cells[42].Value.ToString() + " virheellinen, rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";
                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(42);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "ORCID virheellinen", rivi.Cells[42].Value.ToString()); 

                                dgvc[42].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }
                    }

             
                    // 12	 AvainsanaTeksti korjataan jos on väärä välimerkki, pilkut puolipisteeksi
                    if ((rivi.Cells[12] != null) && (rivi.Cells[12].Value.ToString().Trim().Length > 0))
                    {
                        string avain_str = rivi.Cells[12].Value.ToString().Trim();

                        // Jos on vain pilkut ei puolipisteitä muutetaan pilkut puolipisteiksi
                        if (avain_str.Contains(",") && !avain_str.Contains(";"))
                        {
                            avain_str = Regex.Replace(avain_str, ",", ";"); // pilkut puolipisteeksi	

                            errorTextBox.AppendText("Korjataan Avainsanat pilkut puolipisteeksi , => ; \n\r\n\r");

                            if (LokiOlemassa(masterloki))
                            {
                                KirjoitaLokiin("Korjataan Avainsanat pilkut puolipisteeksi: " + rivi.Cells[12].Value.ToString() + " => " + avain_str + "\n\r\n\r", masterloki);
                            }

                            rivi.Cells[12].Value = avain_str;
                            XMLdataGridView.Rows[num].Cells[12].Style.BackColor = Color.Green;
                        }                      
                        
                    }


                    // YhteisjulkaisuYritysKytkin
                    if (vuosiIlmo == 2016 || vuosiIlmo == 2017 || vuosiIlmo == 2018)
                    {
                        // YhteisjulkaisuYritysKytkin oltava numeerinen 0 tai 1
                        if ((rivi.Cells[33] != null) || (rivi.Cells[33].Value.ToString().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[33].Value.ToString().Trim(), out nollaYx))
                            {
                                XMLdataGridView.Rows[num].Cells[33].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "YhteisjulkaisuYritysKytkin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(33);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuYritysKytkin 0 (ei) tai 1 (kyllä) virheellinen arvo ", rivi.Cells[33].Value.ToString()); 

                                dgvc[33].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                                Boolean val = CheckKytkinkoodi(rivi.Cells[33].Value.ToString(), rivi.Cells[0].Value.ToString().Trim());

                                if (!val)
                                {
                                    XMLdataGridView.Rows[num].Cells[33].Style.BackColor = Color.Yellow;
                                    TallennaXMLButton.Enabled = false;

                                    virheita += "YhteisjulkaisuYritysKytkin koodi rivillä " + (num + 1) + " virheellinen: " + rivi.Cells[33].Value.ToString() + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                    virheIlmoitus2();

                                    virheRivi.Add(num + 1);
                                    virheSarake.Add(33);

                                    RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuYritysKytkin 0 (ei) tai 1 (kyllä) virheellinen arvo ", rivi.Cells[33].Value.ToString()); 

                                    dgvc[33].HeaderCell.Style.BackColor = Color.Yellow;
                                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                    rowStyle.BackColor = Color.Yellow;
                                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                }
                            }
                            else if (nollaYx < 0 || nollaYx > 1)
                            {
                                XMLdataGridView.Rows[num].Cells[33].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "YhteisjulkaisuYritysKytkin pitää olla 0 (ei) tai 1 (kyllä), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(33);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "YhteisjulkaisuYritysKytkin 0 (ei) tai 1 (kyllä) virheellinen arvo ", rivi.Cells[33].Value.ToString()); 

                                dgvc[33].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }
                        }
                    }

                    // RinnakkaistallennettuKytkin
                    if (vuosiIlmo == 2016 || vuosiIlmo == 2017 || vuosiIlmo == 2018)
                    {
                        // RinnakkaistallennettuKytkin on oltava numeerinen 0 tai 1
                        if ((rivi.Cells[34] != null) || (rivi.Cells[34].Value.ToString().Length > 0))
                        {
                            int nollaYx;

                            if (!Int32.TryParse(rivi.Cells[34].Value.ToString().Trim(), out nollaYx))
                            {
                                XMLdataGridView.Rows[num].Cells[34].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "RinnakkaistallennettuKytkin pitää olla numeerinen arvo (0 tai 1), rivillä " + (num + 1) + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(34);


                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "RinnakkaistallennettuKytkin 0 (ei) tai 1 (kyllä) virheellinen arvo ", rivi.Cells[34].Value.ToString()); 

                                dgvc[34].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                                Boolean val = CheckKytkinkoodi(rivi.Cells[34].Value.ToString(), rivi.Cells[0].Value.ToString().Trim());

                                if (!val)
                                {
                                    XMLdataGridView.Rows[num].Cells[34].Style.BackColor = Color.Yellow;
                                    TallennaXMLButton.Enabled = false;

                                    virheita += "RinnakkaistallennettuKytkin koodi rivillä " + (num + 1) + " virheellinen: " + rivi.Cells[34].Value.ToString() + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                    virheIlmoitus2();

                                    virheRivi.Add(num + 1);
                                    virheSarake.Add(34);

                                    RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "RinnakkaistallennettuKytkin 0 (ei) tai 1 (kyllä) virheellinen arvo ", rivi.Cells[34].Value.ToString()); 

                                    dgvc[34].HeaderCell.Style.BackColor = Color.Yellow;
                                    rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                    rowStyle.BackColor = Color.Yellow;
                                    XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                                }
                            }
                            else if (nollaYx < 0 || nollaYx > 1)
                            {
                                XMLdataGridView.Rows[num].Cells[34].Style.BackColor = Color.Red;
                                TallennaXMLButton.Enabled = false;

                                virheita += "RinnakkaistallennettuKytkin pitää olla 0 (ei) tai 1 (kyllä), rivillä " + (num + 1) + ", numero: " + nollaYx + " on virheellinen arvo." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(34);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "RinnakkaistallennettuKytkin 0 (ei) tai 1 (kyllä) virheellinen arvo ", rivi.Cells[34].Value.ToString()); 

                                dgvc[34].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;
                            }

                            if (nollaYx == 1 && rivi.Cells[35].Value.ToString().Trim().Length == 0)
                            {
                                XMLdataGridView.Rows[num].Cells[34].Style.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].Cells[35].Style.BackColor = Color.Red;

                                TallennaXMLButton.Enabled = false;

                                virheita += "Jos rinnakkaistallennettu julkaisun arvo on 1 (kyllä), rivillä " + (num + 1) + ", pitää olla myös rinnakkaistallennetun version verkkoosoite ." + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(35);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Rinnakkaistallennetun verkko-osoite puuttuu", rivi.Cells[35].Value.ToString()); 

                                dgvc[34].HeaderCell.Style.BackColor = Color.Red;
                                dgvc[35].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }

                            if (nollaYx == 1 && !TarkistaOsoite(rivi.Cells[35].Value.ToString().Trim()))
                            {
                                XMLdataGridView.Rows[num].Cells[35].Style.BackColor = Color.Red;

                                TallennaXMLButton.Enabled = false;

                                virheita += "Rinnakkaistallennetun verkko-osoite: " + rivi.Cells[35].Value.ToString().Trim() + " virheellinen, " + virhejulkaisuorgtunnus + ".\n\r\n\r";

                                virheIlmoitus2();

                                virheRivi.Add(num + 1);
                                virheSarake.Add(35);

                                RekisteroiVirheet("VIRHE", rivi.Cells[4].Value.ToString().Trim(), (num + 1).ToString(), "Rinnakkaistallennetun verkko-osoite", rivi.Cells[35].Value.ToString()); 

                                dgvc[35].HeaderCell.Style.BackColor = Color.Red;
                                rowStyle = XMLdataGridView.Rows[num].HeaderCell.Style;
                                rowStyle.BackColor = Color.Red;
                                XMLdataGridView.Rows[num].HeaderCell.Style = rowStyle;

                            }
                        }

                    }                    


                    // Lisää tarkistuksiä ... ? Ei.

                    num++;  // Rivinro laskuri
                }
                catch (Exception ex)
                {
                    // Jahas... Miksi näin?
                    if (ex.Message.Length > 0)
                    {
                        MessageBox.Show("Validointi suoritettu!");
                    }                  
                }             
                                
            } // ...Loppu.


            if (tiedotettavaa.Trim().Length > 0)
            {
                errorTextBox.AppendText(tiedotettavaa);

                if (LokiOlemassa(masterloki))
                {
                    KirjoitaLokiin(tiedotettavaa, masterloki);
                }

            }

            // Jos on kertynyt virheilmoituksia matkalla, kaikki esiin ja lokiin...
            if (virheita.Trim().Length > 0)
            {

                // Jos on tosi paljon ei täytetä
                if (virheita.Trim().Length > 0 && virheita.Trim().Length < 2048) { 
                    
                    MessageBox.Show(virheita,"Virheitä aineistossa!");
                    errorTextBox.AppendText(virheita);

                    if (LokiOlemassa(masterloki))                    {                   

                        KirjoitaLokiin(virheita, masterloki);
                    }

                    CSV_virhe_LokiTiedosto(csvtarkistusloki);

                    KirjoitaCSVTarkistustiedosto(virheTaulukko);

                }
                else
                {                
                    if (LokiOlemassa(masterloki))
                    {
                        MessageBox.Show("Tosi paljon virheitä. Katso ihmeessä lokitiedosto", "Paljon virheitä aineistossa!");
                        errorTextBox.AppendText("Tosi paljon virheitä. Tarkista loki-tiedosto: Työkalut/Avaa lokitiedosto valikon kautta tai paina CTRL+L");

                        KirjoitaLokiin(virheita, masterloki);
                    }

                    CSV_virhe_LokiTiedosto(csvtarkistusloki);

                    KirjoitaCSVTarkistustiedosto(virheTaulukko);

                }

            }
            else
            {
                MessageBox.Show("CSV-aineisto tarkistettu! Tallenna aineisto XML-tiedostoon!", "CSV-aineisto ok.");
                
                nollaGridVarit(); // Nollaa väritykset

                TallennaXMLButton.Enabled = true;

                statusLabel.Text  = "Ok.";
                Statuslabel2.Text = "Nyt voit luoda XML-tiedoston!";

                
                numericUpDownVirhe.Enabled = false;

                virheRivi.Clear();
                virheSarake.Clear();

                if (LokiOlemassa(masterloki))
                {
                    KirjoitaLokiin("Tiedot tarkistettu. XML-tiedoston voi nyt luoda", masterloki);
                    
                    CSV_virhe_LokiTiedosto(csvtarkistusloki);
                }

            }
        }
        // Validoitu lähdedata xml-datagridistä loppuu tähän


        // Avataan samppanjapullo! :) 
        private void AvaaXMLButton_Click(object sender, EventArgs e)
        {         
            if (getXMLTiedosto().Length > 0)
            {                
                System.Diagnostics.Process.Start("notepad.exe", getXMLTiedosto());
                ValidoiXMLButton.Enabled = true;
            }
            else
            {
                MessageBox.Show("Ei voida avata!" + getXMLTiedosto() );
            }
        }


        // Kirjoitetaan csv-tarkistustiedosto
        private void KirjoitaCSVTarkistustiedosto(DataTable dt)
        {
            StreamWriter strom = new StreamWriter("tarkistusloki.csv", true);           

            int sarLkm = dt.Columns.Count;

            // Sarakkeet
            for (int i = 0; i < sarLkm; i++)
            {
                strom.Write(dt.Columns[i]);

                if (i < sarLkm )
                {
                    strom.Write(";");
                }
            }
            strom.Write(strom.NewLine);

            // virheet per rivi, sarakkeittain
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < sarLkm; i++)
                {

                    strom.Write(dr[i].ToString());
                }
                strom.Write(strom.NewLine);
            }
            strom.Close();
        }



        // Validoidaan viimeisin XML-tiedosto julkaisut.xsd scheemaa vasten
        private void ValidoiXMLButton_Click(object sender, EventArgs e)
        {           

            bool vali = false;

            if (getXMLTiedosto().Length > 0)
            {
                errorTextBox.Clear();

                vali = false;

                try
                {   // Scheman .xsd on oltava juuressa!
                    XmlSchemaSet schemaSet = new XmlSchemaSet();
                    schemaSet.Add("urn:mace:funet.fi:julkaisut/2015/03/01", "julkaisut.xsd");                    
                    
                    vali = ValidoiXMLTiedosto(getXMLTiedosto(), schemaSet);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Virhe: Voi ei! Error: " + ex.Message);
                }


                //if (LokiOlemassa(xmlvirheloki))
                //{
                //    MessageBox.Show("xml virheitä");
                //}

                if ( vali && !LokiOlemassa(xmlvirheloki))
                {
                    MessageBox.Show(UusiXMLFile + "\n\n" + " on validi.", "XML validoinnin tulos");

                    errorTextBox.Clear();

                    errorTextBox.AppendText("XML: " + UusiXMLFile + " on validi. :)");

                    if (LokiOlemassa(masterloki))
                    {
                        KirjoitaLokiin("XML: " + UusiXMLFile + " on validi.", masterloki);
                    }

                }
                else
                {

                    MessageBox.Show(UusiXMLFile + "\n\n" + " EI OLE validi!", "XML validoinnin tulos");
                    errorTextBox.Clear();

                    errorTextBox.AppendText("XML: " + UusiXMLFile + " EI OLE validi!!! :(");

                    if (LokiOlemassa(masterloki))
                    {
                        KirjoitaLokiin("XML: " + UusiXMLFile + " EI OLE validi!", masterloki);
                    }

                    // Avataan XML-virheloki
                    if (LokiOlemassa(xmlvirheloki))
                    {
                        System.Diagnostics.Process.Start("notepad.exe", xmlvirheloki);
                    }
                    else
                    {
                        MessageBox.Show("Ei voida avata " + xmlvirheloki + " tiedostoa!");
                    }



                }


            }
            else
            {
                MessageBox.Show("XML-tiedostoa " + getXMLTiedosto() + " ei voitu avata!");
            }
        }

        // Mahdolliset validointivirheet, ei löytynyt skemaa?
        private static void ValidointiVirheet(object sender, ValidationEventArgs args)
        {
            if (args.Severity == XmlSeverityType.Warning)
            {
                MessageBox.Show("XML-schemaa " + args.Message + " ei löytynyt!");
            }
            else
            {
                MessageBox.Show("XML ei ole validi!!!" + "\n\n" + args.Message);
            }
        }

        // XML-tiedoston validointi XMLReaderillä
        private static bool ValidoiXMLTiedosto(String filename, XmlSchemaSet schemaSet)
        {  
            XmlSchema compiledSchema = null;

            foreach (XmlSchema schema in schemaSet.Schemas())
            {
                compiledSchema = schema;
            }

            XmlReaderSettings settings = new XmlReaderSettings();

            settings.Schemas.Add(compiledSchema);
            settings.ValidationEventHandler += new ValidationEventHandler(ValidointiVirheet);
            settings.ValidationType = ValidationType.Schema;

            XmlTextReader reader = new XmlTextReader(filename);

            XmlReader vreader = XmlReader.Create(filename, settings);

            bool err = false;

            try { 
                     while ( vreader.Read() ) {} 
            }
            catch (Exception esx)
            {
                // MessageBox.Show("Virhe: Ei validi! Error: " + esx.Message);
                err = true;
                
                // Checkpoint

                if (LokiOlemassa(masterloki))
                {
                    KirjoitaLokiin("XML ei validi" + esx.Message, masterloki);
                }

                StreamWriter xmlvirheloki = new StreamWriter("xmlvirheloki.txt", true);
                xmlvirheloki.WriteLine("xml-error:" + esx.Message);
                xmlvirheloki.Close();

            }

            vreader.Close();

            if (err)
            {            
                if (LokiOlemassa(masterloki))              
                {
                    KirjoitaLokiin("XML " + UusiXMLFile + " ei ole validi!", masterloki);                    
                }

                return false;
            }

            else {
                
                return true;              
            }         

            
        }

        // Ohjeet Windows helppi tiedostossa "csv-xml-ohje.chm"
        private void contentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Help.ShowHelp(this, "csv-xml-ohje.chm");
            }
            catch (Exception emessa)
            {
                MessageBox.Show("Virhe: " + emessa.Message);
            }
        }

        // Jos valikon kautta valitse validoi
        private void optionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ( ValidoiButton.Enabled ) 
            {
                ValidoiButton_Click( sender, e);
            }
            else
            {
                MessageBox.Show("Ei voi validoida!");
            }

        }

        // Tietoja ohjelmasta
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox tietoja = new AboutBox();
            tietoja.ShowDialog();

        }

        // Asetukset , mikä vuosi käsitellään
        private void asetuksetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VirtaSettingsForm asetuksetForm = new VirtaSettingsForm();

            asetuksetForm.ShowDialog();


        }

        // Tyhjännä taulut ja lue uusi csv-aineisto
        private void tyhjennäVanhatTiedotJaLueUusiLähdetiedostoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
                try
                {
                    LahdeDataGridView.DataSource = null;
                    XMLdataGridView.DataSource = null;
                    tiedostoNimiLabel.Text = "";
                    errorTextBox.Clear();
                    LueLahdeTiedostoButton.Enabled = true;

                    LahdeDataGridView.Rows.Clear();
                    LahdeDataGridView.Columns.Clear();
                                       
                    XMLdataGridView.Rows.Clear();
                    XMLdataGridView.Columns.Clear();

                    virheRivi.Clear();
                    virheSarake.Clear();

                    
                    LueLahdeTiedostoButton_Click( sender,   e);


                }
                catch (Exception err)
                {
                    MessageBox.Show("Error: " + err.Message);

                }

                if (LokiOlemassa(masterloki))
                {
                    string tiedostonimi_str = tiedostoNimiLabel.Text;

                    string tiedostonimi = "";

                    if (tiedostonimi_str.Trim().Length > 0)
                    {
                        tiedostonimi = tiedostonimi_str;
                    }

                    KirjoitaLokiin("Uusi lähdetiedosto " + tiedostonimi + " luettiin.", masterloki);
                }


        }

        // Avataan masteriloki ohjelman valikon kautta
        private void avaaLokitiedostoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (LokiOlemassa(masterloki))
            {
                System.Diagnostics.Process.Start("notepad.exe", masterloki);
            }
            else
            {
                MessageBox.Show("Ei voida avata " + masterloki + " tiedostoa!");
            }

        }

        // Näytetään tarkistus csv-data Datagridissä
        private void DrawTarkistusDataGridView(List<string[]> parsedData, DataGridView dgv, int n)
        {
            dgv.Rows.Clear();
            dgv.ColumnCount = n; // Note to myself Dynaamiseksi kiitos! 

            int maximi = dgv.ColumnCount;

            try
            {
                // Nice-to-have sarakkeiden nimet luetusta csv:stä
                for (int i = 0; i < maximi; i++)
                {
                    var sb = new StringBuilder(parsedData[0][i]);
                    dgv.Columns[i].Name = sb.ToString();
                }

                foreach (string[] row in parsedData)
                {
                    dgv.Rows.Add(row);
                }

                dgv.Rows.Remove(dgv.Rows[0]);

            }
            catch (Exception e)
            {
                MessageBox.Show("Virhe tarkistus csv-aineistoa luettaessa: " + e.Message);
            }

        }

        // Avataan csv-tarkistustiedosto
        private void avaaCsvtarkistustiedostoToolStripMenuItem_Click(object sender, EventArgs e)
        {      
            try { 

                    TarkistusAineisto csvAineisto = new TarkistusAineisto();

                    csvAineisto.TarkistusDataGridView.DataSource = null;

                    int numRiveja = MontakoRivia(Path.GetFileName(csvtarkistusloki));

                    if (numRiveja > 0)
                    { 
                            OleDbConnection yhteys = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + ";Extended Properties=\"text;HDR=Yes;Format=Delimited(;)\";");

                            OleDbDataAdapter dataAdapteri = new OleDbDataAdapter("select * from [" + Path.GetFileName(csvtarkistusloki) + "]", yhteys);

                            DataTable dataTaulu = new DataTable();
                            dataAdapteri.Fill(dataTaulu);

                            csvAineisto.TarkistusDataGridView.DataSource = dataTaulu;             
                            yhteys.Close();

                            csvAineisto.TarkistusDataGridView.DataSource = null;
                            List<string[]> tarkistusCsvData = null;

                            tarkistusCsvData = LueCSVDataa( Path.GetFileName(csvtarkistusloki), ';');

                            int numSarakkeita = MontakoSaraketta( Path.GetFileName(csvtarkistusloki), ';');
                   
                            DrawTarkistusDataGridView(tarkistusCsvData, csvAineisto.TarkistusDataGridView, numSarakkeita);                   
                            csvAineisto.ShowDialog();
                    }

            }
            catch (Exception Excelex)
            {
                MessageBox.Show("Error: " + Excelex.Message);
            }          


        }


        // Avataan Excelissä muutetaan puolipiste kopion kautta sarkainerotin niin Excel osaa ryhmittää text to columns automaattisesti
        public static void AvaaCSVExcelissa()
        {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {

                MessageBox.Show("Exceliä ei ole asennettuna!");

            }
            else { 

                var ExcelApp = new Excel.Application();

                var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string line, line2;

                StreamReader file = new StreamReader(Path.GetFileName(csvtarkistusloki));
                StreamWriter file2 = new StreamWriter("tarkistusaineisto.txt", false);

                while ((line = file.ReadLine()) != null)
                {
                    line2 = line.Substring(0, line.Length -1 );
                    line2 = line2.Replace(";", "\t");               
                    file2.WriteLine(line2);                
                }

                file.Close();
                file2.Close();         

                ExcelApp.Workbooks.OpenText( path + "\\" + Path.GetFileName("tarkistusaineisto.txt")
                                                , Excel.XlPlatform.xlWindows
	                                            , 1
	                                            , Excel.XlTextParsingType.xlDelimited
                                                , Excel.XlTextQualifier.xlTextQualifierDoubleQuote                                            
	                                            , Tab:true                                    
	                                            );                                                               
                ExcelApp.Visible = true;

            }


        }

        // Virhenavigaattori, focus virheen kohdalle DataViewGridissä
        private void numericUpDownVirhe_ValueChanged(object sender, EventArgs e)
        {  
            int r = (int)((NumericUpDown)sender).Value;

            numericUpDownVirhe.Minimum = 1;
            numericUpDownVirhe.Maximum = virheRivi.Count; // max määrä virheitä jota voi hakea

            if (virheRivi.Count == 0)
            {
                numericUpDownVirhe.Enabled = false;           
            }                                  

            DataGridViewRow valittuRivi = null;     // valittu virherivi

            if (virheRivi.Count > 0 &&  r <= numericUpDownVirhe.Maximum)
            {
                numericUpDownRiviNro.Value = virheRivi[r-1];              

                int radNro = virheRivi[r-1]-1;
                int kolNro = virheSarake[r-1];

                if (radNro >= 0 && kolNro > 0 && r <= virheRivi.Count )
                {
                    try {

                        valittuRivi = XMLdataGridView.Rows[virheRivi[r-1]-1];

                        XMLdataGridView.CurrentCell = XMLdataGridView.Rows[radNro].Cells[kolNro]; // Valittu virhe (solu)
                        XMLdataGridView.BeginEdit(true);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }

        }

        
        //  Loppu slut! Stig upp å gaa ut ...

    } // ... ja kaikki hyvä myös loppuu aikanaan ...


}
