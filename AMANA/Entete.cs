using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using OfficeOpenXml;
using OfficeOpenXml.Style;



using System.IO;
using System.Drawing;
using System.Xml.Linq;

namespace bamEpplus
{
    class Entete : Ligne
    {
       public String type;




        public string lire_carac(String file)
        {   
            try
            {
                StreamReader monStreamReader = new StreamReader(@file);
                string ligne = ""; String[] mots; int k = 0;// declaration des variables neccessaires
                for (int i = 0; i < this.nombre_element; i++)
                {

                    Donnee donne;
                    ligne = monStreamReader.ReadLine();
                    mots = ligne.Split(',');
                    if (this.delimiteur.CompareTo('i') == 0)
                    {
                        donne = new Donnee { libelle = mots[0], type = "index", index_debut = Convert.ToInt32(mots[1]), index_fin = Convert.ToInt32(mots[2]), width = Convert.ToInt32(mots[3]) };
                    }
                    else
                    {
                        donne = new Donnee { libelle = mots[0], type = "virgule", index = k, width = Convert.ToDouble(mots[2]) };
                        k++;
                    }
                    if (this.liste == null) { this.liste = new List<Donnee>(); }
                   // Console.WriteLine(donne);/*****************************/
                    this.liste.Add(donne);

                }

                monStreamReader.Close();
            }


            catch (Exception ex)
            {
                return ex.Message;
            }
            return "bien";
        }


        public override void ecrire_ligne(int index, ExcelWorksheet worksheet)
        {
            int i = 3;
            worksheet.Row(i).Height= 30;
            var range = worksheet.Cells[index, index, 1, nombre_element];
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.ShrinkToFit = false;
            range.Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
            range.Style.Font.Color.SetColor(Color.Black);
            foreach (Donnee element in this.liste)  
            {
               // worksheet.Column(i).Width = 5*element.width;
                worksheet.Cells[index, i].Value = element.libelle;
                
                i++;
            }


        }

        public string lire_XML(XElement nodes)
        { //IEnumerable<XElement> nodes
            try
            {
                var elements = from element in nodes.Elements("ELEMENTS").Elements()

                               select element;

                int i = 1;
                foreach (var element in elements)
                {

                    Donnee donne;

                    donne = new Donnee
                    {
                        libelle = (string)element.Attribute("ID"),
                        type = "index",
                        index = i,
                        index_debut = Convert.ToInt32((string)element.Attribute("INDDB")),
                        index_fin = Convert.ToInt32((string)element.Attribute("INDFN")),
                        longueur = Convert.ToInt32((string)element.Attribute("LONGUEUR"))
                    };

                    if (this.liste == null) 
                        { 
                            this.liste = new List<Donnee>();
                        }
                    this.liste.Add(donne);
                    i++;
                }

            }


            catch (Exception ex)
            {
                return ex.Message;
            }
            return "bien";
        }


        public  void affiche()
        {
            foreach (var donne in this.liste) {
                Console.WriteLine("Identificateur: "+this.Identificateur+"\t type: "+this.type+"");
                Console.WriteLine(donne.affiche());
            
            
            }
        }
    }
}
