using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace bamEpplus
{
    class LigneContenu : Ligne
    {
    public string lire_ligne(String ligne, Entete entete,int indexe)
        {
            try
            {
               // Console.WriteLine(ligne);

                this.liste = new List<Donnee>();
                 if (entete.delimiteur.CompareTo('i') == 0)
                 {
                     foreach (Donnee entetedonnee in entete.liste)
                     {
                         Donnee donne = new Donnee { libelle = entetedonnee.libelle, contenu = ligne.Substring(entetedonnee.index_debut-1, entetedonnee.longueur) };
                         this.liste.Add(donne);
                     }
                 }
                 else {
                     String[] mots = ligne.Split(entete.delimiteur);
                     foreach (Donnee entetedonnee in entete.liste)
                     {
                         Donnee donne = new Donnee { libelle = entetedonnee.libelle, contenu = mots[entetedonnee.index] };
                         this.liste.Add(donne);
                     }
                 }
                //System.Console.WriteLine(liste);
                //Console.WriteLine("hello");
                //System.Console.ReadKey();
            }


            catch (Exception ex)
            {
                return ex.Message;
            }
            return "bien";
        }
    public override void ecrire_ligne(int index, ExcelWorksheet worksheet)
        {
            int i = 1;
             foreach (Donnee element in this.liste) {
                 worksheet.Cells[index, i].StyleID = worksheet.Cells[4, i].StyleID;
                 worksheet.Cells[index, i].Value = element.contenu;
                 /*worksheet.Cells[index, i].Style.Font.Bold = worksheet.Cells[4, i].Style.Font.Bold;
                 //worksheet.Cells[index, i].Style.Font.Color.SetColor(worksheet.Cells[4, i].Style.Font.Color);
                 //worksheet.Cells[index, i].Style.VerticalAlignment = worksheet.Cells[4, i].Style.VerticalAlignment;
                 //worksheet.Cells[index, i].Style.WrapText = worksheet.Cells[4, i].Style.WrapText;
                 //worksheet.Cells[index, i].Style.HorizontalAlignment = worksheet.Cells[4, i].Style.HorizontalAlignment;
                 //worksheet.Cells[index, i].Style.Fill.PatternType = worksheet.Cells[4, i].Style.Fill.PatternType;
                 //worksheet.Cells[index, i].Style.Font = worksheet.Cells[4, i].Style.Font;*/
                 i++;
            }
          
            
        }
    public void affiche()
    {
        foreach (var donne in this.liste)
        {
            Console.WriteLine("Identificateur: " + this.Identificateur );
            Console.WriteLine(donne.affiche());


        }
    }
}
}
