using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EconomyAnalyzer.Entitites
{
    public class Row
    {

        public Row(string s)
        {
            //Bogført;Tekst;Rentedato;Beløb;Saldo
            //24-12-2017;Dankort-nota Føtex Malmø  45321;27-12-2017;-32,25;3669,16

            if (s.StartsWith("Bog") || string.IsNullOrWhiteSpace(s))
            {
                this.Comment = string.Empty;

                return;
            }

            var columns = s.Split(';');

            this.Date = DateTime.ParseExact(columns[2], "dd-MM-yyyy", CultureInfo.InvariantCulture);
            this.Comment = columns[1];
            this.Amount = decimal.Parse(columns[3], CultureInfo.CreateSpecificCulture("da-DK"));
            this.Balance = !string.IsNullOrWhiteSpace(columns[4]) ? decimal.Parse(columns[4], CultureInfo.CreateSpecificCulture("da-DK")) : 0;
        }

        public decimal Amount { get; private set; }
        public decimal Balance { get; private set; }
        public string Comment { get; private set; }
        public string Category { get; set; }

        public string SanitizedComment
        {
            get
            {
                // Comment Example 1:
                // "Dankort-nota H&M 896         5834"

                // Comment Example 2:
                // "Visa køb DKK      97,00            VISTAPR*VistaPrint.d               Den 23.11"

                // Comment Example 3:
                // "Pengeautomat den 13.08. kl. 17.09  2344-45 eksnr. 2144"
                // "Fr. automat den 20.08. kl. 12.37   3450 eksnr.     9704"

                // Comment Example 4;
                // "Bgs egen konto"

                // Comment Example 5:
                // "Bs betaling 11571532-00021"

                // Comment Example 6:
                // "Ovf. fra, konto nr. 5678-345-124"

                // the real comment is always 14 chars long

                var result = "";

                if (string.IsNullOrWhiteSpace(this.Comment))
                {
                    result = string.Empty;
                }

                else if (this.Comment.StartsWith("Ovf."))
                {
                    result = this.Comment;
                }

                else if(this.Comment.StartsWith("Bs betaling"))
                {
                    result = $"Betalings Service: {this.Comment.Replace("Bs betaling ", "")}";
                }

                else if (this.Comment.StartsWith("Bgs"))
                {
                    result = $"Overførsel: {this.Comment.Replace("Bgs ", "")}";
                }

                else if (this.Comment.StartsWith("Dankort"))
                {
                    result = this.Comment.Replace("Dankort-nota ", "").Substring(0, 14);
                }

                else if (this.Comment.StartsWith("Pengeautomat") || this.Comment.StartsWith("Fr. automat"))
                {
                    result = "Pengeautomat";
                }

                else if (this.Comment.StartsWith("Visa"))
                {
                    result = this.Comment.Substring(35, 19);
                }
                else
                {
                    result = this.Comment;
                }

                return result.Replace(",", "-");
            }
        }

        internal string ToString(string[] categories)
        {
            var result = this.ToString();

            foreach (var category in categories)
            {
                if(this.Category == category)
                {
                    result += this.Amount * -1;
                }

                result += ",";
            }


            return result;
        }

        public DateTime Date { get; private set; }

        public override string ToString()
        {
            if(this.Date == null)
            {
                return string.Empty;
            }

            return $"{Date.ToString("yyyy-MM-dd")},{SanitizedComment.Trim()},{Amount},{Balance},";
        }

    }
}
