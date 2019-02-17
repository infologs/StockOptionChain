using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using HtmlAgilityPack;
using System.Data.SqlClient;


namespace ReadOptionChain
{
    class Program
    {

        public double RemoveChar(string value)
        {
            try
            {
                return Convert.ToDouble(value.Replace(",", ""));
            } catch(Exception e) {
                return Convert.ToDouble(value.Replace(value.Substring(value.IndexOf("<!--"), (value.IndexOf("-->") + 3)), ""));
            }
        }
        static void Main(string[] args)
        {
            Program p = new Program();
            WebRequest request = WebRequest.Create(@"https://nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?segmentLink=17&instrument=OPTIDX&symbol=NIFTY&date=28FEB2019");

            WebResponse response = request.GetResponse();

            Stream dataStream = response.GetResponseStream();

            StreamReader reader = new StreamReader(dataStream);

            string rt = reader.ReadToEnd();

            Console.WriteLine(rt);

            File.WriteAllText("d:\\abhishek.txt", rt);

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(rt);



            List<List<string>> table = doc.DocumentNode.SelectSingleNode("//table[@id='octable']")
            .Descendants("tr")
            .Skip(1)
            .Where(tr => tr.Elements("td").Count() > 1)
            .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
            .ToList();


            SqlConnection cn = new SqlConnection(@"Data Source=(LocalDB)\v11.0;AttachDbFilename='C:\Users\Abhishek\documents\visual studio 2013\Projects\ReadOptionChain\ReadOptionChain\Database1.mdf';Integrated Security=True");
            /*SqlCommand cmd = new SqlCommand("INSERT INTO MasterStockList VALUES (@Code, @Name, @AddedDate)",cn);
            cmd.Parameters.AddWithValue("@Code", "NIFTY");
            cmd.Parameters.AddWithValue("@Name", "NIFTY");
            cmd.Parameters.AddWithValue("@AddedDate", DateTime.Now);
            cn.Open();
            cmd.ExecuteNonQuery();
            cn.Close();

            cmd = new SqlCommand("INSERT INTO TimeEntry VALUES (@StockID, @AddedDate)",cn);
            cmd.Parameters.AddWithValue("@StockID", 1);
            cmd.Parameters.AddWithValue("@AddedDate", DateTime.Now);
            cn.Open();
            cmd.ExecuteNonQuery();
            cn.Close();
            */

            string _entry = "";
            SqlCommand cmd = null;

            DateTime dt = DateTime.Now;
            
            string dateTime = DateTime.Now.ToString();

            for (int i = 0; i < table.Count - 1; i++)
            {
                cmd = new SqlCommand("INSERT INTO DetailData VALUES (@StrikePrice, @StockPrice, @OI, @ChangeOI, @Volumne, @IV, @LTP, @NetChange, @BidQty, @BidPrice, @AskPrice, @AskQty, @Type, @AddedDate)");

                string StockPrices = doc.DocumentNode.SelectSingleNode("//b").InnerText;
                double price = Convert.ToDouble(StockPrices.Replace("NIFTY ", "").ToString());

                cmd.Parameters.AddWithValue("@StockPrice", price);
                cmd.Connection = cn;

                cmd.Parameters.AddWithValue("@StrikePrice", p.RemoveChar(table[i][11]));
                cmd.Parameters.AddWithValue("@OI", p.RemoveChar(table[i][1] == "-" ? "0" : table[i][1]));
                cmd.Parameters.AddWithValue("@ChangeOI", p.RemoveChar(table[i][2] == "-" ? "0": table[i][2]));
                cmd.Parameters.AddWithValue("@Volumne", p.RemoveChar(table[i][3] == "-" ? "0" : table[i][3]));
                cmd.Parameters.AddWithValue("@IV", p.RemoveChar(table[i][4] == "-" ? "0" : table[i][4]));
                cmd.Parameters.AddWithValue("@LTP", p.RemoveChar(table[i][5] == "-" ? "0" : table[i][5]));
                cmd.Parameters.AddWithValue("@NetChange", p.RemoveChar(table[i][6] == "-" ? "0" : table[i][6]));
                cmd.Parameters.AddWithValue("@BidQty", p.RemoveChar(table[i][7] == "-" ? "0" : table[i][7]));
                cmd.Parameters.AddWithValue("@BidPrice", p.RemoveChar(table[i][8] == "-" ? "0" : table[i][8]));
                cmd.Parameters.AddWithValue("@AskPrice", p.RemoveChar(table[i][9] == "-" ? "0" : table[i][9]));
                cmd.Parameters.AddWithValue("@AskQty", p.RemoveChar(table[i][10] == "-" ? "0" : table[i][10]));
                cmd.Parameters.AddWithValue("@Type", 1);
                cmd.Parameters.AddWithValue("@AddedDate",dt );

                cn.Open();
                cmd.ExecuteNonQuery();
                cn.Close();

                cmd = new SqlCommand("INSERT INTO DetailData VALUES (@StrikePrice, @StockPrice, @OI, @ChangeOI, @Volumne, @IV, @LTP, @NetChange, @BidQty, @BidPrice, @AskPrice, @AskQty, @Type, @AddedDate)");
                cmd.Parameters.AddWithValue("@StockPrice", price);
                cmd.Connection = cn;

                cmd.Parameters.AddWithValue("@StrikePrice", p.RemoveChar(table[i][11] == "-" ? "0" : table[i][11]));
                cmd.Parameters.AddWithValue("@OI", p.RemoveChar(table[i][21] == "-" ? "0" : table[i][21]));
                cmd.Parameters.AddWithValue("@ChangeOI", p.RemoveChar(table[i][20] == "-" ? "0" : table[i][20]));
                cmd.Parameters.AddWithValue("@Volumne", p.RemoveChar(table[i][19] == "-" ? "0" : table[i][19]));
                cmd.Parameters.AddWithValue("@IV", p.RemoveChar(table[i][18] == "-" ? "0" : table[i][18]));
                cmd.Parameters.AddWithValue("@LTP", p.RemoveChar(table[i][17] == "-" ? "0" : table[i][17]));
                cmd.Parameters.AddWithValue("@NetChange", p.RemoveChar(table[i][16] == "-" ? "0" : table[i][16]));
                cmd.Parameters.AddWithValue("@AskQty", p.RemoveChar(table[i][15] == "-" ? "0" : table[i][15]));
                cmd.Parameters.AddWithValue("@AskPrice", p.RemoveChar(table[i][14] == "-" ? "0" : table[i][14]));
                cmd.Parameters.AddWithValue("@BidPrice", p.RemoveChar(table[i][13] == "-" ? "0" : table[i][13]));
                cmd.Parameters.AddWithValue("@BidQty", p.RemoveChar(table[i][12] == "-" ? "0" : table[i][12]));
                                
                cmd.Parameters.AddWithValue("@Type", 2);                
                cmd.Parameters.AddWithValue("@AddedDate", dt);

                cn.Open();
                cmd.ExecuteNonQuery();
                cn.Close();
            }

            reader.Close();
            response.Close();
        }
    }
}
