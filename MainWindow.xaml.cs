using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace hw7
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string filename = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Open_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select a File to Process";
            dialog.Filter = "Excel Files| *.xls; *.xlsx; *.xlsm";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (dialog.ShowDialog() == true) //if the user chose a file
            {
                filename = dialog.FileName;
            }
        }
        private void Menu_Exit(object sender, RoutedEventArgs e)
        {
            MessageBoxResult answer;
            answer = MessageBox.Show("Really Exit?", "", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer == MessageBoxResult.Yes)
                Application.Current.Shutdown();
        }
        private void run_Click(object sender, RoutedEventArgs e)
        {
            // variables for using Excel
            Excel.Application myApp;
            Excel.Workbook myBook;
            Excel.Worksheet mySheet;
            Excel.Range myRange;

            // connect to the Excel file data
            myApp = new Excel.Application();
            myApp.Visible = false;
            myBook = myApp.Workbooks.Open(filename);
            mySheet = myBook.Sheets[1];
            myRange = mySheet.UsedRange;
            Dictionary<string, int> units_sold = new Dictionary<string, int>();
            Dictionary<string, int> unitsPerRep = new Dictionary<string, int>();
            Dictionary<string, int> unitsPerRegion = new Dictionary<string, int>();
            Dictionary<string, int> items = new Dictionary<string, int>();
            Dictionary<string, int> region = new Dictionary<string, int>();
            Dictionary<string, double> salesRepRev = new Dictionary<string, double>();
            Dictionary<string, double> revenue = new Dictionary<string, double>();
            Dictionary<string, double> regionRev = new Dictionary<string, double>();

            string next_item;    // item on the next row, such as pencil or desk
            int next_units;        // how many items were sold
            string next_region;
            string next_rep;
            double next_revenue;

            // loop through all the rows to build the dictionary
            for (int r = 2; r <= myRange.Rows.Count; r++)
            {
                // retrieve next row's data from the Excel file
                next_item = (string)(mySheet.Cells[r, 4] as Excel.Range).Value;
                next_units = (int)(mySheet.Cells[r, 5] as Excel.Range).Value;
                next_region = (string)(mySheet.Cells[r, 2] as Excel.Range).Value;
                next_rep = (string)(mySheet.Cells[r, 3] as Excel.Range).Value;
                next_revenue = (double)(mySheet.Cells[r, 7] as Excel.Range).Value;

                // is that a new item for the dictionary?
                if (!units_sold.ContainsKey(next_item))
                {
                    units_sold.Add(next_item, next_units);
                }
                // if item already in list, then add the new values
                else
                {
                    units_sold[next_item] += next_units;
                }
                if (!revenue.ContainsKey(next_item))
                {
                    revenue.Add(next_item, next_revenue);
                }
                else
                {
                    revenue[next_item] += next_revenue;
                }
                if (!unitsPerRegion.ContainsKey(next_region))
                {
                    unitsPerRegion.Add(next_region, next_units);
                }
                else
                {
                    unitsPerRegion[next_region] += next_units;
                }
                if (!unitsPerRep.ContainsKey(next_rep))
                {
                    unitsPerRep.Add(next_rep, next_units);
                }
                else
                {
                    unitsPerRep[next_rep] += next_units;
                }
                if (!salesRepRev.ContainsKey(next_rep))
                {
                    salesRepRev.Add(next_rep, next_revenue);
                }
                else
                {
                    salesRepRev[next_rep] += next_revenue;
                }
                if (!regionRev.ContainsKey(next_region))
                {
                    regionRev.Add(next_region, next_revenue);
                }
                else
                {
                    regionRev[next_region] += next_revenue;
                }

            }
            myBook.Close();
            myApp.Quit();
            string msg = "";
            if (highVal.IsChecked == true && itemsCheck.IsChecked == true && units.IsChecked == true)
            {
                lblBig.Content = "The item with the most units sold is : " + units_sold.OrderByDescending(x => x.Value).First().Key; //high val, items, revenue
            }
            if (lowVal.IsChecked == true && itemsCheck.IsChecked == true && units.IsChecked == true)
            {
                lblBig.Content = "The item with the least units sold is : " + units_sold.OrderBy(x => x.Value).First().Key;
            }
            if (highVal.IsChecked == true && itemsCheck.IsChecked == true && revenueCheck.IsChecked == true)
            {
                lblBig.Content = "The item with the highest reveue is : " + revenue.OrderByDescending(x => x.Value).First().Key;
            }
            if (lowVal.IsChecked == true && itemsCheck.IsChecked == true && revenueCheck.IsChecked == true)
            {
                lblBig.Content = "The item with the lowest revenue is : " + revenue.OrderBy(x => x.Value).First().Key;
            }
            if (highVal.IsChecked == true && regionCheck.IsChecked == true && units.IsChecked == true)
            {
                lblBig.Content = "The region with the most units sold is : " + unitsPerRegion.OrderByDescending(x => x.Value).First().Key;
            }
            if (lowVal.IsChecked == true && regionCheck.IsChecked == true && units.IsChecked == true)
            {
                lblBig.Content = "The region with the least units sold is : " + unitsPerRegion.OrderBy(x => x.Value).First().Key;
            }
            if (highVal.IsChecked == true && repCheck.IsChecked == true && units.IsChecked == true)
            {
                lblBig.Content = "The sales rep with the most units sold is : " + unitsPerRep.OrderByDescending(x => x.Value).First().Key;
            }
            if (lowVal.IsChecked == true && repCheck.IsChecked == true && units.IsChecked == true)
            {
                lblBig.Content = "The sales rep with the least units sold is : " + unitsPerRep.OrderBy(x => x.Value).First().Key;
            }
            if (highVal.IsChecked == true && repCheck.IsChecked == true && revenueCheck.IsChecked == true)
            {
                lblBig.Content = " The sales rep with the most revenue is : " + salesRepRev.OrderByDescending(x => x.Value).First().Key;
            }
            if (lowVal.IsChecked == true && repCheck.IsChecked == true && revenueCheck.IsChecked == true)
            {
                lblBig.Content = "The sales rep with the least revenue is : " + salesRepRev.OrderBy(x => x.Value).First().Key;
            }
            if (highVal.IsChecked == true && regionCheck.IsChecked == true && revenueCheck.IsChecked == true)
            {
                lblBig.Content = "The region with the most revenue is : " + regionRev.OrderByDescending(x => x.Value).First().Key;
            }
            if (lowVal.IsChecked == true && regionCheck.IsChecked == true && revenueCheck.IsChecked == true)
            {
                lblBig.Content = "The region with the least revenue is : " + regionRev.OrderBy(x => x.Value).First().Key;
            }
            if (allVal.IsChecked == true && itemsCheck.IsChecked == true && units.IsChecked == true)
            {
                msg += "All Units Sold per Item:\n---------------------------\n";
                foreach (var item in units_sold)
                {
                    msg += "Item: " + item.Key + "      Units: " + item.Value + "\n";
                }
                lblBig.Content = msg;
            }
            if (allVal.IsChecked == true && itemsCheck.IsChecked == true && revenueCheck.IsChecked == true)
            {
                msg += "Revenue per Item:\n---------------------------\n";
                foreach (var item in revenue)
                {
                    msg += "Item: " + item.Key + "      Revenue: $" + item.Value + "\n";
                }
                lblBig.Content = msg;
            }
            if (allVal.IsChecked == true && repCheck.IsChecked == true && units.IsChecked == true)
            {
                msg += "All Units Sold per Sales Rep:\n---------------------------\n";
                foreach (var rep in unitsPerRep)
                {
                    msg += "Sales Rep: " + rep.Key + "      Units: " + rep.Value + "\n";
                }
                lblBig.Content = msg;
            }
            if (allVal.IsChecked == true && repCheck.IsChecked == true && revenueCheck.IsChecked == true)
            {
                msg += "All Revenue per Sales Rep:\n---------------------------\n";
                foreach (var rep in salesRepRev)
                {
                    msg += "Sales Rep: " + rep.Key + "      Revenue: $" + rep.Value + "\n";
                }
                lblBig.Content = msg;
            }
            if (allVal.IsChecked == true && regionCheck.IsChecked == true && units.IsChecked == true)
            {
                msg += "All Units Sold per Region:\n---------------------------\n";
                foreach (var rep in unitsPerRegion)
                {
                    msg += "Region: " + rep.Key + "      Units: " + rep.Value + "\n";
                }
                lblBig.Content = msg;
            }
            if (allVal.IsChecked == true && regionCheck.IsChecked == true && revenueCheck.IsChecked == true)
            {
                msg += "All Revenue per Region:\n---------------------------\n";
                foreach (var rep in unitsPerRegion)
                {
                    msg += "Region: " + rep.Key + " Revenue: $" + rep.Value + "\n";
                }
                lblBig.Content = msg;
            }


        }
    }
 
}
