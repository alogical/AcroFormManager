###############################################################################
### Icon Extraction C# Definition :: Shell32.dll Function Import
###############################################################################
$Csharp = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@

Add-Type -TypeDefinition $Csharp -ReferencedAssemblies System.Drawing

#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#!! TODO !! ID=0006
#!! Add support for sorting the Size column as integers.
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
###############################################################################
### ListViewSorter
###############################################################################
$Csharp = @"
using System;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ListViewSorter
{
    public class ItemComparer : IComparer
    {
        private int col;
        private SortOrder order;
        public ItemComparer()
        {
            col = 0;
            order = SortOrder.Ascending;
        }
        
        public ItemComparer(int column, SortOrder order)
        {
            col = column;
            this.order = order;
        }
        
        public int Compare(object x, object y)
        {
            int returnVal;
            
            // Determine if items being sorted are dates
            try
            {
                // Parse the two objects passed as parameters as a DateTime
                System.DateTime DateX = DateTime.Parse(((ListViewItem)x).SubItems[col].Text);
                System.DateTime DateY = DateTime.Parse(((ListViewItem)y).SubItems[col].Text);
                
                // Compare the two dates.
                returnVal = DateTime.Compare(DateX, DateY);
            }
            // If the compared objects don't have a valid date format, compare as string.
            catch
            {
                returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
            }
            
            // Determine if sort order is descending.
            if(order == SortOrder.Descending)
            {
                // Invert the return value
                returnVal *= -1;
            }
            
            return returnVal;
        }
    }
}
"@

$Assem = (
    "System.Windows.Forms",
    "System.Drawing"
    )
    
Add-Type -TypeDefinition $Csharp -ReferencedAssemblies $Assem

Remove-Variable Csharp
Remove-Variable Assem