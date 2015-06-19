using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base {
  public static class ListViewExtension {

    public static bool ToExcel(System.Windows.Forms.ListView lv) {
      //Excel.Application excel = new Excel.Application();
      object excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));

      //excel.Visible = true;
      excel.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, excel, new Object[] { true });

      //excel.Workbooks.Add(System.Reflection.Missing.Value);
      object workbooks = excel.GetType().InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, excel, null);
      object workbook = workbooks.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, workbooks, null);

      //Excel.Worksheet sheet = (Excel.Worksheet)excel.ActiveWorkbook.ActiveSheet;
      //object sheets = workbook.GetType().InvokeMember("Worksheets", System.Reflection.BindingFlags.GetProperty, null, workbook, null);
      object sheet = workbook.GetType().InvokeMember("ActiveSheet", System.Reflection.BindingFlags.GetProperty, null, workbook, null);

      string[,] dataArray = new string[lv.Items.Count + 1, lv.Columns.Count];
      List<string> fontBold = new List<string>();
      Dictionary<System.Drawing.Color, List<string>> foreColor = new Dictionary<System.Drawing.Color, List<string>>();
      Dictionary<System.Drawing.Color, List<string>> backColor = new Dictionary<System.Drawing.Color, List<string>>();

      int row = 1;
      for (int i = 0; i < lv.Columns.Count; i++) {
        dataArray[row - 1, i] = lv.Columns[i].Text;
      }

      row++;
      foreach (System.Windows.Forms.ListViewItem lvi in lv.Items) {
        for (int i = 0; i < lvi.SubItems.Count; i++) {
          dataArray[row - 1, i] = lvi.SubItems[i].Text;
          if (lvi.UseItemStyleForSubItems == false) {
            if (lvi.SubItems[i].Font.Bold == true) {
              fontBold.Add(string.Format("{0}{1}", columnIndexToLetter(i+1), row));
            }
            if (lvi.SubItems[i].ForeColor != System.Drawing.SystemColors.WindowText) {
              if (!foreColor.ContainsKey(lvi.SubItems[i].ForeColor)) {
                foreColor.Add(lvi.SubItems[i].ForeColor, new List<string>());
              }
              foreColor[lvi.SubItems[i].ForeColor].Add(string.Format("{0}{1}", columnIndexToLetter(i+1), row));
            }
            if (lvi.SubItems[i].BackColor != System.Drawing.SystemColors.Window) {
              if (!backColor.ContainsKey(lvi.SubItems[i].BackColor)) {
                backColor.Add(lvi.SubItems[i].BackColor, new List<string>());
              }
              backColor[lvi.SubItems[i].BackColor].Add(string.Format("{0}{1}", columnIndexToLetter(i+1), row));
            }
          }
        }
        row++;
      }

      if (dataArray.Length > 0) {
        //sheet.get_Range(columnIndexToLetter(1) + "1", columnIndexToLetter(lv.Columns.Count) + (row-1)).Value2 = dataArray;
        object range = sheet.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, sheet, new Object[] { columnIndexToLetter(1) + "1", columnIndexToLetter(lv.Columns.Count) + (row - 1) });
        range.GetType().InvokeMember("Value2", System.Reflection.BindingFlags.SetProperty, null, range, new Object[] { dataArray });
        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
      }

      if (fontBold.Count > 0) {
        //sheet.get_Range(listToRangeString(fontBold), System.Reflection.Missing.Value).Font.Bold = true;
        object range = sheet.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, sheet, new Object[] { listToRangeString(fontBold), System.Reflection.Missing.Value });
        object font = range.GetType().InvokeMember("Font", System.Reflection.BindingFlags.GetProperty, null, range, null);
        font.GetType().InvokeMember("Bold", System.Reflection.BindingFlags.SetProperty, null, font, new Object[] { true });
        System.Runtime.InteropServices.Marshal.ReleaseComObject(font);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
      }
      if (foreColor.Count > 0) {
        foreach (System.Drawing.Color color in foreColor.Keys) {
          //sheet.get_Range(listToRangeString(foreColor[color]), System.Reflection.Missing.Value).Font.Color = System.Drawing.ColorTranslator.ToOle(color);
          object range = sheet.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, sheet, new Object[] { listToRangeString(foreColor[color]), System.Reflection.Missing.Value });
          object font = range.GetType().InvokeMember("Font", System.Reflection.BindingFlags.GetProperty, null, range, null);
          font.GetType().InvokeMember("Color", System.Reflection.BindingFlags.SetProperty, null, font, new Object[] { System.Drawing.ColorTranslator.ToOle(color) });
          System.Runtime.InteropServices.Marshal.ReleaseComObject(font);
          System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
        }
      }
      if (backColor.Count > 0) {
        foreach (System.Drawing.Color color in backColor.Keys) {
          //sheet.get_Range(listToRangeString(backColor[color]), System.Reflection.Missing.Value).Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
          object range = sheet.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, sheet, new Object[] { listToRangeString(backColor[color]), System.Reflection.Missing.Value });
          object interior = range.GetType().InvokeMember("Interior", System.Reflection.BindingFlags.GetProperty, null, range, null);
          interior.GetType().InvokeMember("Color", System.Reflection.BindingFlags.SetProperty, null, interior, new Object[] { System.Drawing.ColorTranslator.ToOle(color) });
          System.Runtime.InteropServices.Marshal.ReleaseComObject(interior);
          System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
        }
      }

      System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
      System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
      System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
      System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

      return true;
    }

    private static string columnIndexToLetter(int dx) {
      int first, second;
      first = dx / 26;
      second = dx % 26;
      if (second == 0) {
        second = 26;
        first--;
      }
      return (first == 0 ? string.Empty : ((char)(64+first)).ToString()) + (second == 0 ? "Z" : ((char)(64+second)).ToString());
    }

    private static string listToRangeString(List<string> cells) {
      return string.Join(",", cells.ToArray());
    }
  }
}
