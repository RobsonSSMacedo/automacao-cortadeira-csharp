using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProtecaoVBA
{
    [Guid("A1B2C3D4-E5F6-4A7B-8C9D-0E1F2A3B4C5D")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICalculador { void ContaUm(object excelApp); }

    [Guid("D5C4B3A2-F6E5-4B7A-9C8D-1F0E2B3A4C5D")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Calculador : ICalculador
    {
        public void ContaUm(object excelApp)
        {
            Excel.Application app = (Excel.Application)excelApp;
            Excel.Worksheet ws = (Excel.Worksheet)app.ActiveSheet;
            
            // Tabela nomeada como "TabelaCliques" no Excel
            Excel.ListObject tabela = ws.ListObjects["TabelaCliques"];
            
            DateTime agora = DateTime.Now;
            string horaAtual = agora.ToString("HH:00 - HH:59");
            double dataSerial = Math.Floor(agora.Date.ToOADate());
            bool encontrou = false;

            // Busca na tabela
            foreach (Excel.ListRow row in tabela.ListRows)
            {
                if (Convert.ToDouble(row.Range[1, 1].Value2) == dataSerial && 
                    row.Range[1, 2].Value.ToString() == horaAtual)
                {
                    row.Range[1, 3].Value = Convert.ToDouble(row.Range[1, 3].Value) + 1;
                    encontrou = true;
                    break;
                }
            }

            // Se não encontrou, adiciona nova linha (herda formatação da tabela)
            if (!encontrou)
            {
                Excel.ListRow novaLinha = tabela.ListRows.Add();
                novaLinha.Range[1, 1].Value = dataSerial;
                novaLinha.Range[1, 2].Value = horaAtual;
                novaLinha.Range[1, 3].Value = 1;
            }

            // Limpeza de memória COM
            if (tabela != null) Marshal.ReleaseComObject(tabela);
            if (ws != null) Marshal.ReleaseComObject(ws);
            if (app != null) Marshal.ReleaseComObject(app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
