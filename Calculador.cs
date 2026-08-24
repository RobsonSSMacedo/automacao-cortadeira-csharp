using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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
            DateTime agora = DateTime.Now;

            // Trava de segurança para 2027
            if (agora.Year >= 2027)
            {
                MessageBox.Show(
                    "Este software expirou em 01/01/2027.\nRobson (74) 99965-3574.", 
                    "Sistema Expirado", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning
                );
                return;
            }

            Excel.Application app = (Excel.Application)excelApp;
            Excel.Worksheet ws = (Excel.Worksheet)app.ActiveSheet;
            Excel.ListObject tabela = ws.ListObjects["TabelaCliques"];
            
            string horaAtual = agora.ToString("HH:00 - HH:59");
            double dataSerial = Math.Floor(agora.Date.ToOADate());
            bool encontrou = false;

            // Percorre as linhas convertendo explicitamente cada celula para o tipo correto do Excel
            foreach (Excel.ListRow row in tabela.ListRows)
            {
                Excel.Range r = (Excel.Range)row.Range;
                Excel.Range celulaData = (Excel.Range)r.Cells[1, 1];
                Excel.Range celulaHora = (Excel.Range)r.Cells[1, 2];
                Excel.Range celulaContagem = (Excel.Range)r.Cells[1, 3];

                if (celulaData.Value2 != null && celulaHora.Value != null)
                {
                    if (Convert.ToDouble(celulaData.Value2) == dataSerial && 
                        celulaHora.Value.ToString() == horaAtual)
                    {
                        celulaContagem.Value = Convert.ToDouble(celulaContagem.Value) + 1;
                        encontrou = true;
                        break;
                    }
                }
            }

            // Se não encontrou, adiciona nova linha e insere os dados iniciais
            if (!encontrou)
            {
                Excel.ListRow novaLinha = tabela.ListRows.Add();
                Excel.Range rNova = (Excel.Range)novaLinha.Range;
                
                ((Excel.Range)rNova.Cells[1, 1]).Value = dataSerial;
                ((Excel.Range)rNova.Cells[1, 2]).Value = horaAtual;
                ((Excel.Range)rNova.Cells[1, 3]).Value = 1;
            }

            // Limpeza manual de memoria COM
            if (tabela != null) Marshal.ReleaseComObject(tabela);
            if (ws != null) Marshal.ReleaseComObject(ws);
            if (app != null) Marshal.ReleaseComObject(app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
