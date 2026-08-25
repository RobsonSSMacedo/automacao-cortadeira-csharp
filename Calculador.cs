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
                    "Este software expirou!\nEntre em contato com o suporte\npara obter nova licensa ;).\nRobson (74) 99965-3574.",
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

            // Uso de dynamic elimina a dependência rígida do assembly 'office'
            foreach (Excel.ListRow row in tabela.ListRows)
            {
                dynamic r = row.Range;
                dynamic celulaData = r.Cells[1, 1];
                dynamic celulaHora = r.Cells[1, 2];
                dynamic celulaContagem = r.Cells[1, 3];

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

            if (!encontrou)
            {
                Excel.ListRow novaLinha = tabela.ListRows.Add();
                dynamic rNova = novaLinha.Range;

                rNova.Cells[1, 1].Value = dataSerial;
                rNova.Cells[1, 2].Value = horaAtual;
                rNova.Cells[1, 3].Value = 1;
            }

            if (tabela != null) Marshal.ReleaseComObject(tabela);
            if (ws != null) Marshal.ReleaseComObject(ws);
            if (app != null) Marshal.ReleaseComObject(app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
