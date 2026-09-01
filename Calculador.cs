using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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

            // 🔒 TRAVA DE SEGURANÇA: Bloqueia a partir de 01/01/2027
            if (agora >= new DateTime(2027, 1, 1))
            {
                MessageBox.Show(
                    "Este software expirou em 01/01/2027.\nRobson (74) 99965-3574.", 
                    "Sistema Expirado", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning
                );
                return;
            }

            try
            {
                dynamic app = excelApp;
                dynamic ws = app.ActiveSheet;
                
                string horaAtual = agora.ToString("HH:00 - HH:59");
                double dataHojeSerial = Math.Floor(agora.Date.ToOADate());
                bool encontrou = false;

                int ultimaLinha = ws.Cells[ws.Rows.Count, 1].End(-4162).Row;
                if (ultimaLinha < 2) ultimaLinha = 1;

                for (int i = 2; i <= ultimaLinha; i++)
                {
                    var valorCelulaA = ws.Cells[i, 1].Value;
                    string valorHora = Convert.ToString(ws.Cells[i, 2].Value);

                    if (valorCelulaA != null && valorHora == horaAtual)
                    {
                        double dataPlanilhaSerial = 0;

                        if (valorCelulaA is double || valorCelulaA is int || valorCelulaA is long)
                        {
                            dataPlanilhaSerial = Math.Floor(Convert.ToDouble(valorCelulaA));
                        }
                        else
                        {
                            DateTime dataConvertida;
                            if (DateTime.TryParse(Convert.ToString(valorCelulaA), out dataConvertida))
                            {
                                dataPlanilhaSerial = Math.Floor(dataConvertida.Date.ToOADate());
                            }
                        }

                        if (dataPlanilhaSerial == dataHojeSerial)
                        {
                            double valorAtual = 0;
                            double.TryParse(Convert.ToString(ws.Cells[i, 3].Value), out valorAtual);
                            ws.Cells[i, 3].Value = valorAtual + 1;
                            encontrou = true;
                            break;
                        }
                    }
                }

                if (!encontrou)
                {
                    int novaLinha = (ultimaLinha < 2 && Convert.ToString(ws.Cells.Value) == "") ? 2 : ultimaLinha + 1;
                    
                    ws.Cells[novaLinha, 1].Value = dataHojeSerial;
                    ws.Cells[novaLinha, 1].NumberFormat = "dd/mm/aaaa"; 
                    
                    ws.Cells[novaLinha, 2].Value = horaAtual;
                    ws.Cells[novaLinha, 3].Value = 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro na integração com a tela do Excel: " + ex.Message);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
