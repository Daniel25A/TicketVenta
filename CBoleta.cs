using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
//using System.Drawing.Printing;
using System.Runtime.InteropServices;
namespace MyTicket
{
    /*
     Autor: Daniel25A ( Oscar Gomez )
     Fecha de Escritura: 09/08/2019
     ¿Hay Código Importado? : Si, Para la Impresión del Ticket Si(Libreria de Microsoft), Pero la Logica del codigo del ticket esta 100% Escrita por el autor Mencionado..
     */
    enum TipodeLinea
    {
        Guion,
        Igual,
        Asterisco
    }
    class CBoleta
    {
        StringBuilder textoTicket = new StringBuilder();
        Int32 MaximoLetras=40;
        string nombreComercio;
        string nombreCajero;
        string numeroFactura;
        static string DocNameifDontPrint = "Ticket";
        public CBoleta(string pComercio,string pCajero,string pNumeroFactura)
        {
            if (pComercio.Length > MaximoLetras) {
                this.nombreComercio = pComercio.Substring(0, MaximoLetras - 1) + "E";
            }
            else
            {
                this.nombreComercio = pComercio;
            }
            if (pCajero.Length > MaximoLetras)
            {
                this.nombreCajero = pCajero.Substring(0, MaximoLetras - 1) + "E";
            }
            else
            {
                this.nombreCajero = pCajero;
            }
            this.numeroFactura = pNumeroFactura;
            textoTicket.AppendLine(nombreComercio);
            textoTicket.AppendLine(nombreCajero);
            textoTicket.AppendLine(numeroFactura);
        }
        public void AgregarEncabezadoTicket()
        {
            textoTicket.AppendLine("DESC.              |MONTO.     |CANT.   ");
        }
        public void AgregarArticulo(string ArtName, double ArtPrecio, double ArtCantidad)
        {
            Int32 MaxCaracterArt = 19;
            Int32 MaxCaracterPrecio = 11;
            Int32 MaxCaracterCant = 8;
            string NombreArticulo = ArtName;
            string PrecioArticulo = ArtPrecio.ToString();
            string CantidadArticulo = ArtCantidad.ToString();
            string TextoRestanteArt = "";
            if (NombreArticulo.Length > MaxCaracterArt)
            {
                NombreArticulo = NombreArticulo.Substring(0, MaxCaracterArt);
                TextoRestanteArt = ArtName.Substring(MaxCaracterArt, ArtName.Length - MaxCaracterArt);
            }
            else if (NombreArticulo.Length < MaxCaracterArt)
            {
                for (int i = 0; i < MaxCaracterArt-ArtName.Length; i++)
                {
                    NombreArticulo += " ";
                }
            }
            if (PrecioArticulo.Length > MaxCaracterPrecio)
            {
                PrecioArticulo = PrecioArticulo.Substring(0, MaxCaracterPrecio - 1) + "E";
            }
            else if (PrecioArticulo.Length < MaxCaracterPrecio)
            {
                for (int i = 0; i < MaxCaracterPrecio - ArtPrecio.ToString().Length; i++)
                {
                    PrecioArticulo += " ";
                }
            }
            if (CantidadArticulo.Length > MaxCaracterCant)
            {
                CantidadArticulo = CantidadArticulo.Substring(0, MaxCaracterCant - 1) + "E";
            }
            if(TextoRestanteArt!=string.Empty){
                textoTicket.AppendLine(string.Format("{0} {1} {2}", NombreArticulo, PrecioArticulo, CantidadArticulo));
                textoTicket.AppendLine(TextoRestanteArt);
            }
            else
                textoTicket.AppendLine(string.Format("{0} {1} {2}", NombreArticulo, PrecioArticulo, CantidadArticulo, TextoRestanteArt));
        }
        public void AgregarLineas(TipodeLinea Tipo)
        {
            string signo="";
            string Lineas = "";
            switch (Tipo)
            {
                case TipodeLinea.Guion:
                    signo = "-";
                    break;
                case TipodeLinea.Igual:
                    signo = "=";
                    break;
                case TipodeLinea.Asterisco:
                    signo = "*";
                    break;
                default:
                    break;
            }
            for (int i = 0; i < MaximoLetras; i++)
            {
                Lineas += signo;
            }
            textoTicket.AppendLine(Lineas);
        }
        public void AgregarTotales(decimal pTotal, string pSigno = "$",string pTexto="Total")
        {
            string ValorTotal;
            string TituloTotal;
            try
            {
                if (pTotal > 0)
                {
                    ValorTotal = pTotal.ToString("###,###,###.00") + pSigno;
                    TituloTotal = pTexto;
                    if ((ValorTotal.Length + TituloTotal.Length) > MaximoLetras)
                    {
                        throw new Exception();
                    }
                    else if ((ValorTotal.Length + TituloTotal.Length) < MaximoLetras)
                    {
                        for (int i = 0; i < (MaximoLetras - (ValorTotal.Length+pTexto.Length)); i++)
                        {
                            TituloTotal += " ";
                        }
                    }
                    textoTicket.AppendLine(string.Format("{0}{1}", TituloTotal, ValorTotal));
                }
            }
            catch (Exception ex)
            {
                textoTicket.AppendLine("Se Excedio el Espacio");
            }
        }
        public void Imprimir(string Impresora)
        {
            Console.WriteLine(textoTicket.ToString());
            Console.WriteLine(RawPrinterHelper.SendStringToPrinter(Impresora, textoTicket.ToString()) ? "Se Imprimio la Boleta" : "No se Imprimio la Boleta");
            textoTicket.Clear();
        }
        public class RawPrinterHelper
        {
            // Structure and API declarions:
            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
            public class DOCINFOA
            {
                [MarshalAs(UnmanagedType.LPStr)]
                public string pDocName;
                [MarshalAs(UnmanagedType.LPStr)]
                public string pOutputFile;
                [MarshalAs(UnmanagedType.LPStr)]
                public string pDataType;
            }
            [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

            [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool ClosePrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

            [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool EndDocPrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool StartPagePrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool EndPagePrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);
            public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
            {
                Int32 dwError = 0, dwWritten = 0;
                IntPtr hPrinter = new IntPtr(0);
                DOCINFOA di = new DOCINFOA();
                bool bSuccess = false;

                di.pDocName = DocNameifDontPrint;
                di.pDataType = "RAW";


                if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
                {

                    if (StartDocPrinter(hPrinter, 1, di))
                    {
                        // Start a page.
                        if (StartPagePrinter(hPrinter))
                        {
                            // Write your bytes.
                            bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                            EndPagePrinter(hPrinter);
                        }
                        EndDocPrinter(hPrinter);
                    }
                    ClosePrinter(hPrinter);
                }

                if (bSuccess == false)
                {
                    dwError = Marshal.GetLastWin32Error();
                }
                return bSuccess;
            }

            public static bool SendStringToPrinter(string szPrinterName, string szString)
            {
                IntPtr pBytes;
                Int32 dwCount;

                dwCount = szString.Length;

                pBytes = Marshal.StringToCoTaskMemAnsi(szString);

                SendBytesToPrinter(szPrinterName, pBytes, dwCount);
                Marshal.FreeCoTaskMem(pBytes);
                return true;
            }
        }
    }
}
