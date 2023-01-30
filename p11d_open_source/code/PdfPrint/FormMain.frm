VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Function IsPDFPrinterDriverEx(ByVal printerName As String, ByVal ABACUS_PRINTER_DRIVER As String) As Boolean
  IsPDFPrinterDriverEx = StrComp(printerName, ABACUS_PRINTER_DRIVER) = 0
End Function
Private Function IsPDFAvailable() As Boolean
  Dim p As Printer
  Dim i As Long
    
  If (m_PDFPrinterAvailable = NOT_SEARCHED) Then
    m_PDFPrinterAvailable = NOT_AVAILABLE
    For Each p In Printers
      For i = 1 To UBound(PDF_DRIVER_NAMES)
        If IsPDFPrinterDriverEx(p.DeviceName, PDF_DRIVER_NAMES(i)) Then
          m_PDFPrinterAvailable = AVAILABLE
          m_PDFPrinterName = p.DeviceName 'PDF_DRIVER_NAMES(i)
        End If
      Next
    Next
  End If
  IsPDFAvailable = m_PDFPrinterAvailable = AVAILABLE
End Function

Public Function IsCurrentPrinterPDF() As Boolean
  Dim i As Long
  Dim s As String
  
  s = Printer.DeviceName
  For i = 1 To UBound(PDF_DRIVER_NAMES)
    If IsPDFPrinterDriverEx(s, PDF_DRIVER_NAMES(i)) Then
      IsCurrentPrinterPDF = True
      Exit Function
    End If
  Next

End Function

Public Function PDFDriverInstall() As Boolean
  Dim serialNumber As String
  
  On Error GoTo err_Err
  
  'cad pdf start
  'PDFDriverInstall = False
  'Exit Function
  'cad pdf end
  If (Not IsCurrentPrinterPDF()) Then GoTo err_End
  If IsPDFAvailable = False Then
    Call Err.Raise(ERR_ERROR, "Reporter", "Pdf driver is not available")
  End If
    
  If (g_cdi Is Nothing) Then
    'Printer.Orientation = 1 'EWPDF
    Set g_cdi = New CDIntfEx.CDIntfEx
  Else
    Call Err.Raise(ERR_ERROR, "Reporter", "Can not install pdf driver as has already been installed")
  End If
                  
  serialNumber = "07EFCDAB01000100E0370FBEA11533CE8484F6E47E12547615123438D593268E13F82AE12E446D05067F24A3DB404331D45D96E74588"
  If IsPDFPrinterDriverEx(m_PDFPrinterName, PDF_DRIVER_ABACUS) Or IsPDFPrinterDriverEx(m_PDFPrinterName, PDF_DRIVER_ONE_SOURCE) Then
    g_cdi.DriverInit m_PDFPrinterName 'PDF_DRIVER_ABACUS
    g_cdi.EnablePrinter "Thomson Reuters (Professional)", serialNumber '"07EFCDAB01000100BC59AEFEBF38B9649CF0EE64C144C982C50A46D02B2B3976AC9E4DA9C88520E529023E1508C89DECE79FA977A0FA6D486C3C40BC75D47E"
  ElseIf IsPDFPrinterDriverEx(m_PDFPrinterName, PDF_DRIVER_SAGE) Then
    g_cdi.DriverInit m_PDFPrinterName
    g_cdi.EnablePrinter "Sage (UK) Limited", serialNumber '"07EFCDAB0100010062AE5DBE1E7D10232F4872A235B5E34B5EFC7E6D2EE27CCF7ACBCBBB214A21D8F384C17CF890221DF0893AFD22EE4DDE0621AD1D1921F4"
  Else
      Call Err.Raise(1, "PDF Driver Install", "Invalid PDF Printer Name:" & m_PDFPrinterName)
  End If
  
  
  g_cdi_horizontal_margin = g_cdi.HorizontalMargin
  g_cdi_vertical_margin = g_cdi.VerticalMargin
  g_cdi_paper_size = g_cdi.PaperSize
  
  g_cdi.HorizontalMargin = A4NonPrintableMicroMeters
  g_cdi.VerticalMargin = A4NonPrintableMicroMeters
  g_cdi.PaperSize = 9 'A4
  PDFDriverInstall = True
  
err_End:
  Exit Function
err_Err:
  Call Err.Raise(Err.Number, "PDFDriverInstall", Err.Description)
End Function

Public Function A4NonPrintableMicroMeters() As Single
  A4NonPrintableMicroMeters = 60#
End Function

Public Function A4NonPrintableMArginTwips() As Single
  A4NonPrintableMArginTwips = (1440 * (A4NonPrintableMicroMeters() / 254)) '340.157480315
End Function

Public Function TextFileLoad(ByVal sPathAndFile As String) As String
  Dim fr As TCSFileread
  Dim s As String
  
  On Error GoTo err_Err
  Set fr = New TCSFileread
  
  If Not fr.OpenFile(sPathAndFile) Then Call Err.Raise(0, "TextFileLoad", "Failed to open file " & sPathAndFile)
  
  
  Call fr.GetFile(s)
  TextFileLoad = s
  
  
err_End:
  Exit Function
err_Err:
  Call Err.Raise(0, "TextFileLoad", Err.Description)
  
End Function

