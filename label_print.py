# Tamanho da etiqueta: ETIQUETA ADESIVA BOPP PEROLADO 100 x 150mm - COUCHE PERSONALIZADA 250 UNID.

from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.units import mm, inch
from reportlab.lib.styles import getSampleStyleSheet
from get_data import LabelInfo
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from reportlab.graphics.barcode import code128
from reportlab.graphics.barcode import code93
from reportlab.graphics.barcode import code39
from reportlab.graphics.barcode import usps
from reportlab.graphics.barcode import usps4s
from reportlab.graphics.barcode import ecc200datamatrix
from reportlab.platypus import SimpleDocTemplate, Paragraph

from win32 import win32print, win32api
import os

class LabelPrint():
  def __init__(self, LabelInfo) -> None:
    self.label_info = LabelInfo

  def get_string_width(self, text, font_name, font_size) -> pdfmetrics.stringWidth:
    return pdfmetrics.stringWidth(text, font_name, font_size)
  # Função, para adaptar o texto a comprimento e largura do pdf.
  def draw_text(self, canvas, text, x, y, max_width, initial_font_size, font_name="Helvetica") -> None:
    font_size = initial_font_size
    text_width = self.get_string_width(text, font_name, font_size)
    
    while text_width > max_width and font_size > 1:
        font_size -= 1
        text_width = self.get_string_width(text, font_name, font_size)
    
    canvas.setFont(font_name, font_size)
    canvas.drawString(x, y, text)
  
  def draw_text_break(self, canvas, text, x, y, max_width, font_size, font_name="Helvetica") -> None:
    doc = SimpleDocTemplate("temp.pdf", pagesize=letter)
    
    # Configurar o estilo do parágrafo
    styles = getSampleStyleSheet()
    style = styles["Normal"]
    style.fontName = font_name
    style.fontSize = font_size
    style.leading = font_size * 1

    # Criar o parágrafo com quebra de linha automática

    paragraph = Paragraph(text, style)

    # Construir o parágrafo no documento temporário
    doc.build([paragraph])

    # Desenhar o parágrafo no canvas
    canvas.saveState()
    canvas.translate(x, y)
    paragraph.wrapOn(canvas, max_width, 500)
    paragraph.drawOn(canvas, 0, 0)
    canvas.restoreState()

  def draw_mwm_barcode(self, canvas, barcode_value, x, y):
    barcode = code39.Standard39(barcode_value, barHeight=8*mm, barWidth=0.24*mm, checksum=True)
    barcode.drawOn(canvas, x, y)

  def draw_desc_barcode(self, canvas, barcode_value, x, y, max_width, readable=True) -> None:
    barcode = code128.Code128(barcode_value, barWidth=0.5*mm, barHeight=15*mm, humanReadable=readable, fontSize=10, fontName="Helvetica")
    barcode_width = barcode.width

    scale = max_width / barcode_width
    # Desenha o código de barras com a escala calculada
    canvas.saveState()
    canvas.scale(scale, 1)
    barcode.drawOn(canvas, x / scale, y)  # Ajusta a posição para considerar a escala
    canvas.restoreState()

  def draw_qtd_barcode(self, canvas, barcode_value, x, y, max_width, barHeight) -> None:
    barcode = code128.Code128(barcode_value, barWidth=0.5*mm, barHeight=barHeight, humanReadable=True, fontSize=10, fontName="Helvetica")
    barcode_width = barcode.width

    scale = max_width / barcode_width
    # Desenha o código de barras com a escala calculada
    canvas.saveState()
    canvas.scale(scale, 1)
    barcode.drawOn(canvas, x / scale, y)  # Ajusta a posição para considerar a escala
    canvas.restoreState()

  def create_mwm_label(self) -> None:
    width = 105*mm # Largura da etiqueta
    height = 75*mm # Altura da Etiqueta
    margin_lefts = 7.5*mm # Margens laterais da etiqueta
    margin_top_bottom = 5.5*mm # Margens de cima e de baixo da etiqueta
    internal_margin = margin_lefts+(14*mm) # Margin interna de 14mm após os elementos Part, Qty, Lot, ID
    max_width = 90*mm # Largura maxima do conteudo

    # Campos estáticos da etiqueta
    square_x, square_y = margin_lefts, margin_top_bottom
    client_x, client_y = 8*mm, 70*mm
    version_x, version_y = 86.5*mm, 70*mm
    part_x, part_y = 8*mm, 66*mm
    manufacturer_x, manufacturer_y = internal_margin, 57*mm
    supplier_x, supplier_y = 8*mm, 46*mm
    date_x, date_y = 73*mm, 46*mm
    qty_x, qty_y = 8*mm, 41*mm
    pcs_x, pcs_y = 83*mm, 36*mm
    lot_x, lot_y = 8*mm, 30.5*mm
    id_x, id_y = 8*mm, 17*mm

    # Campos dinâmicos da etiqueta
    partnumber_x, partnumber_y = internal_margin, 60*mm
    partnumber_barcode_x, partnumber_barcode_y = 15*mm, 48.5*mm
    today_date_x, today_date_y = 80*mm, 46*mm
    qty_barcode_x, qty_barcode_y = 15*mm, 36*mm
    qty_number_x, qty_number_y = 73*mm, 36*mm
    op_number_x, op_number_y = internal_margin, 30.5*mm
    op_barcode_x, op_barcode_y = 15*mm, 22*mm
    id_number_x, id_number_y = internal_margin, 17*mm
    id_barcode_x, id_barcode_y = 15*mm, 8*mm

    # pdfmetrics.registerFont(TTFont('SpaceMono', './SpaceMono-Regular.ttf'))
    pdf = canvas.Canvas(f"etq_mwm.pdf", pagesize=landscape((width,height)))
    pdfmetrics.registerFont(TTFont('YuGothicUISemibold', 'C:\Windows\Fonts\YUGOTHB.TTC'))

    # Campos estáticos da Etiqueta
    # Desenhando Quadrado
    pdf.rect(square_x, square_y, 90*mm, 64*mm)

    # Desenhando as linhas 
    pdf.line(margin_lefts, 45*mm, 97.5*mm, 45*mm)
    pdf.line(margin_lefts, 34*mm, 97.5*mm, 34*mm)
    pdf.line(margin_lefts, 21*mm, 97.5*mm, 21*mm)

    # Fornecedor e versão da etiqueta
    pdf.setFont("YuGothicUISemibold", 9)
    self.draw_text(pdf, "MWM MOTORES E GERADORES", client_x, client_y, max_width, 9)
    pdf.setFont("YuGothicUISemibold", 10)
    self.draw_text(pdf, "V.2.3.4", version_x, version_y, max_width, 10)

    # Part:
    pdf.setFont("Helvetica-Bold", 10)
    self.draw_text(pdf, "Part:", part_x, part_y, max_width, 12)
    # Fornecedor/Fabricante
    pdf.setFont("Helvetica-Bold", 6)
    self.draw_text(pdf, "fornecedor/fabricante", manufacturer_x, manufacturer_y, max_width, 6)
    # Supplier:
    pdf.setFont("Helvetica-Bold", 10)
    self.draw_text(pdf, "Supplier: 15175", supplier_x, supplier_y, max_width, 8)
    # Date
    self.draw_text(pdf, "Date:", date_x, date_y, max_width, 8)
    # Qty:
    pdf.setFont("Helvetica-Bold", 10)
    self.draw_text(pdf, "Qty:", qty_x, qty_y, max_width, 12)
    # Lot:
    pdf.setFont("Helvetica-Bold", 10)
    self.draw_text(pdf, "Lot:", lot_x, lot_y, max_width, 12)
    # ID:
    pdf.setFont("Helvetica-Bold", 10)
    self.draw_text(pdf, "ID:", id_x, id_y, max_width, 12)

    #Campos Dinâmicos
    pdf.setFont("Helvetica-Bold", 26)
    # Partnumber
    self.draw_text(pdf, "970000770526", partnumber_x, partnumber_y, max_width, 26)
    # Partnumber barcode
    self.draw_mwm_barcode(pdf, "970000770526", partnumber_barcode_x, partnumber_barcode_y)
    # Today Date
    pdf.setFont("Helvetica-Bold", 9)
    self.draw_text(pdf, f"{self.label_info.date}", today_date_x, today_date_y, max_width, 9)
    # Qty Barcode
    self.draw_mwm_barcode(pdf, "5", qty_barcode_x, qty_barcode_y)
    # Qty Number
    pdf.setFont("Helvetica-Bold", 24)
    self.draw_text(pdf, f"{self.label_info.quantity}", qty_number_x, qty_number_y, max_width, 24)
    # PCs
    pdf.setFont("Helvetica-Oblique", 16)
    self.draw_text(pdf, "PCs", pcs_x, pcs_y, max_width, 16)
    # Lot number
    pdf.setFont("Helvetica-Bold", 9)
    self.draw_text(pdf, "218768", op_number_x, op_number_y, max_width, 9)
    # Lot barcode
    self.draw_mwm_barcode(pdf, "218768", op_barcode_x, op_barcode_y)
    # ID Number
    id = f"1517520240829000001"
    pdf.setFont("Helvetica-Bold", 9)
    self.draw_text(pdf, id, id_number_x, id_number_y, max_width, 9)
    # ID BarCode
    self.draw_mwm_barcode(pdf, id, id_barcode_x, id_barcode_y)
    
    pdf.save()

  def create_label(self) -> None:
    # Dimensões da etiqueta
    width = 150*mm
    height = 100*mm

    margin_left = 5*mm
    max_width = 140*mm

    # Posição dos elementos
    # Data
    date_x, date_y = 10*mm, height-25*mm

    # Lote das caixas
    boxes_x, boxes_y = 50*mm, height-25*mm

    # Linha horizontal
    line_x1, line_y1, line_x2, line_y2 = 0, height-30*mm, width, height-30*mm
    # Cliente
    client_x, client_y = margin_left, height-40*mm
    # Código
    code_x, code_y = margin_left, height-50*mm
    # Descrição
    description_x, description_y = margin_left, height-62*mm
    # Código de barras descrição.
    desc_barcode_x, desc_barcode_y = 0, height-80*mm
    # Quantidade
    qtd_x, qtd_y = margin_left, height-95*mm
    # Código de barras quantidade.
    qtd_barcode_x, qtd_barcode_y = 60*mm, height-95*mm

    weight_x, weight_y = 120*mm, height-95*mm

    pdf = canvas.Canvas(f"etq.pdf", pagesize=landscape((width,height)))
    #Inserindo elementos
    # Data
    self.draw_text(pdf, f"{self.label_info.date}", date_x, date_y, max_width, 12)

    # Lote das caixas
    self.draw_text(pdf, f"{self.label_info.boxes}", boxes_x, boxes_y, max_width, 12)

    # Linha Divisória
    pdf.line(line_x1, line_y1, line_x2, line_y2)

    # Cliente
    self.draw_text(pdf, f"CLIENTE: {self.label_info.client}", client_x, client_y, max_width, 18)

    # Código
    self.draw_text(pdf, f"CÓDIGO: {self.label_info.code}", code_x, code_y, max_width, 20)

    # Descrição
    self.draw_text_break(pdf, f"DESCRIÇÃO: {str(self.label_info.description)}", description_x, description_y, max_width, 12)

    # Código de Barras
    self.draw_desc_barcode(pdf, f"{self.label_info.set_barcode_data()}", desc_barcode_x, desc_barcode_y, max_width)

    # Quantidade
    self.draw_text(pdf, f"QUANTIDADE: {self.label_info.quantity} UND", qtd_x, qtd_y, max_width, 11)

    # Codigo de barras quantidade
    self.draw_qtd_barcode(pdf, f"{self.label_info.quantity}", qtd_barcode_x, qtd_barcode_y, 40*mm, 10*mm)

    # Peso
    self.draw_text(pdf, f"{self.label_info.weight} KG", weight_x, weight_y, 145*mm, 12)

    pdf.save()

  def print_label(self, file_path: str = './etq.pdf') -> None:
    abs_file_path: str = os.path.abspath(file_path)
    
    # Configuração impressora
    default_printer: str = win32print.GetDefaultPrinter()
    hprinter = win32print.OpenPrinter(default_printer)
    try:  
      win32api.ShellExecute(0,"print", abs_file_path, None, ".", 0)
    except Exception as e:
      print("Erro ao enviar para impressão: ", e)
    finally:
      win32print.ClosePrinter(hprinter)

teste = LabelPrint(LabelInfo("MWM", "MWM039","CHICOTAO", 10, 2, 5.00))
teste.create_mwm_label()