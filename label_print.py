# Tamanho da etiqueta: ETIQUETA ADESIVA BOPP PEROLADO 100 x 150mm - COUCHE PERSONALIZADA 250 UNID.
from PIL import Image, ImageDraw, ImageFont
from barcode import Code128
from barcode.writer import ImageWriter
from win32 import win32print, win32api
import os

class LabelPrint():
  def __init__(self, LabelInfo) -> None:
    self.label_info = LabelInfo

  def mm_to_px(self, mm, DPI=300) -> int:
    return int(mm * (DPI / 25.4))

  def wrap_text(self, text: str) -> str:
    # Inicializa variáveis
    MAX_CHARS: int = 50
    words: list[str] = text.split(' ')
    lines: list[str] = []
    current_line: str = ""

    for word in words:
      if len(current_line) + len(word) + 1 > MAX_CHARS:
          lines.append(current_line) 
          current_line = word
      else:
        if current_line:
            current_line += ' ' + word
        else:
            current_line = word  

    if current_line:
      lines.append(current_line)
    
    wrapped_text: str = '\n'.join(lines)

    return wrapped_text

  def description_barcode(self, text) -> None:
    barcode: Code128 = Code128(text, writer=ImageWriter())
    barcode.save("description_barcode", options={"module_width": 0.2, "module_height": 10, "quiet_zone": 1, "font_path": "FiraCode-Regular.ttf", "font_size": 10})

  def quantity_barcode(self, text) -> None:
    barcode: Code128 = Code128(text, writer=ImageWriter())
    barcode.save("quantity_barcode", options={"module_width": 0.3, "module_height": 5, "quiet_zone": 1, "font_path": "FiraCode-Regular.ttf", "font_size": 10})

  def create_label(self) -> None:
    WIDTH: int = self.mm_to_px(150)
    HEIGHT: int = self.mm_to_px(100)
    FONT_NAME: str = "FiraCode-Regular.ttf"

    # Posições dos elementos da etiqueta (label)
    positions: dict[str, int] = {
      'date': (self.mm_to_px(10), self.mm_to_px(15)),
      'boxes': (self.mm_to_px(80), self.mm_to_px(15)),
      'line': (self.mm_to_px(0), self.mm_to_px(20), self.mm_to_px(150), self.mm_to_px(20)),
      'client': (self.mm_to_px(5), self.mm_to_px(25)),
      'code': (self.mm_to_px(5), self.mm_to_px(35)),
      'description': (self.mm_to_px(5), self.mm_to_px(45)),
      'barcode': (self.mm_to_px(5), self.mm_to_px(65)),
      'quantity': (self.mm_to_px(5), self.mm_to_px(90)),
      'quantity_barcode': (self.mm_to_px(70), self.mm_to_px(88)),
      'weight': (self.mm_to_px(120), self.mm_to_px(90))
      }

    # Criar imagem em branco (label)
    label: Image = Image.new('RGB', (WIDTH, HEIGHT), color='white')
    font: ImageFont = ImageFont.truetype(FONT_NAME, size=50)
    # Criar objeto para desenhar na imagem
    draw: ImageDraw.Draw = ImageDraw.Draw(label)

    client: str = self.wrap_text(self.label_info.client)
    description: str = self.wrap_text(self.label_info.description)
    self.description_barcode(self.label_info.barcode)
    self.quantity_barcode(str(self.label_info.quantity))

    draw.text(positions.get('date'), f'Data: {self.label_info.date}', fill='black', font=font)
    draw.text(positions.get('boxes'), str(self.label_info.boxes), fill='black', font=font)
    draw.line(positions.get('line'), fill='black', width=3)
    draw.text(positions.get('client'), f"CLIENTE: {client}",fill='black', font=font)
    draw.text(positions.get('code'), f"CÓDIGO: {self.label_info.code}", fill='black', font=font)
    draw.text(positions.get('description'), f"DESCRIÇÃO: {description}", fill='black', font=font)

    # Codigo de barras da descrição
    description_barcode = Image.open('description_barcode.png')
    label.paste(description_barcode, positions.get('barcode'))

    draw.text(positions.get('quantity'), f"QUANTIDADE: {str(self.label_info.quantity)} UND", fill='black', font=font)

    # Codig de barras da quantidade
    quantity_barcode = Image.open('quantity_barcode.png')
    label.paste(quantity_barcode, positions.get('quantity_barcode'))

    draw.text(positions.get('weight'), f"{self.label_info.weight} KG", fill='black', font=font)

    label.thumbnail((WIDTH, HEIGHT))
    label.save('etq.png')

  def print_label(self, file_path: str = './etq.png') -> None:
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