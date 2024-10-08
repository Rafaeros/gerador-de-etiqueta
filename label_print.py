# Tamanho da etiqueta: ETIQUETA ADESIVA BOPP PEROLADO 100 x 150mm - COUCHE PERSONALIZADA 250 UNID.
from PIL import Image, ImageDraw, ImageFont
from barcode import Code128
from barcode.writer import ImageWriter
from win32 import win32print, win32api
import os
from io import BytesIO

class LabelPrint():
  def __init__(self, LabelInfo) -> None:
    self.label_info = LabelInfo

  def mm_to_px(self, mm: float, DPI: int =300) -> int:
    return int(mm * (DPI / 25.4))

  def wrap_text(self, text: str, max_chars: int = 35) -> str:
    # Inicializa variáveis
    MAX_CHARS: int = max_chars
    words: list[str] = text.split(' ')
    lines: list[str] = []
    current_line: str = ""

    # Verifica se a quantidade de caracteres excedeu o limite definido e quebra a linha do texto
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

  def draw_label_elements(self, image: Image, draw: ImageDraw, attributes: list[dict], font_path: str) -> None:
    for attribute in attributes:
      if 'width' in attribute:
        barcode_buffer: BytesIO = BytesIO()
        barcode: Code128 = Code128(attribute['text'], writer=ImageWriter())
        options: dict = {
          "module_width": attribute['width'],                
          "module_height": attribute['height'], 
          "quiet_zone": 1, 
          "font_path": font_path, 
          "font_size": 10
        }
        barcode.write(barcode_buffer, options)
        barcode_image: Image = Image.open(barcode_buffer)
        image.paste(barcode_image, attribute['xy'])
      else:
        draw.text(attribute['xy'], attribute['text'], font=attribute['font'], fill='black')

  def create_label(self) -> None:
    WIDTH: int = self.mm_to_px(150)
    HEIGHT: int = self.mm_to_px(100)
    FONT_PATH: str = "./assets/fonts/FiraCode-Regular.ttf"

    # Posições dos elementos da etiqueta (label)
    attr: list[dict] = [
      {'xy': (self.mm_to_px(10), self.mm_to_px(20)), 'text': f'{self.label_info.date}', 'font': ImageFont.truetype(FONT_PATH, size=62)},
      {'xy': (self.mm_to_px(50), self.mm_to_px(20)), 'text': f'{self.label_info.boxes}', 'font': ImageFont.truetype(FONT_PATH, size=62)},
      {'xy': (self.mm_to_px(5), self.mm_to_px(27)), 'text': f'CLIENTE: {self.wrap_text(self.label_info.client)}', 'font': ImageFont.truetype(FONT_PATH, size=55)},
      {'xy': (self.mm_to_px(5), self.mm_to_px(40)), 'text': f'CÓDIGO: {self.label_info.code}', 'font': ImageFont.truetype(FONT_PATH, size=80)},
      {'xy': (self.mm_to_px(5), self.mm_to_px(50)), 'text': f'DESCRICÃO: {self.wrap_text(self.label_info.description)}', 'font': ImageFont.truetype(FONT_PATH, size=62)},
      {'xy': (self.mm_to_px(5), self.mm_to_px(70)), 'text': f'{self.label_info.barcode}', 'font': ImageFont.truetype(FONT_PATH, size=62), 'width': 0.35, 'height': 12},  
      {'xy': (self.mm_to_px(5),self.mm_to_px(90)), 'text': f'QUANTIDADE: {self.label_info.quantity}', 'font': ImageFont.truetype(FONT_PATH, size=62)},
      {'xy': (self.mm_to_px(70), self.mm_to_px(88)), 'text': f'{self.label_info.quantity}', 'font': ImageFont.truetype(FONT_PATH, size=62), 'width': 0.35, 'height': 10},
      {'xy': (self.mm_to_px(120), self.mm_to_px(90)), 'text': f'{self.label_info.weight} Kg', 'font': ImageFont.truetype(FONT_PATH, size=62)},
    ]

    # Criar imagem em branco (label)
    label: Image = Image.new('RGB', (WIDTH, HEIGHT), color='white')

    # Criar objeto para desenhar na imagem
    draw: ImageDraw.Draw = ImageDraw.Draw(label)

    # Desenhando elementos do label
    draw.line((self.mm_to_px(0), self.mm_to_px(26), self.mm_to_px(160), self.mm_to_px(26)), fill='black', width=2)
    self.draw_label_elements(label, draw, attr, FONT_PATH)

    label.thumbnail((WIDTH, HEIGHT))
    label.save(f'./tmp/etq.png')

  def print_label(self, file_path: str = './tmp/etq.png') -> None:
    abs_file_path: str = os.path.abspath(file_path)
    
    # Configuração impressora
    default_printer: str = win32print.GetDefaultPrinter()
    hprinter = win32print.OpenPrinter(default_printer)
      
    try:
      win32print.StartDocPrinter(hprinter, 1, ("Print PNG", None, "RAW"))
      win32print.StartPagePrinter(hprinter)
      with open(abs_file_path, 'rb') as img:
        win32print.WritePrinter(hprinter, img.read())
      win32print.EndPagePrinter(hprinter)
      win32print.EndDocPrinter(hprinter)
      #win32api.ShellExecute(0, "print", abs_file_path, None, ".", 0)
    except Exception as e:
      print("Erro ao enviar para impressão: ", e)
    finally:
      win32print.ClosePrinter(hprinter)