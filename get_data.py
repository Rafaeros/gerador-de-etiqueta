from datetime import datetime as dt
import pandas as pd
import re

file_date = dt.now().strftime("%d_%m_%Y")

class LabelInfo:
  def __init__(self, op: str, client: str, code: str, description: str, quantity: int, boxes: int, weight: int) -> None:
    today: str = dt.now()
    self.date: str = today.strftime("%d/%m/%Y")
    self.mwm_date: str = today.strftime("%Y%m%d")
    self.op: str = op
    self.client: str = client
    self.code: str = code
    self.barcode: str = ""
    self.client_code: str = ""
    self.description: str = description
    self.quantity: int = quantity
    self.weight: int = weight
    self.boxes: int = boxes
    self.set_barcode_data()

  def get_client_code(self, text: str) -> str:
    search = re.search(r'\((.*?)\)', text)
    if search:
      return search.group(1)
    return ''

  def set_barcode_data(self) -> str:
    self.client_code = self.get_client_code(self.description)
    self.barcode = self.code + " " + f"({self.client_code})"
    return self.barcode

class LabelData:
  def __init__(self, file_path: str = f"./ordens_{file_date}.xlsx") -> None:
    self.file_path: str = file_path
    self.label_data: pd.DataFrame = self.load_data(self.file_path)
    self.format_data()

  def load_data(self, file_path: str) -> pd.DataFrame | None:
    try:
      return pd.read_excel(file_path)
    except Exception as e:
      print(f"Error loading data: {e}")
      return None
    
  def format_data(self) -> None:
    self.label_data = self.label_data[['Código', 'Cliente', 'Cód. Material', 'Material', 'Quantidade']]

  def print_data(self) -> None:
    print(self.label_data)

  def get_data(self, op: str, boxes: int, kg: str) -> LabelInfo | None:
    try:
      current_data = self.label_data.loc[self.label_data['Código']==op, ['Cliente', 'Cód. Material', 'Material', 'Quantidade']]
      current_label = LabelInfo(op, current_data.loc[:,['Cliente']].to_string(index=False, header=False), current_data.loc[:,['Cód. Material']].to_string(index=False, header=False), current_data.loc[:,['Material']].to_string(index=False, header=False), current_data.loc[:,['Quantidade']].to_string(index=False, header=False), boxes, kg)
      return current_label
    except Exception as e:
      print(f"not found {e}")
      return None