import tkinter as tk
from tkinter import filedialog, messagebox
import requests
import pandas as pd
from datetime import datetime
import pytz
import time

CLIENT_ID = ""
CLIENT_SECRET = ""
CODE = ""
URI = ""

class App:
    def __init__(self, master):
        self.master = master
        master.title("Publicación masiva en Mercado Libre")
        
        self.label = tk.Label(master, text="Seleccione un archivo de Excel:")
        self.label.pack()
        self.select_button = tk.Button(master, text="Seleccionar archivo", command=self.load_file)
        self.select_button.pack()

        self.label2 = tk.Label(master, text="¿En qué categoría quieres hacer las publicaciones?")
        self.label3 = tk.Label(master, text="1 = Cascos, 2 = , 3 = ")
        self.label2.pack()
        self.label3.pack()
        self.category_entry = tk.Entry(master)
        self.category_entry.pack()

        self.refresh_token_label = tk.Label(master, text="Ingrese el token de refresco, (solo después de la segunda vez):")
        self.refresh_token_label.pack()
        self.refresh_token_entry = tk.Entry(master)
        self.refresh_token_entry.pack()
        self.run_button = tk.Button(master, text="Ejecutar", command=self.run)
        self.run_button.pack()
        
        self.output_frame = tk.Frame(master)
        self.output_frame.pack()
        self.process_output = tk.Text(self.output_frame, wrap=tk.WORD, height=10)
        self.process_output.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar = tk.Scrollbar(self.output_frame, command=self.process_output.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.process_output.config(yscrollcommand=self.scrollbar.set)

        self.filepath = None 
        self.token_info = None  

        self.id = CLIENT_ID
        self.secret = CLIENT_SECRET
        self.code = CODE
        self.uri = URI

    def load_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.process_output.insert(tk.END, f"Archivo seleccionado: {self.filepath}\n")

    def Get_Token(self):
        url = "https://api.mercadolibre.com/oauth/token"
        headers = {
            "accept": "application/json",
            "content-type": "application/x-www-form-urlencoded"
        }

        data = {
            "grant_type": "authorization_code",
            "client_id": self.id,
            "client_secret": self.secret,
            "code": self.code,
            "redirect_uri": self.uri
        }

        response_data = requests.post(url, headers=headers, data=data)
        response = response_data.json()

        if "access_token" in response:
            return {
                "success": True,
                "token": response["access_token"],
                "refresh_token": response.get("refresh_token")
            }
        elif "error" in response and response.get("status") == 400:
            if (response.get("message") == "Error validating grant. Your authorization code or refresh token may be expired or it was already used"):

                new_data = {
                "grant_type": "refresh_token",
                "client_id": self.id,
                "client_secret": self.secret,
                "refresh_token": self.refresh_token_entry.get()
                }

                new_response = requests.post(url, headers=headers, data=new_data)
                new_response_data = new_response.json()

            if "access_token" in new_response_data:
                return {
                    "success": True,
                    "token": new_response_data["access_token"],
                    "refresh_token": new_response_data.get("refresh_token"),
                }
            else:
                return {
                    "success": False,
                    "message": new_response_data.get(
                        "message", "Error al obtener un nuevo token."
                    ),
                }
        elif response.get("message") == "invalid client_id or client_secret":
            return {
                "success": False,
                "message": "El token o el usuario son incorrectos.",
            }
        else:
            return {"success": False, "message": "La solicitud ha salido mal."}

    
    def log_to_file(self, info):

        colombia_tz = pytz.timezone("America/Bogota")
        now = datetime.now(colombia_tz)
        timestamp = now.strftime("%d/%m/%Y %H:%M:%S")

        with open("../../materials/registros.txt", "a") as file:
            file.write("-" * 100 + "\n")
            file.write(f"Fecha y hora: {timestamp}\n")
            file.write(f"El token es: {info['token']}\n")
            file.write(f"El token de refresco es: {info['refresh_token']}\n")
            file.write("-" * 100 + "\n")

    def create_product(self, data, category):

        if category == 1:
            producto = {
                "title": data["Nombre"],  # nombre
                "category_id": "MCO21947",  # categoria
                "price": data["Precio"],  # precio
                "currency_id": "COP",  # moneda
                "available_quantity": data["Cantidad"],  # cantidad
                "buying_mode": "buy_it_now",  # precio fijo
                "condition": "new",  # condiccion
                "description": data["Descripcion"],
                "listing_type_id": "free",  # "free", "bronze", "silver", "gold_special", "gold_pro", "gold_premium"
                "sale_terms": [
                    {
                        "id": "WARRANTY_TYPE",
                        "value_name": data["Garantia"],
                    },  # garantia
                    {"id": "WARRANTY_TIME", "value_name": data["Garantia_Tiempo"]},  # tiempo de garantia
                ],
                "pictures": [  # imagenes
                    {
                        "source": data["Url_Imagen_1"]
                    },
                    {
                        "source": data["Url_Imagen_2"]
                    },
                    {
                        "source": data["Url_Imagen_3"]
                    },
                ],
                "attributes": [
                    {
                        "id": "BRAND", 
                        "value_name": data["Marca"]},  # Marca del casco
                    {
                        "id": "LINE", 
                        "value_name": data["Linea"]},  # Línea o serie del casco
                    {
                        "id": "MODEL",
                        "value_name": data["Modelo"],
                    },  # Nombre del modelo específico
                    {
                        "id": "ALPHANUMERIC_MODEL",
                        "value_name": data["Codigo_Alfanumerico"],
                    },  # Código alfanumérico que representa el modelo
                    {
                        "id": "COLOR", 
                        "value_name": data["Color"]},  # Color general del casco
                    {
                        "id": "DESIGN", 
                        "value_name": data["Diseño"] },  # Estilo del diseño exterior
                    {
                        "id": "FINISH",
                        "value_name": data["Acabado"],
                    },  # Tipo de acabado (mate, brillante, etc.)
                    {
                        "id": "HELMET_SIZE",
                        "value_name": data["Talla"],
                    },  # Talla del casco (por ejemplo: S, M, L)
                    {
                        "id": "HEAD_CIRCUMFERENCE_SIZE",
                        "value_name": data["Circunferencia_Cabeza"],
                    },  # Circunferencia de la cabeza para esta talla
                    {
                        "id": "MAIN_COLOR",
                        "value_name": data["Color_Principal"],
                    },  # Color principal del casco
                    {
                        "id": "DETAILED_MODEL",
                        "value_name": data["Modelo_especifico"],
                    },  # Modelo con más detalles específicos
                    {
                        "id": "HELMET_TYPE",
                        "value_name": data["Tipo_Casco"],
                    },  # Tipo de casco (ej. integral, jet, etc.)
                    {
                        "id": "INNER_MATERIALS",
                        "value_name": data["Material_Interno"],
                    },  # Material interno del casco
                    {
                        "id": "OUTER_MATERIALS",
                        "value_name": data["Material_Externo"],
                    },  # Material externo del casco
                    {
                        "id": "AGE_GROUP",
                        "value_name": data["Grupo_para_el_que_esta_diseñado"],
                    },  # Grupo de edad para el que está diseñado
                    {
                        "id": "MOTORCYCLE_RIDING_STYLE",
                        "value_name": data["Estilo_Conduccion"],
                    },  # Estilo de conducción recomendado (urbano, deportivo, etc.)
                    {
                        "id": "SAFETY_REGULATIONS",
                        "value_name": data["Certificado_de_Segurida"],
                    },  # Certificación de seguridad que cumple el casco
                    {
                        "id": "GTIN",
                        "value_name": data["Codigo_de_Barras"],
                    },  # Número Global de Artículo Comercial (código de barras)
                    {
                        "id": "VISOR_TYPE",
                        "value_name": data["Tipo_Visor"],
                    },  # Tipo de visor (transparente, polarizado, etc.)
                    {
                        "id": "VISOR_MATERIALS",
                        "value_name": data["Meterial_Visor"],
                    },  # Material del visor
                    {
                        "id": "VISOR_THICKNESS", 
                        "value_name": data["Espesor_Visor"]},  # Espesor del visor
                    {
                        "id": "VENTILATION_POSITIONS",
                        "value_name": data["Numero_Posiciones_Ventilacion"]
                    },  # Número de posiciones de ventilación disponibles
                    {
                        "id": "CLOSURE_TYPE",
                        "value_name": data["Tipo_Cierre"]
                    },  # Tipo de cierre del casco
                    {
                        "id": "WITH_UV_PROTECTION_VISOR",
                        "value_name": data["Proteccion_Uv"],
                    },  # Indica si el visor tiene protección UV
                    {
                        "id": "WITH_SCRATCH_RESISTANT_VISOR",
                        "value_name": data["Resistencia_del_Visor_a_Arañasos"],
                    },  # Indica si el visor es resistente a los arañazos
                    {
                        "id": "WITH_PINLOCK_READY_VISOR",
                        "value_name": data["Compatibilidad_pinLock"],
                    },  # Indica si el casco es compatible con pinlock (sistema antivaho)
                    {
                        "id": "WITH_BREATHABLE_LINING",
                        "value_name": data['Posee_Forro'],
                    },  # Indica si el casco tiene forro transpirable
                    {
                        "id": "IS_DETACHABLE",
                        "value_name": data["Interior_Desmontable"],
                    },  # Indica si el interior del casco es desmontable
                    {
                        "id": "IS_HYPOALLERGENIC_PRODUCT",
                        "value_name": data["Es_Hipoalergenico"],
                    },  # Indica si el producto es hipoalergénico
                    {
                        "id": "MPN",
                        "value_name": data["Numero_Casco_por_parte_Fabricante"],
                    },  # Número de parte del fabricante
                    {
                        "id": "ITEM_CONDITION",
                        "value_name": data["Condicion"],
                    },  # Condición del artículo (nuevo, usado, etc.)
                ],
                "shipping": {
                    "mode": "me2",  # indica que se usa la logistica de mercado libre
                    "local_pick_up": True,  # recogida donde el vendedor
                    "free_shipping": data["Envio"],  # envio gratis o no
                },
            }

        elif category == 2:
            print("Opcion 2")
        elif category == 3:
            print("Opcion 3")
        url_item = "https://api.mercadolibre.com/items"
        headers_item = {
            "Authorization": f"Bearer {self.token_info['token']}",
            "Content-Type": "application/json",
        }
        response_item = requests.post(url_item, json=producto, headers=headers_item)
        return response_item.status_code
    
    def run(self):
        if not self.filepath:
            messagebox.showerror("Error", "Por favor, seleccione un archivo primero.")
            return

        self.process_output.insert(tk.END, "Ejecutando procesos...\n")
        self.process_output.see(tk.END)
        self.master.update()

        try:
            category_value = int(self.category_entry.get()) 
        except ValueError:
            messagebox.showerror("Error", "¡La categoría debe ser un número!")
            return

        if category_value in [1, 2, 3, 4, 5, 6]:
            try:
                self.token_info = self.Get_Token()

                if self.token_info["success"]:
                    self.log_to_file(self.token_info)

                    self.process_output.insert(tk.END, "_" * 93)
                    self.process_output.see(tk.END)
                    self.master.update()

                    self.process_output.insert(tk.END, f"el token es {self.token_info['token']}")
                    self.process_output.see(tk.END)
                    self.master.update()

                    self.process_output.insert(tk.END, "_" * 93)
                    self.process_output.see(tk.END)
                    self.master.update()

                    self.process_output.insert(tk.END, f"El token de refresco es: {self.token_info['refresh_token']}")
                    self.process_output.see(tk.END)
                    self.master.update()


                    df = pd.read_excel(self.filepath)
                    counter = 0
                    counter_success = 0
                    counter_error = 0
                    products_error = []
                    for index, row in df.iterrows():
                        result = self.create_product(row.to_dict(), category_value)
                        self.process_output.insert(tk.END, "_" * 93)
                        self.process_output.see(tk.END)
                        self.master.update()
                        if result == 201:
                            counter_success += 1
                            self.process_output.insert(tk.END, f'El producto : {index + 1} se subio corretamente.')
                            self.process_output.see(tk.END)
                            self.master.update()
                        else:
                            data_row = row.to_dict()
                            products_error.append(data_row['Nombre'])
                            counter_error += 1
                            self.process_output.insert(tk.END, f'El producto : {index + 1} no se subio.')
                            self.process_output.see(tk.END)
                            self.master.update()
                        self.process_output.insert(tk.END, "_" * 93)
                        self.process_output.see(tk.END)
                        self.master.update()
                        counter += 1
                        time.sleep(120)
                    self.process_output.insert(tk.END, "--proceso finalizado--")
                    self.process_output.see(tk.END)
                    self.master.update()
                    self.process_output.insert(tk.END, "_" * 93)
                    self.process_output.see(tk.END)
                    self.master.update()
                    list_products_error = str(products_error)
                    self.process_output.insert(tk.END, f"se subieron un total de {counter} productos, {counter_success} se subieron correctamente y {counter_error} fallaron, los productos que fallaron son {list_products_error}")
                    self.process_output.see(tk.END)
                    self.master.update()
                else:
                    messagebox.showerror("Error", f"Algo salió mal al obtener el token, error: '{self.token_info['message']}'")
                    return
            except Exception as e:
                messagebox.showerror("Error", e)
                return
        else: 
            messagebox.showerror("Error", "¡La categoria no es valida!.")
            return

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root) 
    root.mainloop()
