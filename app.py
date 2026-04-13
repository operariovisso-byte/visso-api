from flask_cors import CORS
from flask import Flask, render_template, request, redirect, jsonify
import base64
import os
import tempfile
import json
#import pythoncom
#import win32com.client as win32

app = Flask(__name__)
CORS(app)

# -------------------------
# ARCHIVOS
# -------------------------
RUTA_PRODUCTOS = "productos.json"
RUTA_COLORES = "colores.json"
RUTA_PEDIDOS = "pedidos.json"

# -------------------------
# DATA EN MEMORIA
# -------------------------
productos = []
colores = []
pedidos = []

# -------------------------
# FUNCIONES JSON
# -------------------------
def cargar_datos():
    global productos, colores, pedidos

    try:
        with open(RUTA_PRODUCTOS) as f:
            productos = json.load(f)
    except:
        productos = []

    try:
        with open(RUTA_COLORES) as f:
            colores = json.load(f)
    except:
        colores = []

    try:
        with open(RUTA_PEDIDOS) as f:
            pedidos = json.load(f)
    except:
        pedidos = []

def guardar_productos():
    with open(RUTA_PRODUCTOS, "w") as f:
        json.dump(productos, f, indent=4)

def guardar_colores():
    with open(RUTA_COLORES, "w") as f:
        json.dump(colores, f, indent=4)

def guardar_pedidos():
    with open(RUTA_PEDIDOS, "w") as f:
        json.dump(pedidos, f, indent=4)

# Cargar al iniciar
cargar_datos()

# -------------------------
# HOME
# -------------------------
@app.route("/")
def index():
    return jsonify({
        "status": "API funcionando",
        "pedidos": len(pedidos)
    })

# -------------------------
# PRODUCTOS (🔥 CON RANGOS)
# -------------------------
@app.route("/agregar_producto", methods=["POST"])
def agregar_producto():
    nuevo = {
        "nombre": request.form["nombre"],

        # 🔥 NUEVO FORMATO
        "largo_min": request.form.get("largo_min", ""),
        "largo_max": request.form.get("largo_max", ""),
        "ancho_min": request.form.get("ancho_min", ""),
        "ancho_max": request.form.get("ancho_max", ""),
        "alto_min": request.form.get("alto_min", ""),
        "alto_max": request.form.get("alto_max", ""),

        "accesorios": request.form.get("accesorios", ""),
        "activo": True
    }

    productos.append(nuevo)
    guardar_productos()  # 🔥 CLAVE

    return redirect("/")

@app.route("/toggle_producto/<int:i>")
def toggle_producto(i):
    productos[i]["activo"] = not productos[i]["activo"]
    guardar_productos()
    return redirect("/")

@app.route("/eliminar_producto/<int:i>")
def eliminar_producto(i):
    productos.pop(i)
    guardar_productos()
    return redirect("/")

@app.route("/obtener_colores")
def obtener_colores():
    return jsonify(colores)

# -------------------------
# COLORES
# -------------------------
@app.route("/agregar_color", methods=["POST"])
def agregar_color():
    nuevo = {
        "nombre": request.form["nombre"],
        "clasificacion": request.form["clasificacion"],
        "activo": True
    }

    colores.append(nuevo)
    guardar_colores()

    return redirect("/")

@app.route("/toggle_color/<int:i>")
def toggle_color(i):
    colores[i]["activo"] = not colores[i]["activo"]
    guardar_colores()
    return redirect("/")

@app.route("/set_fecha_entrega", methods=["POST"])
def set_fecha_entrega():
    data = request.json
    index = data["index"]

    pedidos[index]["fecha_entrega"] = data["fecha"]
    pedidos[index]["tipo_entrega"] = data["tipo"]

    guardar_pedidos()

    return jsonify({"ok": True})

@app.route("/eliminar_color/<int:i>")
def eliminar_color(i):
    colores.pop(i)
    guardar_colores()
    return redirect("/")

# -------------------------
# PEDIDOS
# -------------------------
@app.route("/guardar_pedido", methods=["POST"])
def guardar_pedido():
    data = request.json

    # 🚫 VALIDAR DUPLICADO
    for p in pedidos:
        if p["numero"] == data["numero"]:
            return jsonify({"error": "El número de pedido ya existe"}), 400

    nuevo = {
        "cliente": data["cliente"],
        "numero": data["numero"],
        "consultora": data["consultora"],
        "productos": data["productos"],
        "pagado": False,
        "enviado": False,
        "metodo_pago": "",
        "imagen_pago": ""
    }

    pedidos.append(nuevo)
    guardar_pedidos()

    return jsonify({"ok": True})

@app.route("/editar_color/<int:i>", methods=["POST"])
def editar_color(i):
    colores[i]["nombre"] = request.form["nombre"]
    colores[i]["clasificacion"] = request.form["clasificacion"]

    guardar_colores()  # 🔥 FALTABA ESTO

    return redirect("/")

@app.route("/editar_producto/<int:i>", methods=["POST"])
def editar_producto(i):
    productos[i].update({
        "nombre": request.form.get("nombre"),
        "largo_min": request.form.get("largo_min"),
        "largo_max": request.form.get("largo_max"),
        "ancho_min": request.form.get("ancho_min"),
        "ancho_max": request.form.get("ancho_max"),
        "alto_min": request.form.get("alto_min"),
        "alto_max": request.form.get("alto_max"),
        "accesorios": request.form.get("accesorios")
    })

    guardar_productos()
    return redirect("/")

@app.route("/obtener_pedidos")
def obtener_pedidos():
    return jsonify(pedidos)

# -------------------------
# REGISTRAR PAGO
# -------------------------
@app.route("/registrar_pago", methods=["POST"])
def registrar_pago():
    data = request.json
    index = data["index"]

    pedidos[index]["pagado"] = True
    pedidos[index]["metodo_pago"] = data["metodo"]
    pedidos[index]["imagen_pago"] = data["imagen"]

    guardar_pedidos()

    return jsonify({"ok": True})

@app.route("/eliminar_pedido", methods=["POST"])
def eliminar_pedido():
    data = request.json
    index = data["index"]

    pedidos.pop(index)
    guardar_pedidos()

    return jsonify({"ok": True})

# -------------------------
# ENVIAR CORREO (ESTABLE)
# -------------------------
#@app.route("/enviar_correo", methods=["POST"])
#def enviar_correo():
    try:
        pythoncom.CoInitialize()

        data = request.get_json()
        index = data.get("index")
        pedido = pedidos[index]

        if not pedido.get("pagado"):
            return jsonify({"error": "Debe registrar pago primero"}), 400

        estandar = [p for p in pedido["productos"] if p["tipo"] == "Estandar"]
        especial = [p for p in pedido["productos"] if p["tipo"] == "Especial"]

        outlook = win32.Dispatch("Outlook.Application")
        outlook.GetNamespace("MAPI")

        # IMAGEN
        ruta_img = None
        if pedido.get("imagen_pago") and pedido.get("metodo_pago") != "PAGO EN LINEA":
            try:
                img_data = pedido["imagen_pago"].split(",")[1]
                ruta_img = os.path.join(tempfile.gettempdir(), f"pago_{index}.jpg")

                with open(ruta_img, "wb") as f:
                    f.write(base64.b64decode(img_data))
            except Exception as e:
                print("Error imagen:", e)

        # HTML
        def crear_html(productos, tipo):
            filas = "".join([
                f"""
                <tr>
                    <td style='padding:8px'>{p['nombre']}</td>
                    <td style='padding:8px; text-align:center'>{p['cantidad']}</td>
                </tr>
                """
                for p in productos
            ])

            return f"""
            <div style="font-family:Arial">
                <h2 style="color:#1f3a5f;">Pedido {pedido['numero']}</h2>
                <p><b>Cliente:</b> {pedido['cliente']}</p>
                <p><b>Consultora:</b> {pedido['consultora']}</p>

                <h3>{tipo}</h3>

                <table border="1" cellspacing="0" style="border-collapse:collapse; width:100%;">
                    <tr style="background:#1f3a5f; color:white;">
                        <th>Producto</th>
                        <th>Cantidad</th>
                    </tr>
                    {filas}
                </table>

                <p style="margin-top:20px;">
                    <b>Método de pago:</b> {pedido['metodo_pago']}
                </p>
            </div>
            """

        # CORREO ESTANDAR
        if estandar:
            mail = outlook.CreateItem(0)
            mail.To = "pcp2@visso.com.pe; pcp1@visso.com.pe; gvalencia@visso.com.pe"
            mail.Subject = f"Pedido {pedido['numero']} - ESTANDAR"
            mail.HTMLBody = crear_html(estandar, "Productos Estandar")

            if ruta_img:
                mail.Attachments.Add(ruta_img)

            mail.Display()

        # CORREO ESPECIAL
        if especial:
            mail = outlook.CreateItem(0)
            mail.To = "achunga@visso.com.pe; evaldez@visso.com.pe; gvalencia@visso.com.pe"
            mail.Subject = f"Pedido {pedido['numero']} - ESPECIAL"
            mail.HTMLBody = crear_html(especial, "Productos Especiales")

            if ruta_img:
                mail.Attachments.Add(ruta_img)

            mail.Display()

        pedidos[index]["enviado"] = True
        guardar_pedidos()

        return jsonify({"ok": True})

    except Exception as e:
        print("ERROR:", e)
        return jsonify({"error": str(e)}), 500

    finally:
        pythoncom.CoUninitialize()

# -------------------------
# RUN
# -------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)