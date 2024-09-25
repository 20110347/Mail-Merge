import json

#################################### FUNCIONES ##############################

# Function para Lectura de los Destinatarios
def read_data(dir):
    print("Ubicacion del la info: " + dir)
    with open(dir, 'r', encoding="utf-8") as file:
        data = json.load(file)
    return data

# Function para Template de la carta
def read_template(dir):
    print("Ubicacion del template: " + dir)
    with open(dir, "r", encoding="utf-8") as template_file:
        # Se usa .join() para convertir la lista en una String
        fullTemplate = "".join(template_file.readlines())
    return fullTemplate

# Function para añadir un nuevo elemento al json
def add_recipient(data, info):
    data["recipient"].append(info)
    write_document(data)

# Function para modificar algun elemento del json
def mod_recipient(data, info, search_email):
    # Busco entre los destinatarios hasta encontrar el que coincida con el correo
    for recipient, obj in enumerate(data["recipient"]):
        if obj['email'] == search_email:
            # Remuevo el elemento en el que coincida el email por medio
            data["recipient"].pop(recipient)
            data["recipient"].append(info)
            break
    write_document(data)

# Function para eliminar algun elemento del json
def del_recipient(data, search_email):
    # Busco entre los destinatarios hasta encontrar el que coincida con el correo
    for recipient, obj in enumerate(data["recipient"]):
        if obj['email'] == search_email:
            # Remuevo el elemento en el que coincida el email por medio
            data["recipient"].pop(recipient)
            break
    write_document(data)

# Function para escribir en el documento
def write_document(file_data):
    # Vuelvo a abrir el archivo para que borre el texto anterior
    # y Sobrescriba con el nuevo json con el elemento elimanado
    with open("data.json", 'w', encoding="utf-8") as file:
        file.write(json.dumps(file_data, indent=4))