import sys
from PyQt6 import QtWidgets
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.uic import load_ui
from crud_json import *
from generate_documents import *

class PrincipalView(QMainWindow):
    def __init__(self):
        super(PrincipalView, self).__init__()
        load_ui.loadUi('gui_correos.ui', self)
        self.setFixedSize(818, 590)

        # Btns para botones de interacción
        self.btn_close.clicked.connect(lambda: self.close())
        self.btn_min.clicked.connect(self.min)

        # Btns para Seleccionar Archivos
        global fname
        fname = ""
        global fnameT
        fnameT = ""
        self.btn_select_data.clicked.connect(self.select_data_file)
        self.btn_select_template.clicked.connect(self.select_template_file)
        self.btn_add.clicked.connect(
            lambda: self.stackedWidget.setCurrentWidget(self.pg_add))
        self.btn_data.clicked.connect(
            lambda: self.stackedWidget.setCurrentWidget(self.pg_data))
        self.btn_data.clicked.connect(
            self.load_table)
        self.btn_modify.clicked.connect(
            lambda: self.stackedWidget.setCurrentWidget(self.pg_modify))
        self.btn_delete.clicked.connect(
            lambda: self.stackedWidget.setCurrentWidget(self.pg_delete))
        self.btn_delete.clicked.connect(
            self.load_table_delete)
        self.btn_gen_docs.clicked.connect(
            lambda: self.stackedWidget.setCurrentWidget(self.pg_gen))
        self.btn_gen_docs.clicked.connect(
            self.load_table_doc)
        self.btn_gen_docs.clicked.connect(
            self.advice)
        
        # CRUD
        # Create
        self.btn_add_recipient.clicked.connect(self.add_rec)
        # Read
         # Se ejecuta siempre que se cambia al menu de Consultar
        # Update
        self.btn_modify_search.clicked.connect(self.load_info_modify)
        self.btn_modify_recipient.clicked.connect(self.mod_rec)
        # Delete
        self.btn_delete_recipient.clicked.connect(self.del_rec)

        # Metodos para los botones generadores de docs}
        self.btn_txt.clicked.connect(self.gen_txt)
        self.btn_pdf.clicked.connect(self.gen_pdf)
        self.btn_docx.clicked.connect(self.gen_docx)
        self.btn_email.clicked.connect(self.gen_email)

        

######################################## FUNCTIONS ##################################

    # Metodo para minimizar ventana
    def min(self):
        self.showMinimized()

    # Metodo para seleccionar archivo json
    def select_data_file(self):
        global fname
        fname = QFileDialog.getOpenFileName(
            self, "Open File", "", "All Files (*);;JSON Files (*.json)")
        self.load_table()

    # Metodo para seleccionar archivo template
    def select_template_file(self):
        global fnameT
        fnameT = QFileDialog.getOpenFileName(
            self, "Open File", "", "All Files (*);;TXT Files (*.txt)")


######################################## CRUD ##################################

######################################## ADD ##################################
    def add_rec(self):
        if fname:
            data = read_data(fname[0])
            dic = {
                "name": self.edt_name.text(),
                "apellido1":  self.edt_apellido1.text(),
                "apellido2": self.edt_apellido2.text(),
                "cargo": self.edt_cargo.text(),
                "empresa": self.edt_empresa.text(),
                "calle": self.edt_calle.text(),
                "noExt": int(self.edt_ext.text()),
                "noInt": int(self.edt_int.text()),
                "colonia": self.edt_col.text(),
                "municipio": self.edt_mun.text(),
                "estado":  self.edt_estado.text(),
                "cp": int(self.edt_cp.text()),
                "tel": int(self.edt_tel.text()),
                "email": self.edt_email.text(),
                "fnac": self.edt_fnac.text()
            }
            add_recipient(data, dic)
            self.lbl_add_msg.setText("Añadido con Exito")
            self.clean_add()
        else:
            self.error404()
######################################## UPDATE ##################################
    
    #Buscar el Email y cambiar los campos
    def load_info_modify(self):
            search_email = self.edt_search_modify.text()
            if fname:
                data = read_data(fname[0])
                for recipient, obj in enumerate(data["recipient"]):
                    if obj['email'] == search_email:
                        self.edt_name_modify.setText(obj['name'])
                        self.edt_apellido1_modify.setText(obj['apellido1'])
                        self.edt_apellido2_modify.setText(obj['apellido2'])
                        self.edt_cargo_modify.setText(obj['cargo'])
                        self.edt_empresa_modify.setText(obj['empresa'])
                        self.edt_calle_modify.setText(obj['calle'])
                        self.edt_ext_modify.setText(str(obj['noExt']))
                        self.edt_int_modify.setText(str(obj['noInt']))
                        self.edt_col_modify.setText(obj['colonia'])
                        self.edt_mun_modify.setText(obj['municipio'])
                        self.edt_estado_modify.setText(obj['estado'])
                        self.edt_cp_modify.setText(str(obj['cp']))
                        self.edt_tel_modify.setText(str(obj['tel']))
                        self.edt_email_modify.setText(obj['email'])
                        self.edt_fnac_modify.setText(obj['fnac'])
            else:
                self.error404()

    #Guardar el Dic y actualizar los datos en el json
    def mod_rec(self):
        if fname:
            data = read_data(fname[0])
            dic = {
                "name": self.edt_name_modify.text(),
                "apellido1":  self.edt_apellido1_modify.text(),
                "apellido2": self.edt_apellido2_modify.text(),
                "cargo": self.edt_cargo_modify.text(),
                "empresa": self.edt_empresa_modify.text(),
                "calle": self.edt_calle_modify.text(),
                "noExt": int(self.edt_ext_modify.text()),
                "noInt": int(self.edt_int_modify.text()),
                "colonia": self.edt_col_modify.text(),
                "municipio": self.edt_mun_modify.text(),
                "estado":  self.edt_estado_modify.text(),
                "cp": int(self.edt_cp_modify.text()),
                "tel": int(self.edt_tel_modify.text()),
                "email": self.edt_email_modify.text(),
                "fnac": self.edt_fnac_modify.text()
            }
            mod_recipient(data, dic, self.edt_search_modify.text())
            self.lbl_mod_msg.setText("Modificado con Exito")
            self.clean_mod()
        else:
            self.error404()

######################################## READ ##################################

    #Cargar los datos en la tabla
    def load_table(self):
        if fname:
            data = read_data(fname[0])
            row = 0
            self.table_data.setRowCount(len(data["recipient"]))
            for recipient, obj in enumerate(data["recipient"]):
                self.table_data.setItem(row, 0, QtWidgets.QTableWidgetItem(obj["name"]))                
                self.table_data.setItem(row, 1, QtWidgets.QTableWidgetItem(obj["apellido1"]))                
                self.table_data.setItem(row, 2, QtWidgets.QTableWidgetItem(obj["apellido2"]))                
                self.table_data.setItem(row, 3, QtWidgets.QTableWidgetItem(obj["cargo"]))                
                self.table_data.setItem(row, 4, QtWidgets.QTableWidgetItem(obj["empresa"]))                
                self.table_data.setItem(row, 5, QtWidgets.QTableWidgetItem(obj["calle"]))                
                self.table_data.setItem(row, 6, QtWidgets.QTableWidgetItem(str(obj["noExt"])))                
                self.table_data.setItem(row, 7, QtWidgets.QTableWidgetItem(str(obj["noInt"])))                
                self.table_data.setItem(row, 8, QtWidgets.QTableWidgetItem(obj["colonia"]))                
                self.table_data.setItem(row, 9, QtWidgets.QTableWidgetItem(obj["municipio"]))                
                self.table_data.setItem(row, 10, QtWidgets.QTableWidgetItem(obj["estado"]))                
                self.table_data.setItem(row, 11, QtWidgets.QTableWidgetItem(str(obj["cp"])))                
                self.table_data.setItem(row, 12, QtWidgets.QTableWidgetItem(str(obj["tel"])))                
                self.table_data.setItem(row, 13, QtWidgets.QTableWidgetItem(obj["email"]))                
                self.table_data.setItem(row, 14, QtWidgets.QTableWidgetItem(obj["fnac"]))                
                row = row + 1
        else:
            self.error404()

######################################## DELETE ##################################

    #Cargar los datos en la tabla de eliminación
    def load_table_delete(self):
        if fname:
            data = read_data(fname[0])
            row = 0
            self.table_data_delete.setRowCount(len(data["recipient"]))
            for recipient, obj in enumerate(data["recipient"]):
                self.table_data_delete.setItem(row, 0, QtWidgets.QTableWidgetItem(obj["name"]))                
                self.table_data_delete.setItem(row, 1, QtWidgets.QTableWidgetItem(obj["apellido1"]))                
                self.table_data_delete.setItem(row, 2, QtWidgets.QTableWidgetItem(obj["apellido2"]))                
                self.table_data_delete.setItem(row, 3, QtWidgets.QTableWidgetItem(obj["cargo"]))                
                self.table_data_delete.setItem(row, 4, QtWidgets.QTableWidgetItem(obj["empresa"]))                
                self.table_data_delete.setItem(row, 5, QtWidgets.QTableWidgetItem(obj["calle"]))                
                self.table_data_delete.setItem(row, 6, QtWidgets.QTableWidgetItem(str(obj["noExt"])))                
                self.table_data_delete.setItem(row, 7, QtWidgets.QTableWidgetItem(str(obj["noInt"])))                
                self.table_data_delete.setItem(row, 8, QtWidgets.QTableWidgetItem(obj["colonia"]))                
                self.table_data_delete.setItem(row, 9, QtWidgets.QTableWidgetItem(obj["municipio"]))                
                self.table_data_delete.setItem(row, 10, QtWidgets.QTableWidgetItem(obj["estado"]))                
                self.table_data_delete.setItem(row, 11, QtWidgets.QTableWidgetItem(str(obj["cp"])))                
                self.table_data_delete.setItem(row, 12, QtWidgets.QTableWidgetItem(str(obj["tel"])))                
                self.table_data_delete.setItem(row, 13, QtWidgets.QTableWidgetItem(obj["email"]))                
                self.table_data_delete.setItem(row, 14, QtWidgets.QTableWidgetItem(obj["fnac"]))                
                row = row + 1
        else:
            self.error404()

    #Borrar un destinatario
    def del_rec(self):
        if fname:
            search_email = self.edt_search_delete.text()
            data = read_data(fname[0])
            del_recipient(data, search_email)
            self.load_table_delete()
            self.clean_del()
            self.lbl_delete_msg.setText("Eliminado con Exito")
        else:
            self.error404()

######################################## GEN DOCs ##################################

    # Generar txt
    def gen_txt(self):
        if fname and fnameT:
            data = read_data(fname[0])
            fullTemplate = read_template(fnameT[0])
            if self.edt_emails_doc.text() == "":
                for recipient, obj in enumerate(data["recipient"]):
                    generate_txt(obj, fullTemplate)  
            else:
                elements = self.edt_emails_doc.text()
                list_elements = list(elements.split(","))
                for recipient, obj in enumerate(data["recipient"]):
                    if obj['email'] in list_elements:
                        generate_txt(obj, fullTemplate)    
        else:
            self.error404()

    # Generar PDF
    def gen_pdf(self):
        if fname and fnameT:
            data = read_data(fname[0])
            fullTemplate = read_template(fnameT[0])
            if self.edt_emails_doc.text() == "":
                for recipient, obj in enumerate(data["recipient"]):
                    generate_pdf(obj, fullTemplate)  
            else:
                elements = self.edt_emails_doc.text()
                list_elements = list(elements.split(","))
                for recipient, obj in enumerate(data["recipient"]):
                    if obj['email'] in list_elements:
                        generate_pdf(obj, fullTemplate)    
        else:
            self.error404()

    # Generar Docx
    def gen_docx(self):
        if fname and fnameT:
            data = read_data(fname[0])
            fullTemplate = read_template(fnameT[0])
            if self.edt_emails_doc.text() == "":
                for recipient, obj in enumerate(data["recipient"]):
                    generate_docx(obj, fullTemplate)  
            else:
                elements = self.edt_emails_doc.text()
                list_elements = list(elements.split(","))
                for recipient, obj in enumerate(data["recipient"]):
                    if obj['email'] in list_elements:
                        generate_docx(obj, fullTemplate)    
        else:
            self.error404()

    # Generar Email
    def gen_email(self):
        if fname and fnameT:
            data = read_data(fname[0])
            fullTemplate = read_template(fnameT[0])
            if self.edt_emails_doc.text() == "":
                for recipient, obj in enumerate(data["recipient"]):
                    generate_mails(obj, fullTemplate)  
            else:
                elements = self.edt_emails_doc.text()
                list_elements = list(elements.split(","))
                for recipient, obj in enumerate(data["recipient"]):
                    if obj['email'] in list_elements:
                        generate_mails(obj, fullTemplate)    
        else:
            self.error404()

    # Cargar Tabla de docs
    def load_table_doc(self):
        if fname:
            data = read_data(fname[0])
            row = 0
            self.table_data_doc.setRowCount(len(data["recipient"]))
            for recipient, obj in enumerate(data["recipient"]):
                self.table_data_doc.setItem(row, 0, QtWidgets.QTableWidgetItem(obj["name"]))                
                self.table_data_doc.setItem(row, 1, QtWidgets.QTableWidgetItem(obj["apellido1"]))                
                self.table_data_doc.setItem(row, 2, QtWidgets.QTableWidgetItem(obj["apellido2"]))                
                self.table_data_doc.setItem(row, 3, QtWidgets.QTableWidgetItem(obj["cargo"]))                
                self.table_data_doc.setItem(row, 4, QtWidgets.QTableWidgetItem(obj["empresa"]))                
                self.table_data_doc.setItem(row, 5, QtWidgets.QTableWidgetItem(obj["calle"]))                
                self.table_data_doc.setItem(row, 6, QtWidgets.QTableWidgetItem(str(obj["noExt"])))                
                self.table_data_doc.setItem(row, 7, QtWidgets.QTableWidgetItem(str(obj["noInt"])))                
                self.table_data_doc.setItem(row, 8, QtWidgets.QTableWidgetItem(obj["colonia"]))                
                self.table_data_doc.setItem(row, 9, QtWidgets.QTableWidgetItem(obj["municipio"]))                
                self.table_data_doc.setItem(row, 10, QtWidgets.QTableWidgetItem(obj["estado"]))                
                self.table_data_doc.setItem(row, 11, QtWidgets.QTableWidgetItem(str(obj["cp"])))                
                self.table_data_doc.setItem(row, 12, QtWidgets.QTableWidgetItem(str(obj["tel"])))                
                self.table_data_doc.setItem(row, 13, QtWidgets.QTableWidgetItem(obj["email"]))                
                self.table_data_doc.setItem(row, 14, QtWidgets.QTableWidgetItem(obj["fnac"]))                
                row = row + 1
        else:
            self.error404()


####################################### LIMPIAR CAMPOS ##################################

    # Limpiar Registro
    def clean_add(self):
        self.edt_name.clear()
        self.edt_apellido1.clear()
        self.edt_apellido2.clear()
        self.edt_cargo.clear()
        self.edt_empresa.clear()
        self.edt_calle.clear()
        self.edt_ext.clear()
        self.edt_int.clear()
        self.edt_col.clear()
        self.edt_mun.clear()
        self.edt_estado.clear()
        self.edt_cp.clear()#
        self.edt_tel.clear()
        self.edt_fnac.clear()

    # Limpiar Actualizaciones
    def clean_mod(self):
        self.edt_search_modify.clear()
        self.edt_name_modify.clear()
        self.edt_apellido1_modify.clear()
        self.edt_apellido2_modify.clear()
        self.edt_cargo_modify.clear()
        self.edt_empresa_modify.clear()
        self.edt_calle_modify.clear()
        self.edt_ext_modify.clear()
        self.edt_int_modify.clear()
        self.edt_col_modify.clear()
        self.edt_mun_modify.clear()
        self.edt_estado_modify.clear()
        self.edt_cp_modify.clear()
        self.edt_tel_modify.clear()
        self.edt_email_modify.clear()
        self.edt_fnac_modify.clear()

    # Limpiar Delete
    def clean_del(self):
        self.edt_search_delete.clear()

    # Instrucciones para generar Docs
    def advice(self):
        dlg = QMessageBox(self)
        dlg.setText("Para generar los Docs ingrese los correos electronicos separados por una coma")
        dlg.setWindowTitle("Instrucciones pt.1")
        dlg.exec()
        dlg_all = QMessageBox(self)
        dlg_all.setText("Para generar todos los Docs deje el campo en blanco")
        dlg_all.setWindowTitle("Instrucciones pt.2")
        dlg_all.exec()

    #En caso de que haya un error  
    def error404(self):
        dlg = QMessageBox(self)
        dlg.setText("File not Found")
        dlg.setWindowTitle("File not Found")
        dlg.exec()
    
# Para abrir la ventana apartir del constructor
if __name__ == '__main__':
    app = QApplication(sys.argv)
    view = PrincipalView()
    view.show()
    sys.exit(app.exec())