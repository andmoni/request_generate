#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''

'''


try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET
#import xml.etree.ElementTree as ET



# ТЕКСТ НИЖЕ НИ В КОЕМ СЛУЧАЕ НЕЛЬЗЯ РЕДАКТИРОВАТЬ!

info = '''
Created on 18 октября 2016 г.
Relise beta-test: 02.11.2016
@author: Андрей (Следователь) 

Дорогой пользователь.
Данная программа предназначен для создания запросов 
характеризующего материала на лиц по уголовным делам.
пожелания в развитии программы принимаются на
электронную почту andmoni@yandex.ru
'''

VERSION = '0.01b'

from PyQt4 import QtGui, QtCore
import pickle
import datetime
import tempfile
import os
import shutil
import zipfile


class docx():
    def __init__(self, filemane, 
                 replace_text, 
                 output_filename
                 ):
        self.zipfile = zipfile.ZipFile(filemane)
        self.old_content = self.zipfile.open('word/document.xml')
        self.tree = ET.parse(self.old_content) 
        self.root = self.tree.getroot() 
        self.node_t = []
        self._findeTextNode(self.root)
        self._replace(replace_text)
        self._saveAndClose(output_filename, 0)
        
    def _findeTextNode(self, root):
        for node in root:
            if len(node)>0:
                self._findeTextNode(node)
            if node.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t':
                self.node_t.append(node)
        
    def _replace(self, replase_text):
        for key in replase_text.keys():
            for node in self.node_t:
                if node.text.strip() == key:
                    node.text = replase_text[key]

    def _saveAndClose(self, output_filename, trying):
        tmp_dir = tempfile.mkdtemp()
        self.zipfile.extractall(tmp_dir)
        with open (os.path.join(tmp_dir, 'word/document.xml'), 'wb') as f:
            f.write(ET.tostring(self.root))
        filenames = self.zipfile.namelist()
        try:
            with zipfile.ZipFile(output_filename, 'w') as docx:
                for filename in filenames:
                    docx.write(os.path.join(tmp_dir, filename), filename)
        except PermissionError:
            trying += 1
            self._saveAndClose(output_filename[:-5] + '({})'.format(trying)+'.docx', trying)
                
        shutil.rmtree(tmp_dir)
    
class RequestTabWidget (QtGui.QDialog):
    '''
    класс виджета запроса
    доработать список хранения:
    - добавить выбор подписи лично-лично+начальник-начальник
    - добавить галочку использования фонарика
    '''
    def __init__(self, request_data, parent=None):
        QtGui.QWidget.__init__(self)
        self.request_data = request_data
        self.init()
        self.setData()
        self.setWindowTitle('Изменение запроса')

    def init(self):
        self.vbox_0 = QtGui.QVBoxLayout()
        self.Gbox = QtGui.QGroupBox('Запрос в {}'.format(self.request_data[0]))
        short_name = QtGui.QLabel('Краткое название :')
        self.le_short_name = QtGui.QLineEdit()
        self.hbox_00 = QtGui.QHBoxLayout()
        self.hbox_00.addWidget(short_name)
        self.hbox_00.addWidget(self.le_short_name)
        self.box_0 = QtGui.QGroupBox('Адресат :')
        self.vbox_1 = QtGui.QVBoxLayout()
        self.le_appointment = QtGui.QLineEdit()
        self.hbox_0 = QtGui.QHBoxLayout()
        self.hbox_0.addWidget(QtGui.QLabel('Должность :'))
        self.hbox_0.addWidget(self.le_appointment)
        self.vbox_1.addLayout(self.hbox_0)
        self.le_name_organization = QtGui.QLineEdit()
        self.hbox_1 = QtGui.QHBoxLayout()
        self.hbox_1.addWidget(QtGui.QLabel('Наименование организации :'))
        self.hbox_1.addWidget(self.le_name_organization)
        self.vbox_1.addLayout(self.hbox_1)
        self.le_rang = QtGui.QLineEdit()
        self.le_surname = QtGui.QLineEdit()
        self.le_initial = QtGui.QLineEdit()
        self.le_initial.setMaximumWidth(50)
        self.hbox_2 = QtGui.QHBoxLayout()
        self.hbox_2.addWidget(QtGui.QLabel('Класный чин или звание :'))
        self.hbox_2.addWidget(self.le_rang)
        self.hbox_2.addWidget(QtGui.QLabel('Фамилия :'))
        self.hbox_2.addWidget(self.le_surname)
        self.hbox_2.addWidget(QtGui.QLabel('Инициалы :'))
        self.hbox_2.addWidget(self.le_initial)
        self.vbox_1.addLayout(self.hbox_2)
        self.box_0.setLayout(self.vbox_1)
        self.vbox_0.addWidget(self.box_0)
        self.box_1 = QtGui.QGroupBox('Адрес :')
        self.vbox_2 = QtGui.QVBoxLayout()
        self.le_index = QtGui.QLineEdit()
        self.le_index.setMaximumWidth(120)
        self.le_country = QtGui.QLineEdit()
        self.le_country.setText('РФ')
        self.le_country.setReadOnly(True)
        self.le_province = QtGui.QLineEdit()
        self.le_district = QtGui.QLineEdit()
        self.hbox_3 = QtGui.QHBoxLayout()
        self.hbox_3.addWidget(QtGui.QLabel('Индекс :'))
        self.hbox_3.addWidget(self.le_index)
        self.hbox_3.addWidget(QtGui.QLabel('Страна :'))
        self.hbox_3.addWidget(self.le_country)
        self.hbox_3.addWidget(QtGui.QLabel('Область, край, республика :'))
        self.hbox_3.addWidget(self.le_province)
        self.hbox_3.addWidget(QtGui.QLabel('Муниципальный район :'))
        self.hbox_3.addWidget(self.le_district)
        self.vbox_2.addLayout(self.hbox_3)
        self.le_town = QtGui.QLineEdit()
        self.le_street = QtGui.QLineEdit()
        self.le_house = QtGui.QLineEdit()
        self.le_house.setMaximumWidth(50)
        self.le_housein = QtGui.QLineEdit()
        self.le_housein.setMaximumWidth(50)
        self.le_corps = QtGui.QLineEdit()
        self.le_corps.setMaximumWidth(50)
        self.le_apartment = QtGui.QLineEdit()
        self.le_apartment.setMaximumWidth(50)
        self.hbox_4 = QtGui.QHBoxLayout()
        self.hbox_4.addWidget(QtGui.QLabel('Населенный пункт :'))
        self.hbox_4.addWidget(self.le_town)
        self.hbox_4.addWidget(QtGui.QLabel('Улица :'))
        self.hbox_4.addWidget(self.le_street)
        self.hbox_4.addWidget(QtGui.QLabel('Дом :'))
        self.hbox_4.addWidget(self.le_house)
        self.hbox_4.addWidget(QtGui.QLabel('Строение :'))
        self.hbox_4.addWidget(self.le_housein)
        self.hbox_4.addWidget(QtGui.QLabel('Корпус :'))
        self.hbox_4.addWidget(self.le_corps)
        self.hbox_4.addWidget(QtGui.QLabel('Офис :'))
        self.hbox_4.addWidget(self.le_apartment)
        self.vbox_2.addLayout(self.hbox_4)
        self.box_1.setLayout(self.vbox_2)
        self.vbox_0.addWidget(self.box_1)
        self.Gbox.setLayout(self.vbox_0)
        self.gbox_2 = QtGui.QGroupBox('Формулировка запрашиваемых сведений :')
        self.vbox_3 = QtGui.QVBoxLayout()
        self.gbox_2.setLayout(self.vbox_3)
        self.hbox_5 = QtGui.QHBoxLayout()
        self.le_need_info = QtGui.QLineEdit()
        self.hbox_5.addWidget(self.le_need_info)
        self.vbox_3.addLayout(self.hbox_5)
        self.vbox_0.addWidget(self.gbox_2)
        self.vbox_3 = QtGui.QVBoxLayout()
        self.vbox_3.addLayout(self.hbox_00)
        self.vbox_3.addWidget(self.Gbox)
        self.vbox_3.addLayout(self.hbox_5)
        self.pb_box = QtGui.QDialogButtonBox()
        self.pb_box.addButton('Сохранить',
                        QtGui.QDialogButtonBox.AcceptRole)
        self.pb_box.addButton('Отменить',
                        QtGui.QDialogButtonBox.RejectRole)
        self.pb_box.accepted.connect(self.accept)
        self.pb_box.rejected.connect(self.reject)
        self.vbox_3.addWidget(self.pb_box)
        self.vbox_3.addStretch()
        self.setLayout(self.vbox_3)

    def setData(self):
        #self.le_country.setReadOnly(False)
        self.le_short_name.setText(self.request_data[0])
        self.le_appointment.setText(self.request_data[1])
        self.le_name_organization.setText(self.request_data[2])
        self.le_rang.setText(self.request_data[3])
        self.le_surname.setText(self.request_data[4])
        self.le_initial.setText(self.request_data[5])
        self.le_index.setText(self.request_data[6])
        #self.le_country.setText(self.request_data[7])
        self.le_province.setText(self.request_data[8])
        self.le_district.setText(self.request_data[9])
        self.le_town.setText(self.request_data[10])
        self.le_street.setText(self.request_data[11])
        self.le_house.setText(self.request_data[12])
        self.le_housein.setText(self.request_data[13])
        self.le_corps.setText(self.request_data[14])
        self.le_apartment.setText(self.request_data[15])
        self.le_need_info.setText(self.request_data[16])
        self.le_short_name.textChanged.connect(self.changeName)
        
    def changeName(self):
        self.Gbox.setTitle('Запрос {}'.format(self.le_short_name.text()))
    

class RequestOverheadDataWidget (QtGui.QWidget):
    '''
    виджет служебной информации
    добавить:
        - список для автозаполнения спец звания 
        - список для автозаполнение дистриктов (край область респ.)
        - подумать над прикручиванием автоопределения индексов 
            (автоматическое определение в зависимости от выбранной улицы и дома)
            и в обратном порядке (по индексу вставка данных в поля кроме улицы и дома)
    '''
    def __init__(self, setting = None, parent=None):
        QtGui.QWidget.__init__(self, parent)   
        self.init()
        if setting != None: self.setData(setting)
    
    def init(self):
        self.vbox_0 = QtGui.QVBoxLayout()
        self.setLayout(self.vbox_0)
        self.gbox_0 = QtGui.QGroupBox('Данные для углового штампа :')
        self.vbox_00 = QtGui.QVBoxLayout()
        self.gbox_0.setLayout(self.vbox_00)
        self.fbox_0 = QtGui.QFormLayout ()
        self.le_short_title_superior = QtGui.QLineEdit ()
        #self.le_short_title_superior.setMinimumWidth(400)
        self.fbox_0.addRow ('Краткое название вышестоящего ОВД* :', self.le_short_title_superior)
        self.le_full_title = QtGui.QLineEdit ()
        #self.le_full_title.setMinimumWidth(400)
        self.fbox_0.addRow ('Полное наименование органа МВД* :', self.le_full_title)
        self.le_short_title = QtGui.QLineEdit ()
        #self.le_short_title.setMinimumWidth(400)
        self.fbox_0.addRow ('Сокращенное наименование органа МВД* :', self.le_short_title)
        self.le_full_unit_name = QtGui.QLineEdit ()
        #self.le_full_unit_name.setMinimumWidth(400)
        self.fbox_0.addRow ('Полное наименование подразделения* :', self.le_full_unit_name)
        self.le_short_unit_name = QtGui.QLineEdit ()
        #self.le_short_unit_name.setMinimumWidth(400)
        self.fbox_0.addRow ('Сокращенное наименование подразделения* :', self.le_short_unit_name)
        self.fbox_0.setLabelAlignment(QtCore.Qt.AlignRight)
        self.fbox_0.setAlignment(QtCore.Qt.AlignJustify)
        self.vbox_00.addLayout(self.fbox_0)
        self.gbox_1 = QtGui.QGroupBox('Адрес подразделения:')
        self.vbox_1 = QtGui.QVBoxLayout()
        self.gbox_1.setLayout(self.vbox_1)
        self.le_index_0 = QtGui.QLineEdit()
        self.le_index_0.setMaximumWidth(120)
        self.le_country_0 = QtGui.QLineEdit()
        self.le_province_0 = QtGui.QLineEdit()
        self.le_district_0 = QtGui.QLineEdit()
        self.hbox_1 = QtGui.QHBoxLayout()
        self.hbox_1.addWidget(QtGui.QLabel('Индекс* :'))
        self.hbox_1.addWidget(self.le_index_0)
        self.hbox_1.addWidget(QtGui.QLabel('Страна :'))
        self.hbox_1.addWidget(self.le_country_0)
        self.hbox_1.addWidget(QtGui.QLabel('Область, край, республика* :'))
        self.hbox_1.addWidget(self.le_province_0)
        self.hbox_1.addWidget(QtGui.QLabel('Район :'))
        self.hbox_1.addWidget(self.le_district_0)
        self.vbox_1.addLayout(self.hbox_1)
        self.le_town_0 = QtGui.QLineEdit()
        self.le_street_0 = QtGui.QLineEdit()
        self.le_house_0 = QtGui.QLineEdit()
        self.le_house_0.setMaximumWidth(50)
        self.le_housein_0 = QtGui.QLineEdit()
        self.le_housein_0.setMaximumWidth(50)
        self.le_corps_0 = QtGui.QLineEdit()
        self.le_corps_0.setMaximumWidth(50)
        self.le_apartment_0 = QtGui.QLineEdit()
        self.le_apartment_0.setMaximumWidth(50)
        self.hbox_2 = QtGui.QHBoxLayout()
        self.hbox_2.addWidget(QtGui.QLabel('Населенный пункт* :'))
        self.hbox_2.addWidget(self.le_town_0)
        self.hbox_2.addWidget(QtGui.QLabel('Улица* :'))
        self.hbox_2.addWidget(self.le_street_0)
        self.hbox_2.addWidget(QtGui.QLabel('Дом* :'))
        self.hbox_2.addWidget(self.le_house_0)
        self.hbox_2.addWidget(QtGui.QLabel('Строение :'))
        self.hbox_2.addWidget(self.le_housein_0)
        self.hbox_2.addWidget(QtGui.QLabel('Корпус :'))
        self.hbox_2.addWidget(self.le_corps_0)
        self.hbox_2.addWidget(QtGui.QLabel('Офис :'))
        self.hbox_2.addWidget(self.le_apartment_0)
        self.vbox_1.addLayout(self.hbox_2)
        self.le_phone_0 = QtGui.QLineEdit()
        self.le_fax_0 = QtGui.QLineEdit()
        self.le_email_0 = QtGui.QLineEdit()
        self.hbox_3 = QtGui.QHBoxLayout()
        self.hbox_3.addWidget(QtGui.QLabel('Телефон д\ч* :'))
        self.hbox_3.addWidget(self.le_phone_0)
        self.hbox_3.addWidget(QtGui.QLabel('Факс :'))
        self.hbox_3.addWidget(self.le_fax_0)
        self.hbox_3.addWidget(QtGui.QLabel('Адрес электронной почты :'))
        self.hbox_3.addWidget(self.le_email_0)
        self.vbox_1.addLayout(self.hbox_3)
        self.vbox_00.addWidget(self.gbox_1)
        self.vbox_0.addWidget(self.gbox_0)
        self.hbox_30 = QtGui.QHBoxLayout()
        self.hbox_30.addWidget(QtGui.QLabel('Наименование подразделения в дательном падеже (кому)* :'))
        self.le_short_title_dat = QtGui.QLineEdit()
        self.hbox_30.addWidget(self.le_short_title_dat)
        self.vbox_0.addLayout(self.hbox_30)
        self.hbox_31 = QtGui.QHBoxLayout()
        self.hbox_31.addWidget(QtGui.QLabel('Наименование подразделения в родительном падеже (кого)* :'))
        self.le_short_title_rod = QtGui.QLineEdit()
        self.hbox_31.addWidget(self.le_short_title_rod)
        self.vbox_0.addLayout(self.hbox_31)
        self.gbox_2 = QtGui.QGroupBox('Начальник подразделения :')
        self.vbox_2 = QtGui.QVBoxLayout()
        self.le_position_head = QtGui.QLineEdit()
        self.le_rank_head = QtGui.QLineEdit()
        self.le_surname_head = QtGui.QLineEdit()
        self.le_name_head = QtGui.QLineEdit()
        self.le_patronymic_head = QtGui.QLineEdit()
        self.hbox_4 = QtGui.QHBoxLayout()
        self.hbox_4.addWidget(QtGui.QLabel('Должность* :'))
        self.hbox_4.addWidget(self.le_position_head)
        self.hbox_4.addWidget(QtGui.QLabel('Специальное звание* :'))
        self.hbox_4.addWidget(self.le_rank_head)
        self.vbox_2.addLayout(self.hbox_4)
        self.hbox_5 = QtGui.QHBoxLayout()
        self.hbox_5.addWidget(QtGui.QLabel('Фамилия* :'))
        self.hbox_5.addWidget(self.le_surname_head)
        self.hbox_5.addWidget(QtGui.QLabel('Имя* :'))
        self.hbox_5.addWidget(self.le_name_head)
        self.hbox_5.addWidget(QtGui.QLabel('Отчество* :'))
        self.hbox_5.addWidget(self.le_patronymic_head)
        self.vbox_2.addLayout(self.hbox_5)
        self.gbox_2.setLayout(self.vbox_2)
        self.vbox_0.addWidget(self.gbox_2)
        self.gbox_3 = QtGui.QGroupBox('Исполнитель запроса :')
        self.vbox_3 = QtGui.QVBoxLayout()
        self.le_position_executive = QtGui.QLineEdit()
        self.le_rank_executive = QtGui.QLineEdit()
        self.le_surname_executive = QtGui.QLineEdit()
        self.le_name_executive = QtGui.QLineEdit()
        self.le_patronymic_executive = QtGui.QLineEdit()
        self.hbox_6 = QtGui.QHBoxLayout()
        self.hbox_6.addWidget(QtGui.QLabel('Должность* :'))
        self.hbox_6.addWidget(self.le_position_executive)
        self.hbox_6.addWidget(QtGui.QLabel('Специальное звание :'))
        self.hbox_6.addWidget(self.le_rank_executive)
        self.vbox_3.addLayout(self.hbox_6)
        self.hbox_7 = QtGui.QHBoxLayout()
        self.hbox_7.addWidget(QtGui.QLabel('Фамилия* :'))
        self.hbox_7.addWidget(self.le_surname_executive)
        self.hbox_7.addWidget(QtGui.QLabel('Имя* :'))
        self.hbox_7.addWidget(self.le_name_executive)
        self.hbox_7.addWidget(QtGui.QLabel('Отчество* :'))
        self.hbox_7.addWidget(self.le_patronymic_executive)
        self.vbox_3.addLayout(self.hbox_7)
        self.le_phone_executive = QtGui.QLineEdit()
        self.le_fax_executive = QtGui.QLineEdit()
        self.le_email_executive = QtGui.QLineEdit()
        self.hbox_8 = QtGui.QHBoxLayout()
        self.hbox_8.addWidget(QtGui.QLabel('Контактный телефон :'))
        self.hbox_8.addWidget(self.le_phone_executive)
        self.hbox_8.addWidget(QtGui.QLabel('Факс :'))
        self.hbox_8.addWidget(self.le_fax_executive)
        self.hbox_8.addWidget(QtGui.QLabel('Адрес электронной почты :'))
        self.hbox_8.addWidget(self.le_email_executive)
        self.vbox_3.addLayout(self.hbox_8)
        self.gbox_3.setLayout(self.vbox_3)
        self.vbox_0.addWidget(self.gbox_3)
        self.vbox_0.addStretch()
    
    def setData(self, setting):
        #self.swichConnect(False)
        self.le_short_title_superior.setText(setting[0][0])
        self.le_full_title.setText(setting[0][1])
        self.le_short_title.setText(setting[0][2])
        self.le_full_unit_name.setText(setting[0][3])
        self.le_short_unit_name.setText(setting[0][4])
        self.le_index_0.setText(setting[0][5])
        self.le_country_0.setText(setting[0][6])
        self.le_province_0.setText(setting[0][7])
        self.le_district_0.setText(setting[0][8])
        self.le_town_0.setText(setting[0][9])
        self.le_street_0.setText(setting[0][10])
        self.le_house_0.setText(setting[0][11])
        self.le_housein_0.setText(setting[0][12])
        self.le_corps_0.setText(setting[0][13])
        self.le_apartment_0.setText(setting[0][14])
        self.le_phone_0.setText(setting[0][15])
        self.le_fax_0.setText(setting[0][16])
        self.le_email_0.setText(setting[0][17])
        self.le_short_title_dat.setText(setting[0][18])
        self.le_short_title_rod.setText(setting[0][19])
        self.le_position_head.setText(setting[1][0])
        self.le_rank_head.setText(setting[1][1])
        self.le_surname_head.setText(setting[1][2])
        self.le_name_head.setText(setting[1][3])
        self.le_patronymic_head.setText(setting[1][4])
        self.le_position_executive.setText(setting[2][0])
        self.le_rank_executive.setText(setting[2][1])
        self.le_surname_executive.setText(setting[2][2])
        self.le_name_executive.setText(setting[2][3])
        self.le_patronymic_executive.setText(setting[2][4])
        self.le_phone_executive.setText(setting[2][5])
        self.le_fax_executive.setText(setting[2][6])
        self.le_email_executive.setText(setting[2][7])
        self.swichConnect()
        
    def verification(self):
        for item in [self.le_short_title_superior,
                     self.le_full_title,
                     self.le_short_title,
                     self.le_full_unit_name,
                     self.le_short_unit_name,
                     self.le_index_0,
                     self.le_province_0,
                     self.le_town_0,
                     self.le_street_0,
                     self.le_house_0,
                     self.le_phone_0]:
            if item.text() == '':
                QtGui.QMessageBox.information(self, "Внимание",
                    'Недостаточно сведений для углового штампа.\nЗаполните все поля отмеченые звездочкой.')
                return True
        for item in [self.le_position_head,
                     self.le_rank_head,
                     self.le_surname_head,
                     self.le_name_head,
                     self.le_patronymic_head]:
            if item.text() == '':
                QtGui.QMessageBox.information(self, "Внимание",
                    'Недостаточно сведений о руководстве подразделения.\nЗаполните все поля отмеченые звездочкой.')
                return True
        for item in [self.le_position_executive,
                     self.le_rank_executive,
                     self.le_surname_executive,
                     self.le_name_executive,
                     self.le_patronymic_executive]:
            if item.text() == '':
                QtGui.QMessageBox.information(self, "Внимание",
                    'Недостаточно сведений об исполнителе (инициаторе) запроса.\nЗаполните все поля отмеченые звездочкой.')
                return True
        return False
                
    def swichConnect(self):
        self.le_full_title.editingFinished.connect(self.saveData)
        self.le_short_title.editingFinished.connect(self.saveData)
        self.le_full_unit_name.editingFinished.connect(self.saveData)
        self.le_short_unit_name.editingFinished.connect(self.saveData)
        self.le_index_0.editingFinished.connect(self.saveData)
        self.le_country_0.editingFinished.connect(self.saveData)
        self.le_province_0.editingFinished.connect(self.saveData)
        self.le_district_0.editingFinished.connect(self.saveData)
        self.le_town_0.editingFinished.connect(self.saveData)
        self.le_street_0.editingFinished.connect(self.saveData)
        self.le_house_0.editingFinished.connect(self.saveData)
        self.le_housein_0.editingFinished.connect(self.saveData)
        self.le_corps_0.editingFinished.connect(self.saveData)
        self.le_apartment_0.editingFinished.connect(self.saveData)
        self.le_phone_0.editingFinished.connect(self.saveData)
        self.le_fax_0.editingFinished.connect(self.saveData)
        self.le_email_0.editingFinished.connect(self.saveData)
        self.le_position_head.editingFinished.connect(self.saveData)
        self.le_rank_head.editingFinished.connect(self.saveData)
        self.le_surname_head.editingFinished.connect(self.saveData)
        self.le_name_head.editingFinished.connect(self.saveData)
        self.le_patronymic_head.editingFinished.connect(self.saveData)
        self.le_position_executive.editingFinished.connect(self.saveData)
        self.le_rank_executive.editingFinished.connect(self.saveData)
        self.le_surname_executive.editingFinished.connect(self.saveData)
        self.le_name_executive.editingFinished.connect(self.saveData)
        self.le_patronymic_executive.editingFinished.connect(self.saveData)
        self.le_phone_executive.editingFinished.connect(self.saveData)
        self.le_fax_executive.editingFinished.connect(self.saveData)
        self.le_email_executive.editingFinished.connect(self.saveData)
        self.le_short_title_dat.editingFinished.connect(self.saveData)
        self.le_short_title_rod.editingFinished.connect(self.saveData)

    def saveData(self):
        try:
            setting_file = open('setting', 'wb')
            setting = [[self.le_short_title_superior.text(),
                          self.le_full_title.text(),
                          self.le_short_title.text(),
                          self.le_full_unit_name.text(),
                          self.le_short_unit_name.text(),
                          self.le_index_0.text(),
                          self.le_country_0.text(),
                          self.le_province_0.text(),
                          self.le_district_0.text(),
                          self.le_town_0.text(),
                          self.le_street_0.text(),
                          self.le_house_0.text(),
                          self.le_housein_0.text(),
                          self.le_corps_0.text(),
                          self.le_apartment_0.text(),
                          self.le_phone_0.text(),
                          self.le_fax_0.text(),
                          self.le_email_0.text(),
                          self.le_short_title_dat.text(),
                          self.le_short_title_rod.text()],
                        [self.le_position_head.text(),
                          self.le_rank_head.text(),
                          self.le_surname_head.text(),
                          self.le_name_head.text(),
                          self.le_patronymic_head.text()],
                        [self.le_position_executive.text(),
                          self.le_rank_executive.text(),
                          self.le_surname_executive.text(),
                          self.le_name_executive.text(),
                          self.le_patronymic_executive.text(),
                          self.le_phone_executive.text(),
                          self.le_fax_executive.text(),
                          self.le_email_executive.text()]] 
            pickle.dump(setting, setting_file)
            setting_file.close()
        except IOError:
            QtGui.QMessageBox.information(self, "Внимание",
                    'Не удается сохранить настройки в файл')

class RequestPersonWidget (QtGui.QWidget):
    '''виджет данных о личности '''
    def __init__(self, person, parent=None):
        QtGui.QWidget.__init__(self, parent)
        self.init()
        if person != None: self.setData(person)
        
    def init(self):
        self.vbox_0 = QtGui.QVBoxLayout()
        self.setLayout(self.vbox_0)
        self.gbox_0 = QtGui.QGroupBox('Данные по уголовному делу :')
        self.vbox_1 = QtGui.QVBoxLayout()
        self.le_number_case = QtGui.QLineEdit()
        self.le_number_case.setMaximumWidth(120)
        self.de_date_case = QtGui.QDateEdit()
        self.de_date_case.setCalendarPopup(True)
        self.hbox_0 = QtGui.QHBoxLayout()
        self.hbox_0.addWidget(QtGui.QLabel('Номер уголовного дела* :'))
        self.hbox_0.addWidget(self.le_number_case)
        self.hbox_0.addWidget(QtGui.QLabel('Дата возбуждения* :'))
        self.hbox_0.addWidget(self.de_date_case)
        self.gbox_00 = QtGui.QGroupBox('Квалификация преступления (по наиболее тяжкому) :')
        self.hbox_00 = QtGui.QHBoxLayout()
        self.hbox_00.addWidget(QtGui.QLabel('Статья* :'))
        self.le_article = QtGui.QLineEdit()
        self.le_article.setMaximumWidth(60)
        self.hbox_00.addWidget(self.le_article)
        self.hbox_00.addWidget(QtGui.QLabel('Часть :'))
        self.le_part = QtGui.QLineEdit()
        self.le_part.setMaximumWidth(60)
        self.hbox_00.addWidget(self.le_part)
        self.hbox_00.addWidget(QtGui.QLabel('Пункты :'))
        self.le_point = QtGui.QLineEdit()
        self.le_point.setMaximumWidth(60)
        self.hbox_00.addWidget(self.le_point)
        self.gbox_00.setLayout(self.hbox_00)
        self.hbox_0.addWidget(self.gbox_00)
        self.hbox_0.addStretch()
        self.vbox_1.addLayout(self.hbox_0)
        self.gbox_0.setLayout(self.vbox_1)
        self.vbox_0.addWidget(self.gbox_0)
        self.gbox_1 = QtGui.QGroupBox('Данные о лице :')
        self.vbox_2 = QtGui.QVBoxLayout()
        self.gbox_1.setLayout(self.vbox_2)
        self.vbox_0.addWidget(self.gbox_1)
        self.le_surname = QtGui.QLineEdit()
        self.le_name = QtGui.QLineEdit()
        self.le_patronymic = QtGui.QLineEdit()
        self.hbox_1 = QtGui.QHBoxLayout()
        self.hbox_1.addWidget(QtGui.QLabel('Фамилия* :'))
        self.hbox_1.addWidget(self.le_surname)
        self.hbox_1.addWidget(QtGui.QLabel('Имя* :'))
        self.hbox_1.addWidget(self.le_name)
        self.hbox_1.addWidget(QtGui.QLabel('Отчество* :'))
        self.hbox_1.addWidget(self.le_patronymic)
        self.vbox_2.addLayout(self.hbox_1)
        self.de_date_ob = QtGui.QDateEdit()
        self.de_date_ob.setCalendarPopup(True)
        self.le_place_ob = QtGui.QLineEdit()
        self.hbox_2 = QtGui.QHBoxLayout()
        self.hbox_2.addWidget(QtGui.QLabel('Дата рождения* :'))
        self.hbox_2.addWidget(self.de_date_ob)
        self.hbox_2.addWidget(QtGui.QLabel('Место рождения (как в паспорте)* :'))
        self.hbox_2.addWidget(self.le_place_ob)
        self.vbox_2.addLayout(self.hbox_2)
        self.gbox_2 = QtGui.QGroupBox('Место жительства :')
        self.vbox_3 = QtGui.QVBoxLayout()
        self.le_index_2 = QtGui.QLineEdit()
        self.le_index_2.setMaximumWidth(120)
        self.le_country_2 = QtGui.QLineEdit()
        self.le_province_2 = QtGui.QLineEdit()
        self.le_district_2 = QtGui.QLineEdit()
        self.hbox_3 = QtGui.QHBoxLayout()
        self.hbox_3.addWidget(QtGui.QLabel('Индекс :'))
        self.hbox_3.addWidget(self.le_index_2)
        self.hbox_3.addWidget(QtGui.QLabel('Страна :'))
        self.hbox_3.addWidget(self.le_country_2)
        self.le_country_2.setText('РФ')
        self.le_country_2.setReadOnly(True)
        self.hbox_3.addWidget(QtGui.QLabel('Область, край, республика* :'))
        self.hbox_3.addWidget(self.le_province_2)
        self.hbox_3.addWidget(QtGui.QLabel('Район :'))
        self.hbox_3.addWidget(self.le_district_2)
        self.vbox_3.addLayout(self.hbox_3)
        self.le_town_2 = QtGui.QLineEdit()
        self.le_street_2 = QtGui.QLineEdit()
        self.le_house_2 = QtGui.QLineEdit()
        self.le_house_2.setMaximumWidth(50)
        self.le_housein_2 = QtGui.QLineEdit()
        self.le_housein_2.setMaximumWidth(50)
        self.le_corps_2 = QtGui.QLineEdit()
        self.le_corps_2.setMaximumWidth(50)
        self.le_apartment_2 = QtGui.QLineEdit()
        self.le_apartment_2.setMaximumWidth(50)
        self.hbox_4 = QtGui.QHBoxLayout()
        self.hbox_4.addWidget(QtGui.QLabel('Населенный пункт* :'))
        self.hbox_4.addWidget(self.le_town_2)
        self.hbox_4.addWidget(QtGui.QLabel('Улица* :'))
        self.hbox_4.addWidget(self.le_street_2)
        self.hbox_4.addWidget(QtGui.QLabel('Дом* :'))
        self.hbox_4.addWidget(self.le_house_2)
        self.hbox_4.addWidget(QtGui.QLabel('Строение :'))
        self.hbox_4.addWidget(self.le_housein_2)
        self.hbox_4.addWidget(QtGui.QLabel('Корпус :'))
        self.hbox_4.addWidget(self.le_corps_2)
        self.hbox_4.addWidget(QtGui.QLabel('Квартира :'))
        self.hbox_4.addWidget(self.le_apartment_2)
        self.vbox_3.addLayout(self.hbox_4)
        self.gbox_2.setLayout(self.vbox_3)
        self.vbox_2.addWidget(self.gbox_2)
        self.cb_bloc_ap = QtGui.QCheckBox('Место жительства совпадает с местом регистрации')
        self.cb_bloc_ap.setCheckState(QtCore.Qt.Checked)
        self.vbox_2.addWidget(self.cb_bloc_ap)
        self.gbox_3 = QtGui.QGroupBox('Место регистрации :')
        self.vbox_4 = QtGui.QVBoxLayout()
        self.le_index_3 = QtGui.QLineEdit()
        self.le_index_3.setMaximumWidth(120)
        self.le_country_3 = QtGui.QLineEdit()
        self.le_country_3.setText('РФ')
        self.le_country_3.setReadOnly(True)
        self.le_province_3 = QtGui.QLineEdit()
        self.le_district_3 = QtGui.QLineEdit()
        self.hbox_5 = QtGui.QHBoxLayout()
        self.hbox_5.addWidget(QtGui.QLabel('Индекс :'))
        self.hbox_5.addWidget(self.le_index_3)
        self.hbox_5.addWidget(QtGui.QLabel('Страна :'))
        self.hbox_5.addWidget(self.le_country_3)
        self.hbox_5.addWidget(QtGui.QLabel('Область, край, республика :'))
        self.hbox_5.addWidget(self.le_province_3)
        self.hbox_5.addWidget(QtGui.QLabel('Район :'))
        self.hbox_5.addWidget(self.le_district_3)
        self.vbox_4.addLayout(self.hbox_5)
        self.le_town_3 = QtGui.QLineEdit()
        self.le_street_3 = QtGui.QLineEdit()
        self.le_house_3 = QtGui.QLineEdit()
        self.le_house_3.setMaximumWidth(50)
        self.le_housein_3 = QtGui.QLineEdit()
        self.le_housein_3.setMaximumWidth(50)
        self.le_corps_3 = QtGui.QLineEdit()
        self.le_corps_3.setMaximumWidth(50)
        self.le_apartment_3 = QtGui.QLineEdit()
        self.le_apartment_3.setMaximumWidth(50)
        self.hbox_6 = QtGui.QHBoxLayout()
        self.hbox_6.addWidget(QtGui.QLabel('Населенный пункт :'))
        self.hbox_6.addWidget(self.le_town_3)
        self.hbox_6.addWidget(QtGui.QLabel('Улица :'))
        self.hbox_6.addWidget(self.le_street_3)
        self.hbox_6.addWidget(QtGui.QLabel('Дом :'))
        self.hbox_6.addWidget(self.le_house_3)
        self.hbox_6.addWidget(QtGui.QLabel('Строение :'))
        self.hbox_6.addWidget(self.le_housein_3)
        self.hbox_6.addWidget(QtGui.QLabel('Корпус :'))
        self.hbox_6.addWidget(self.le_corps_3)
        self.hbox_6.addWidget(QtGui.QLabel('Квартира :'))
        self.hbox_6.addWidget(self.le_apartment_3)
        self.vbox_4.addLayout(self.hbox_6)
        self.gbox_3.setLayout(self.vbox_4)
        self.vbox_2.addWidget(self.gbox_3)
        self.gbox_4 = QtGui.QGroupBox('Паспортные данные :')
        self.vbox_5 = QtGui.QVBoxLayout()
        self.vbox_2.addWidget(self.gbox_4)
        self.gbox_4.setLayout(self.vbox_5)
        self.le_serial = QtGui.QLineEdit()
        self.le_serial.setMaximumWidth(80)
        self.le_number = QtGui.QLineEdit()
        self.le_number.setMaximumWidth(120)
        self.de_date_issue = QtGui.QDateEdit()
        self.de_date_issue.setCalendarPopup(True)
        self.hbox_7 = QtGui.QHBoxLayout()
        self.hbox_7.addWidget(QtGui.QLabel('Серия :'))
        self.hbox_7.addWidget(self.le_serial)
        self.hbox_7.addWidget(QtGui.QLabel('Номер :'))
        self.hbox_7.addWidget(self.le_number)
        self.hbox_7.addWidget(QtGui.QLabel('Дата выдачи :'))
        self.hbox_7.addWidget(self.de_date_issue)
        self.hbox_7.addStretch()
        self.vbox_5.addLayout(self.hbox_7)
        self.le_place_issue = QtGui.QLineEdit()
        self.le_code_place_issue = QtGui.QLineEdit()
        self.le_code_place_issue.setMaximumWidth(120)
        self.hbox_8 = QtGui.QHBoxLayout()
        self.hbox_8.addWidget(QtGui.QLabel('Кем выдан :'))
        self.hbox_8.addWidget(self.le_place_issue)
        self.hbox_8.addWidget(QtGui.QLabel('Код подразделения :'))
        self.hbox_8.addWidget(self.le_code_place_issue)
        self.vbox_5.addLayout(self.hbox_8)
        self.gbox_4.setLayout(self.vbox_5)
        self.vbox_0.addStretch()
        self.swich_gbox_3()
        self.cb_bloc_ap.toggled['bool'].connect(self.swich_gbox_3)
        
    def swich_gbox_3(self):
        if self.cb_bloc_ap.isChecked():
            self.gbox_3.hide()
        else:
            self.gbox_3.show()
            
    def setData(self, person):
        self.le_number_case.setText(person[0])
        self.de_date_case.setDate(person[1])
        self.le_article.setText(person[2])
        self.le_part.setText(person[3])
        self.le_point.setText(person[4])
        self.le_surname.setText(person[5])
        self.le_name.setText(person[6])
        self.le_patronymic.setText(person[7])
        self.de_date_ob.setDate(person[8])
        self.le_place_ob.setText(person[9])
        self.le_index_2.setText(person[10])
        #self.le_country_2.setText(person[11])
        self.le_province_2.setText(person[12])
        self.le_district_2.setText(person[13])
        self.le_town_2.setText(person[14])
        self.le_street_2.setText(person[15])
        self.le_house_2.setText(person[16])
        self.le_housein_2.setText(person[17])
        self.le_corps_2.setText(person[18])
        self.le_apartment_2.setText(person[19])
        self.le_index_3.setText(person[20])
        #self.le_country_3.setText(person[21])
        self.le_province_3.setText(person[22])
        self.le_district_3.setText(person[23])
        self.le_town_3.setText(person[24])
        self.le_street_3.setText(person[25])
        self.le_house_3.setText(person[26])
        self.le_housein_3.setText(person[27])
        self.le_corps_3.setText(person[28])
        self.le_apartment_3.setText(person[29])
        self.le_serial.setText(person[30])
        self.le_number.setText(person[31])
        self.de_date_issue.setDate(person[32])
        self.le_place_issue.setText(person[33])
        self.le_code_place_issue.setText(person[34])
        self.cb_bloc_ap.setCheckState(person[35])
        self.swichConnect()
        
    def verification(self):
        for item in [self.le_number_case,
                     self.de_date_case,
                     self.le_article]:
            if item.text() == '':
                QtGui.QMessageBox.information(self, "Внимание",
                    'Недостаточно сведений об уголовном деле.\nЗаполните все поля отмеченые звездочкой.')
                return True
        for item in [self.le_surname,
                     self.le_name,
                     self.le_patronymic,
                     self.le_place_ob]:
            if item.text() == '':
                QtGui.QMessageBox.information(self, "Внимание",
                    'Недостаточно сведений о лице.\nЗаполните все поля отмеченые звездочкой.')
                return True
        return False
    
    def swichConnect(self):
        self.le_number_case.editingFinished.connect(self.saveDataP)
        self.de_date_case.editingFinished.connect(self.saveDataP)
        self.le_article.editingFinished.connect(self.saveDataP)
        self.le_part.editingFinished.connect(self.saveDataP)
        self.le_point.editingFinished.connect(self.saveDataP)
        self.le_surname.editingFinished.connect(self.saveDataP)
        self.le_name.editingFinished.connect(self.saveDataP)
        self.le_patronymic.editingFinished.connect(self.saveDataP)
        self.de_date_ob.editingFinished.connect(self.saveDataP)
        self.le_place_ob.editingFinished.connect(self.saveDataP)
        self.le_index_2.editingFinished.connect(self.saveDataP)
        #self.le_country_2.editingFinished.connect(self.saveDataP)
        self.le_province_2.editingFinished.connect(self.saveDataP)
        self.le_district_2.editingFinished.connect(self.saveDataP)
        self.le_town_2.editingFinished.connect(self.saveDataP)
        self.le_street_2.editingFinished.connect(self.saveDataP)
        self.le_house_2.editingFinished.connect(self.saveDataP)
        self.le_housein_2.editingFinished.connect(self.saveDataP)
        self.le_corps_2.editingFinished.connect(self.saveDataP)
        self.le_apartment_2.editingFinished.connect(self.saveDataP)
        self.le_index_3.editingFinished.connect(self.saveDataP)
        #self.le_country_3.editingFinished.connect(self.saveDataP)
        self.le_province_3.editingFinished.connect(self.saveDataP)
        self.le_district_3.editingFinished.connect(self.saveDataP)
        self.le_town_3.editingFinished.connect(self.saveDataP)
        self.le_street_3.editingFinished.connect(self.saveDataP)
        self.le_house_3.editingFinished.connect(self.saveDataP)
        self.le_housein_3.editingFinished.connect(self.saveDataP)
        self.le_corps_3.editingFinished.connect(self.saveDataP)
        self.le_apartment_3.editingFinished.connect(self.saveDataP)
        self.le_serial.editingFinished.connect(self.saveDataP)
        self.le_number.editingFinished.connect(self.saveDataP)
        self.de_date_issue.editingFinished.connect(self.saveDataP)
        self.le_place_issue.editingFinished.connect(self.saveDataP)
        self.le_code_place_issue.editingFinished.connect(self.saveDataP) 
    
    def saveDataP(self):
        person = [self.le_number_case.text(),
                       self.de_date_case.date().toPyDate(),
                       self.le_article.text(),
                       self.le_part.text(),
                       self.le_point.text(),
                       self.le_surname.text(),
                       self.le_name.text(),
                       self.le_patronymic.text(),
                       self.de_date_ob.date().toPyDate(),
                       self.le_place_ob.text(),
                       self.le_index_2.text(),
                       self.le_country_2.text(),
                       self.le_province_2.text(),
                       self.le_district_2.text(),
                       self.le_town_2.text(),
                       self.le_street_2.text(),
                       self.le_house_2.text(),
                       self.le_housein_2.text(),
                       self.le_corps_2.text(),
                       self.le_apartment_2.text(),
                       self.le_index_3.text(),
                       self.le_country_3.text(),
                       self.le_province_3.text(),
                       self.le_district_3.text(),
                       self.le_town_3.text(),
                       self.le_street_3.text(),
                       self.le_house_3.text(),
                       self.le_housein_3.text(),
                       self.le_corps_3.text(),
                       self.le_apartment_3.text(),
                       self.le_serial.text(),
                       self.le_number.text(),
                       self.de_date_issue.date().toPyDate(),
                       self.le_place_issue.text(),
                       self.le_code_place_issue.text(),
                       self.cb_bloc_ap.checkState()]
        try:
            person_file = open('person', 'wb')
            pickle.dump(person, person_file)
            person_file.close()
        except IOError:
                QtGui.QMessageBox.information(self, "Внимание",
                    'Не удается сохранить данные о лице в файл')
        
    
class EditButton(QtGui.QPushButton):
    def __init__(self, index, parent=None):
        QtGui.QPushButton.__init__(self, parent)
        self.setMaximumSize(20, 20)
        self.index = index
        self.setIcon(QtGui.QIcon('images/pencil_0.png'))
    
    def getIndex(self):
        return self.index
    
    
class DeleteButton(QtGui.QPushButton):
    def __init__(self, index, parent=None):
        QtGui.QPushButton.__init__(self, parent)
        self.setMaximumSize(20, 20)
        self.index = index
        self.setIcon(QtGui.QIcon('images/cancel_0.png'))

    def getIndex(self):
        return self.index
    

class Request132Widget(QtGui.QDialog):
    '''диалог изменения данных требования о судимости
    добавить чекбокс подписывается лично исполнителем
    '''
    
    def __init__(self, request_data, parent=None):
        QtGui.QDialog.__init__(self, parent)
        self.request_data = request_data
        self.init()
        self.setData()
    
    def init(self):
        self.vbox_0 = QtGui.QVBoxLayout()
        self.gbox_0 = QtGui.QGroupBox('{}'.format(self.request_data[0]))
        self.vbox_1 = QtGui.QVBoxLayout()
        self.le_name_GIAC = QtGui.QLineEdit()
        self.le_name_GIAC.setMinimumWidth(300)
        self.hbox_0 = QtGui.QHBoxLayout()
        self.hbox_0.addWidget(QtGui.QLabel('Наименование подразделения ГИАЦ :'))
        self.hbox_0.addWidget(self.le_name_GIAC)
        self.vbox_1.addLayout(self.hbox_0)
        self.le_sity_GIAC = QtGui.QLineEdit()
        self.hbox_00 = QtGui.QHBoxLayout()
        self.hbox_00.addWidget(QtGui.QLabel('Город нахождения подразделения ГИАЦ :'))
        self.hbox_00.addWidget(self.le_sity_GIAC)
        self.vbox_1.addLayout(self.hbox_00)
        self.le_name_IC = QtGui.QLineEdit()
        self.hbox_1 = QtGui.QHBoxLayout()
        self.hbox_1.addWidget(QtGui.QLabel('Наименование подразделения ИЦ :'))
        self.hbox_1.addWidget(self.le_name_IC)
        self.vbox_1.addLayout(self.hbox_1)
        self.le_sity_IC = QtGui.QLineEdit()
        self.hbox_10 = QtGui.QHBoxLayout()
        self.hbox_10.addWidget(QtGui.QLabel('Город нахождения подразделения ИЦ :'))
        self.hbox_10.addWidget(self.le_sity_IC)
        self.vbox_1.addLayout(self.hbox_10)
        self.le_basis_verification = QtGui.QLineEdit()
        self.hbox_2 = QtGui.QHBoxLayout()
        self.hbox_2.addWidget(QtGui.QLabel('Основание проверки :'))
        self.hbox_2.addWidget(self.le_basis_verification)
        self.vbox_1.addLayout(self.hbox_2)
        self.le_necessary_info = QtGui.QLineEdit()
        self.hbox_3 = QtGui.QHBoxLayout()
        self.hbox_3.addWidget(QtGui.QLabel('Необходимая информация :'))
        self.hbox_3.addWidget(self.le_necessary_info)
        self.vbox_1.addLayout(self.hbox_3)
        self.gbox_0.setLayout(self.vbox_1)
        self.vbox_0.addWidget(self.gbox_0)
        self.cb_sign = QtGui.QCheckBox('Подписывается лично')
        self.vbox_0.addWidget(self.cb_sign)
        self.pb_box = QtGui.QDialogButtonBox()
        self.pb_box.addButton('Сохранить',
                        QtGui.QDialogButtonBox.AcceptRole)
        self.pb_box.addButton('Отменить',
                        QtGui.QDialogButtonBox.RejectRole)
        self.pb_box.accepted.connect(self.accept)
        self.pb_box.rejected.connect(self.reject)
        self.vbox_0.addWidget(self.pb_box)
        self.vbox_0.addStretch()
        self.setLayout(self.vbox_0)

    def setData(self):
        self.le_name_GIAC.setText(self.request_data[1])
        self.le_sity_GIAC.setText(self.request_data[2])
        self.le_name_IC.setText(self.request_data[3])
        self.le_sity_IC.setText(self.request_data[4])
        self.le_basis_verification.setText(self.request_data[5])
        self.le_necessary_info.setText(self.request_data[6])
        if self.request_data[7]: self.cb_sign.setCheckState(QtCore.Qt.Checked)


class RequestsListWidget(QtGui.QWidget):
    
    def __init__(self, request_list, parent=None):
        QtGui.QWidget.__init__(self, parent)
        self.request_list = request_list
        self.init()
        
    def init(self):
        self.btns_editRequest = []
        self.ch_boxs = []
        self.vbox_0 = QtGui.QVBoxLayout()
        self.setLayout(self.vbox_0)
        self.gbox_0 = QtGui.QGroupBox('Сформировать запросы в :')
        self.vbox_0.addWidget(self.gbox_0)
        self.vbox_1 = QtGui.QVBoxLayout()
        self.gbox_0.setLayout(self.vbox_1)
        self.btns_editRequest = []
        self.ch_boxs = []
        self.hbox_lay = []
        self.btns_delRequest = []
        self.btns_delRequest.append(DeleteButton(0))
        self.bt_run = QtGui.QPushButton('Начать создание запросов!')
        self.btn_add = QtGui.QPushButton('Добавить запрос')
        self.connect(self.btn_add, QtCore.SIGNAL('clicked()'),
                         self.addRequest)
        for i in range(0, len(self.request_list)):
            self.hbox_lay.append(QtGui.QHBoxLayout())
            self.ch_boxs.append(QtGui.QCheckBox(self.request_list[i][0]))
            self.hbox_lay[i].addWidget(self.ch_boxs[i])
            self.btns_editRequest.append(EditButton(i))
            self.connect(self.btns_editRequest[i],
                         QtCore.SIGNAL('clicked()'),
                         self.editRequest)
            self.hbox_lay[i].addWidget(self.btns_editRequest[i])
            if i != 0:
                self.btns_delRequest.append(DeleteButton(i))
                self.connect(self.btns_delRequest[i],
                         QtCore.SIGNAL('clicked()'),
                         self.deleteRequest)
                self.hbox_lay[i].addWidget(self.btns_delRequest[i])
            self.hbox_lay[i].addStretch()
            self.vbox_1.addLayout(self.hbox_lay[i])
        self.vbox_1.addStretch()
        self.vbox_0.addWidget(self.btn_add)
        self.vbox_0.addWidget(self.bt_run)
    
    def addRequest(self):
        new_request_window = RequestTabWidget(['', '', '', '', '', '', '', '', '',
                                               '', '', '', '', '', '', '', ''])
        new_request_window.setWindowTitle('Введите данные нового запроса')
        result = new_request_window.exec_()
        if result:
            self.request_list.append([new_request_window.le_short_name.text(),
                                      new_request_window.le_appointment.text(),
                                      new_request_window.le_name_organization.text(),
                                      new_request_window.le_rang.text(),
                                      new_request_window.le_surname.text(),
                                      new_request_window.le_initial.text(),
                                      new_request_window.le_index.text(),
                                      new_request_window.le_country.text(),
                                      new_request_window.le_province.text(),
                                      new_request_window.le_district.text(),
                                      new_request_window.le_town.text(),
                                      new_request_window.le_street.text(),
                                      new_request_window.le_house.text(),
                                      new_request_window.le_housein.text(),
                                      new_request_window.le_corps.text(),
                                      new_request_window.le_apartment.text(),
                                      new_request_window.le_need_info.text()])
        else:
            return
        self.hbox_lay.append(QtGui.QHBoxLayout())
        self.ch_boxs.append(QtGui.QCheckBox(self.request_list[-1][0]))
        self.hbox_lay[-1].addWidget(self.ch_boxs[-1])
        self.btns_editRequest.append(EditButton(len(self.request_list) - 1))
        self.connect(self.btns_editRequest[-1],
                         QtCore.SIGNAL('clicked()'),
                         self.editRequest)
        self.hbox_lay[-1].addWidget(self.btns_editRequest[-1])
        self.btns_delRequest.append(DeleteButton(len(self.request_list) - 1))
        self.connect(self.btns_delRequest[-1],
                        QtCore.SIGNAL('clicked()'),
                         self.deleteRequest)
        self.hbox_lay[-1].addWidget(self.btns_delRequest[-1])
        self.hbox_lay[-1].addStretch()
        self.vbox_1.insertLayout(len(self.request_list) - 1, self.hbox_lay[-1])
        self.saveRequest()
    
    def deleteRequest(self):
        index = self.sender().getIndex()
        if index == 0:
            QtGui.QMessageBox.information(self, "Внимание",
                    'Удалять требование о судимости нельзя!')
            return
        confirm_window = QtGui.QMessageBox(QtGui.QMessageBox.Question,
                                           'Внимание!', 'Вы точно хотите удалить запрос и все его данные?',
                                           buttons=QtGui.QMessageBox.Ok | QtGui.QMessageBox.Cancel,
                                           parent=self)
        result = confirm_window.exec_()
        if result:
            for i in range(index + 1, len(self.request_list)):
                self.btns_editRequest[i].index = i - 1
                self.btns_delRequest[i].index = i - 1
            self.ch_boxs[index].hide()
            del self.ch_boxs[index]
            del self.request_list[index]
            self.disconnect(self.btns_editRequest[index],
                         QtCore.SIGNAL('clicked()'),
                         self.editRequest)
            self.btns_editRequest[index].hide()
            del self.btns_editRequest[index]
            self.disconnect(self.btns_delRequest[index],
                        QtCore.SIGNAL('clicked()'),
                         self.deleteRequest)
            self.btns_delRequest[index].hide()
            del self.btns_delRequest[index]
            del self.hbox_lay[index]
            self.vbox_1.update()
        self.saveRequest()
    
    def editRequest(self):
        index = self.sender().getIndex()
        old_data_request = self.request_list.pop(index)
        if index == 0:
            req_edit_windows = Request132Widget(old_data_request)
            req_edit_windows.setWindowTitle('Изменить данные требования о судимости')
            result = req_edit_windows.exec_()
            if result:
                self.request_list.insert(index, ['Требование о судимости',
                                                 req_edit_windows.le_name_GIAC.text(),
                                                 req_edit_windows.le_sity_GIAC.text(),
                                                 req_edit_windows.le_name_IC.text(),
                                                 req_edit_windows.le_sity_IC.text(),
                                                 req_edit_windows.le_basis_verification.text(),
                                                 req_edit_windows.le_necessary_info.text(),
                                                 req_edit_windows.cb_sign.isChecked()])
            else:
                self.request_list.insert(index, old_data_request)
        else:
            req_edit_windows = RequestTabWidget(old_data_request)
            result = req_edit_windows.exec_()
            if result:
                self.request_list.insert(index, [req_edit_windows.le_short_name.text(),
                                             req_edit_windows.le_appointment.text(),
                                            req_edit_windows.le_name_organization.text(),
                                            req_edit_windows.le_rang.text(),
                                            req_edit_windows.le_surname.text(),
                                            req_edit_windows.le_initial.text(),
                                            req_edit_windows.le_index.text(),
                                            req_edit_windows.le_country.text(),
                                            req_edit_windows.le_province.text(),
                                            req_edit_windows.le_district.text(),
                                            req_edit_windows.le_town.text(),
                                            req_edit_windows.le_street.text(),
                                            req_edit_windows.le_house.text(),
                                            req_edit_windows.le_housein.text(),
                                            req_edit_windows.le_corps.text(),
                                            req_edit_windows.le_apartment.text(), 
                                            req_edit_windows.le_need_info.text()])
                self.ch_boxs[index].setText(self.request_list[index][0])
            else:
                self.request_list.insert(index, old_data_request)
        self.saveRequest()

    def saveRequest(self):
        try:
            requests_file = open('request', 'wb')
            pickle.dump(self.request_list, requests_file)
            requests_file.close()
        except IOError:
                QtGui.QMessageBox.information(self, "Внимание",
                    'Не удается сохранить запросы в файл')


class MainWindow (QtGui.QMainWindow):
    '''основное окно программы'''
    def __init__(self):
        QtGui.QMainWindow.__init__(self)
        try:
            with open('setting', 'rb') as setting_file:
                self.setting = pickle.load(setting_file)
        except IOError:
            self.ferstStart()
        try:
            with open('request', 'rb') as requests_file:
                self.requests_list = pickle.load(requests_file)
        except IOError:
            QtGui.QMessageBox.information(self, "Внимание",
                    'Отсутствует файл с данными запросов.\nПриняты значения по умолчанию.')
            self.requests_list = [['Требование о судимости', '', '', '', '', 'привлечение к уголовной ответственности',
                                   'о судимости', False],
                                  ['Психолог', 'Главному врачу', 'КГБУЗ ""', '', '', '', '', '', '', '',
                                   '', '', '', '', '', '', ' справку в которой отразить состоит (состоял) ли на учете, (проходил ли, проходит ли лечение)'],
                                  ['Нарколог', 'Главному врачу', 'КГБУЗ ""', '', '', '', '', '', '', '',
                                   '', '', '', '', '', '', ' справку в которой отразить состоит (состоял) ли на учете, (проходил ли, проходит ли лечение)'],
                                  ['Характеристика', 'Начальнику', '', '', '', '', '', '', '', '',
                                   '', '', '', '', '', '', ' социально-бытовую характеристику участкового уполномоченного по месту жительства '], ]
        try:
            with open('person', 'rb') as person_file:
                self.person = pickle.load(person_file)
        except IOError:
            self.person = None
        QtGui.QMessageBox(QtGui.QMessageBox.Information,
                        'Одну Минуточку!', info,
                        buttons = QtGui.QMessageBox.Ok, parent = self)
        self.info.exec_()
        self.init()
        self.readSettings()
            
    def ferstStart(self):
        #print ('Первый запуск')
        m_window_setting = QtGui.QDialog()
        vbox = QtGui.QVBoxLayout()
        m_window_setting.setLayout(vbox)
        m_window_setting.setWindowModality(QtCore.Qt.ApplicationModal)
        central_widget = RequestOverheadDataWidget()
        vbox.addWidget(central_widget)
        db_box = QtGui.QDialogButtonBox()
        db_box.addButton('Сохранить настройки',
                        QtGui.QDialogButtonBox.AcceptRole)
        db_box.addButton('Отменить',
                        QtGui.QDialogButtonBox.RejectRole)
        vbox.addWidget(db_box)
        db_box.accepted.connect(m_window_setting.accept)
        db_box.rejected.connect(m_window_setting.reject)
        result = m_window_setting.exec_()
        if result == 1:
            self.setting = [[central_widget.le_short_title_superior.text(),
                          central_widget.le_full_title.text(),
                          central_widget.le_short_title.text(),
                          central_widget.le_full_unit_name.text(),
                          central_widget.le_short_unit_name.text(),
                          central_widget.le_index_0.text(),
                          central_widget.le_country_0.text(),
                          central_widget.le_province_0.text(),
                          central_widget.le_district_0.text(),
                          central_widget.le_town_0.text(),
                          central_widget.le_street_0.text(),
                          central_widget.le_house_0.text(),
                          central_widget.le_housein_0.text(),
                          central_widget.le_corps_0.text(),
                          central_widget.le_apartment_0.text(),
                          central_widget.le_phone_0.text(),
                          central_widget.le_fax_0.text(),
                          central_widget.le_email_0.text(),
                          central_widget.le_short_title_dat.text(),
                          central_widget.le_short_title_rod.text()],
                        [central_widget.le_position_head.text(),
                          central_widget.le_rank_head.text(),
                          central_widget.le_surname_head.text(),
                          central_widget.le_name_head.text(),
                          central_widget.le_patronymic_head.text()],
                        [central_widget.le_position_executive.text(),
                          central_widget.le_rank_executive.text(),
                          central_widget.le_surname_executive.text(),
                          central_widget.le_name_executive.text(),
                          central_widget.le_patronymic_executive.text(),
                          central_widget.le_phone_executive.text(),
                          central_widget.le_fax_executive.text(),
                          central_widget.le_email_executive.text()]] 
        else:
            self.close()
    
    def init(self):
        
        self.setWindowTitle('Автоматизатор запросов характеризующего')
        self.createStatusBar()
        
        '''# определяем область механизм прокрутки (QScrollArea)
        # self.scrollArea = QtGui.QScrollArea()
        # self.scrollArea.setWidgetResizable(True) #разрешаем проктурку
        # добавляем на область виджет, с ранее добавленным на него слоем слоем
        # self.scrollArea.setWidget(self.scrollWidget)
        # создаём главный вертикальный слой
        # self.mainLayout = QtGui.QVBoxLayout()
 
        # добавляем элементы на главный слой
        # self.mainLayout.addWidget(self.addButton) # добавляем основную кнопку
        # self.mainLayout.addWidget(self.scrollArea) # добавляем область прокрутки
 
        # определяем "центральный виджет"
        # self.centralWidget = QtGui.QWidget()
        # self.centralWidget.setLayout(self.mainLayout)'''
        
        self.main_tab_windget = QtGui.QTabWidget(self) 
        if self.setting:
            self.overhead_data = RequestOverheadDataWidget(self.setting)
            self.main_tab_windget.addTab(self.overhead_data, 'Служебные данные')
            self.overhead_data.saveData()

            self.request_person = RequestPersonWidget(self.person)
            self.main_tab_windget.addTab(self.request_person, 'Данные о лице')
            self.request_person.saveDataP()
            
            self.request_widget = RequestsListWidget(self.requests_list)
            self.main_tab_windget.addTab(self.request_widget, 'Запросы')
            # подключение кнопки магии
            self.request_widget.bt_run.clicked.connect(self.magic)
            self.setCentralWidget(self.main_tab_windget)
        else:
            return
        self.main_tab_windget.setCurrentIndex(1)
        self.statusBar().showMessage("Готов к работе")
        
    def magic(self):
        verif = True
        for ch_box in self.request_widget.ch_boxs:
            if ch_box.isChecked():
                verif = False
        if verif:
            QtGui.QMessageBox.information(self, "Внимание",
                    'Не выбрано запросов! Куда шлем? Определись!')
            return
        if self.overhead_data.verification():
            self.main_tab_windget.setCurrentIndex(0)
            return
        elif self.request_person.verification():
            self.main_tab_windget.setCurrentIndex(1)
            return
        requ = 0
        if self.request_widget.ch_boxs[0].checkState() == QtCore.Qt.Checked:
            self.generate132Document()
            requ += 1
        for i in range(1, len(self.request_widget.request_list)):
            if self.request_widget.ch_boxs[i].checkState() == QtCore.Qt.Checked:
                self.generateRequestDocument(self.request_widget.request_list[i])
                requ += 1
        if requ > 0:
            QtGui.QMessageBox.information(self, "Поздравляю",
                    'Успешно создано {} запросов. Смотри папку "Запросы" в папке с программой'.format(requ))
        for ch_box in self.request_widget.ch_boxs:
            ch_box.setCheckState(QtCore.Qt.Unchecked)

    def generate132Document(self):
        #print ('    требование о судимости')
        patern_file = 'patern/132.docx'
        new_doc = 'запросы/{} уд {} 132.docx'.format(self.request_person.le_surname.text(),
                                                     self.request_person.le_number_case.text())
        replace_text = dict(
            NAME_IC = self.requests_list[0][3],
            NAME_GIAC = self.requests_list[0][1],
            CITY_GIAC = self.requests_list[0][2],
            CITY_IC = self.requests_list[0][4],
            SURNAME = '{}'.format(self.request_person.le_surname.text()),
            NAME = '{}'.format(self.request_person.le_name.text()),
            PATRONYMIC = ' {}'.format(self.request_person.le_patronymic.text()),
            DATE_OB = self.request_person.de_date_ob.date().toPyDate().strftime("%d.%m.%Y"),
            PLACE_OB = '{}'.format(self.request_person.le_place_ob.text()),
            ADRESS = '{province}{district}{town} ул. {street}, д. {house}{housein}{corps}{apartment}'.format(
                province = '' if self.request_person.le_province_2.text() == '' else ' {},'.format(self.request_person.le_province_2.text()),
                district = '' if self.request_person.le_district_2.text() == '' else ' {},'.format(self.request_person.le_district_2.text()),
                town = '' if self.request_person.le_town_2.text() == '' else ' {},'.format(self.request_person.le_town_2.text()),
                street = self.request_person.le_street_2.text(),
                house = self.request_person.le_house_2.text(),
                housein = '' if self.request_person.le_housein_2.text() == '' else ', стр. {}'.format(self.request_person.le_housein_2.text()),
                corps = '' if self.request_person.le_corps_2.text() == '' else ', корп. {}'.format(self.request_person.le_corps_2.text()),
                apartment = '' if self.request_person.le_apartment_2.text() == '' else ', кв. {}'.format(self.request_person.le_apartment_2.text())) if self.request_person.cb_bloc_ap.isChecked() else ', зарегистрированный:{province}{district}{town} ул. {street}, д. {house}{housein}{corps}{apartment}'.format(
                province = '' if self.request_person.le_province_3.text() == '' else ' {},'.format(self.request_person.le_province_3.text()),
                district = '' if self.request_person.le_district_3.text() == '' else ' {},'.format(self.request_person.le_district_3.text()),
                town = '' if self.request_person.le_town_3.text() == '' else ' {},'.format(self.request_person.le_town_3.text()),
                street = self.request_person.le_street_3.text(),
                house = self.request_person.le_house_3.text(),
                housein = '' if self.request_person.le_housein_3.text() == '' else ', стр. {}'.format(self.request_person.le_housein_3.text()),
                corps = '' if self.request_person.le_corps_3.text() == '' else ', корп. {}'.format(self.request_person.le_corps_3.text()),
                apartment = '' if self.request_person.le_apartment_3.text() == '' else ', кв. {}'.format(self.request_person.le_apartment_3.text())),
            BASIS_VERIFICATION = self.requests_list[0][5],
            NEED_INFO = self.requests_list[0][6],
            SIGN = '''{POSITION}
{SHORT_TITLE_ROD}
{RANK}              {INITIAL}{SURNAME}'''.format(
    POSITION = self.overhead_data.le_position_executive.text(),
    SHORT_TITLE_ROD = self.overhead_data.le_short_title_rod.text(),
    RANK = self.overhead_data.le_rank_executive.text(),
    SURNAME = ' {}'.format(self.overhead_data.le_surname_executive.text()),
    INITIAL = '{}.{}. '.format(
                        self.overhead_data.le_name_executive.text()[0],
                        self.overhead_data.le_patronymic_executive.text()[0]),
) if self.requests_list[0][7] else '''{POSITION}
{SHORT_TITLE_ROD}
{RANK}              {INITIAL}{SURNAME}'''.format(
    POSITION = self.overhead_data.le_position_head.text(),
    SHORT_TITLE_ROD = self.overhead_data.le_short_title_rod.text(),
    RANK = self.overhead_data.le_rank_head.text(),
    INITIAL = '{}.{}. '.format(
        self.overhead_data.le_name_head.text()[0],
        self.overhead_data.le_patronymic_head.text()[0]
),
    SURNAME = '{}'.format(self.overhead_data.le_surname_head.text()),
),
            SURNAME_EXECUTIVE = '{}'.format(self.overhead_data.le_surname_executive.text()),
            INITIAL_EXECUTIVE = ' {}.{}.'.format(
                        self.overhead_data.le_name_executive.text()[0],
                        self.overhead_data.le_patronymic_executive.text()[0]),
            PHONE_EXECUTIVE = '' if self.overhead_data.le_phone_executive.text() == '' else 'тел. {}'.format(self.overhead_data.le_phone_executive.text()),
            SHORT_TITLE = '{}'.format(self.overhead_data.le_short_title.text()),
            ADRESS_00 = 'ул. {street}, д. {house}{housein}{corps}{apartment}'.format(
                street = self.overhead_data.le_street_0.text(),
                house = self.overhead_data.le_house_0.text(),
                housein = '' if self.overhead_data.le_housein_0.text() == '' else ', стр. {},'.format(self.overhead_data.le_housein_0.text()),
                apartment = '' if self.overhead_data.le_apartment_0.text() == '' else ', кв. {},'.format(self.overhead_data.le_apartment_0.text()),
                corps = '' if self.overhead_data.le_corps_0.text() == '' else ', корп. {},'.format(self.overhead_data.le_corps_0.text())),
            ADRESS_0 = '{town}{district}{province}{index}'.format(
                index = '' if self.overhead_data.le_index_0.text() == '' else ' {}'.format(self.overhead_data.le_index_0.text()),
                province = '' if self.overhead_data.le_province_0.text() == '' else ' {},'.format(self.overhead_data.le_province_0.text()),
                district = '' if self.overhead_data.le_district_0.text() == '' else ' {},'.format(self.overhead_data.le_district_0.text()),
                town = '' if self.overhead_data.le_town_0.text() == '' else '{},'.format(self.overhead_data.le_town_0.text())),
            DATE_PRINT = datetime.date.today().strftime("%d.%m.%Y "),)
        docx(patern_file, replace_text, new_doc)
        
    def generateRequestDocument(self, request):
        #print ('запрос в {}'.format(request[0]))
        patern_file = 'patern/request_exec 2.docx'
        new_doc = 'запросы/{} уд {} запрос в {}.docx'.format(self.request_person.le_surname.text(),
                                                                self.request_person.le_number_case.text(),
                                                                request[0])
        replace_text = dict(
            SHORT_TITLE_SUPERIOR = self.overhead_data.le_short_title_superior.text(),
            FULL_TITLE = self.overhead_data.le_full_title.text(),
            SHORT_TITLE = '({})'.format(self.overhead_data.le_short_title.text()),
            FULL_UNIT_NAME = self.overhead_data.le_full_unit_name.text(),
            ADRESS_00 = 'ул. {street}, д. {house}{housein}{corps}{apartment}'.format(
                street = self.overhead_data.le_street_0.text(),
                house = self.overhead_data.le_house_0.text(),
                housein = '' if self.overhead_data.le_housein_0.text() == '' else ', стр. {},'.format(self.overhead_data.le_housein_0.text()),
                apartment = '' if self.overhead_data.le_apartment_0.text() == '' else ', кв. {},'.format(self.overhead_data.le_apartment_0.text()),
                corps = '' if self.overhead_data.le_corps_0.text() == '' else ', корп. {},'.format(self.overhead_data.le_corps_0.text())),
            ADRESS_0 = '{town}{district}{province}{index}'.format(
                index = '' if self.overhead_data.le_index_0.text() == '' else ' {}'.format(self.overhead_data.le_index_0.text()),
                province = '' if self.overhead_data.le_province_0.text() == '' else ' {},'.format(self.overhead_data.le_province_0.text()),
                district = '' if self.overhead_data.le_district_0.text() == '' else ' {},'.format(self.overhead_data.le_district_0.text()),
                town = '' if self.overhead_data.le_town_0.text() == '' else '{},'.format(self.overhead_data.le_town_0.text())),
            PHONEFAX_0 = '{}{}'.format('' if self.overhead_data.le_phone_0.text() == '' else 'тел:{}'.format(self.overhead_data.le_fax_0.text()),
                '' if self.overhead_data.le_fax_0.text() == '' else ' факс:{}'.format(self.overhead_data.le_fax_0.text())),
            EMAIL_0 = '' if self.overhead_data.le_email_0.text() == '' else 'email:{}'.format(self.overhead_data.le_email_0.text()),
            APPOINTMENT = '{}'.format(request[1]),
            NAME_ORGANIZATION = '{}'.format(request[2]),
            RANG = '{}'.format(request[3]),
            INITIAL_0 = '{}'.format(request[5]),
            SURNAME_0 = ' {}'.format(request[4]),
            ADRESS_1 = '''ул. {street}, д. {house}{housein}{corps}{apartment}
{town}{district}{province}{index}'''.format(
                index = '' if request[6] == '' else ' {}'.format(request[6]),
                province = '' if request[8] == '' else ' {},'.format(request[8]),
                district = '' if request[9] == '' else ' {},'.format(request[9]),
                town = '' if request[10] == '' else '{},'.format(request[10]),
                street = request[11],
                house = request[12],
                housein = '' if request[13] == '' else ', стр. {},'.format(request[13]),
                corps = '' if request[14] == '' else ', корп. {},'.format(request[14]),
                apartment = '' if request[15] == '' else ', кв. {},'.format(request[15])),
            DATE_OUT_REG = datetime.date.today().strftime("%d.%m.%Y"),
            NUMBER_OUT_REG = '  /',
            NUMBER_IN_REG = '',
            DATE_IN_REG = '',
            NUMBER_CASE = '{} '.format(self.request_person.le_number_case.text()),
            DATE_CASE = self.request_person.de_date_case.date().toPyDate().strftime("%d.%m.%Y "),
            ARTICLE = '{}{}{}'.format(
                '' if self.request_person.le_point.text() == '' else 'п. {} '.format(self.request_person.le_point.text()),
                '' if self.request_person.le_part.text() == '' else 'ч. {} '.format(self.request_person.le_part.text()),
                'ст. {} '.format(self.request_person.le_article.text())),
            SHORT_UNIT_NAME = self.overhead_data.le_short_unit_name.text(),
            SHORT_UNIT_NAME_DAT = self.overhead_data.le_short_title_dat.text(),
            NEED_INFO = request[16],
            SURNAME = ' {}'.format(self.request_person.le_surname.text()),
            NAME = ' {}'.format(self.request_person.le_name.text()),
            PATRONYMIC = ' {}'.format(self.request_person.le_patronymic.text()),
            DATE_OB = self.request_person.de_date_ob.date().toPyDate().strftime(" %d.%m.%Y "),
            PLACE_OB = ' {}'.format(self.request_person.le_place_ob.text()),
            ADRESS_2 = ' проживающий:{province}{district}{town} ул. {street}, д. {house}{housein}{corps}{apartment}'.format(
                province = '' if self.request_person.le_province_2.text() == '' else ' {},'.format(self.request_person.le_province_2.text()),
                district = '' if self.request_person.le_district_2.text() == '' else ' {},'.format(self.request_person.le_district_2.text()),
                town = '' if self.request_person.le_town_2.text() == '' else ' {},'.format(self.request_person.le_town_2.text()),
                street = self.request_person.le_street_2.text(),
                house = self.request_person.le_house_2.text(),
                housein = '' if self.request_person.le_housein_2.text() == '' else ', стр. {}'.format(self.request_person.le_housein_2.text()),
                corps = '' if self.request_person.le_corps_2.text() == '' else ', корп. {}'.format(self.request_person.le_corps_2.text()),
                apartment = '' if self.request_person.le_apartment_2.text() == '' else ', кв. {}'.format(self.request_person.le_apartment_2.text())),
            ADRESS_3 = '' if self.request_person.cb_bloc_ap.isChecked() else ', зарегистрированный:{province}{district}{town} ул. {street}, д. {house}{housein}{corps}{apartment}'.format(
                province = '' if self.request_person.le_province_3.text() == '' else ' {},'.format(self.request_person.le_province_3.text()),
                district = '' if self.request_person.le_district_3.text() == '' else ' {},'.format(self.request_person.le_district_3.text()),
                town = '' if self.request_person.le_town_3.text() == '' else ' {},'.format(self.request_person.le_town_3.text()),
                street = self.request_person.le_street_3.text(),
                house = self.request_person.le_house_3.text(),
                housein = '' if self.request_person.le_housein_3.text() == '' else ', стр. {}'.format(self.request_person.le_housein_3.text()),
                corps = '' if self.request_person.le_corps_3.text() == '' else ', корп. {}'.format(self.request_person.le_corps_3.text()),
                apartment = '' if self.request_person.le_apartment_3.text() == '' else ', кв. {}'.format(self.request_person.le_apartment_3.text())),
            DATE_EXECUTE = (datetime.date.today() + datetime.timedelta(10)).strftime("%d.%m.%Y"),
            FAX_EXECUTIVE = ', предварительно направив по факсу {}'.format(self.overhead_data.le_fax_executive.text()) if self.overhead_data.le_fax_executive.text() != '' else '',
            EMAIL_EXECUTIVE = ', или на электронную почту {}'.format(self.overhead_data.le_email_executive.text()) if self.overhead_data.le_email_executive.text() != '' else '', 
            POSITION_EXECUTIVE = self.overhead_data.le_position_executive.text(),
            RANK_EXECUTIVE = self.overhead_data.le_rank_executive.text(),
            INITIAL_EXECUTIVE = '{}.{}. '.format(
                        self.overhead_data.le_name_executive.text()[0],
                        self.overhead_data.le_patronymic_executive.text()[0]),
            SURNAME_EXECUTIVE = '{}'.format(self.overhead_data.le_surname_executive.text()),
            PHONE_EXECUTIVE = '' if self.overhead_data.le_phone_executive.text() == '' else 'тел. {}'.format(self.overhead_data.le_phone_executive.text()))
        docx(patern_file, replace_text, new_doc)
        
    def createStatusBar(self):
        self.statusBar().showMessage("Идет загрузка данных")

    def readSettings(self):
        settings = QtCore.QSettings('Следователь(с)', 'АРМ-следователь')
        pos = settings.value('pos', QtCore.QPoint(0, 0))
        size = settings.value('size', QtCore.QSize(800, 1200))
        self.move(pos)
        self.resize(size)

    def writeSettings(self):
        settings = QtCore.QSettings('Следователь(с)', 'АРМ-следователь')
        settings.setValue('pos', self.pos())
        settings.setValue('size', self.size())
    
    def closeEvent(self, event):
        self.writeSettings()
        event.accept()

    
if __name__ == '__main__':
    import sys
    app = QtGui.QApplication(sys.argv)
    winWig = MainWindow()
    print ('не закрывайте это окно во время работы программы')
    winWig.show()
    print ('\nтеперь это окно можно закрыть.\nспасибо что воспользовались моей программой.')
    sys.exit(app.exec_())
