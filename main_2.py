import os
import shutil
import pandas as pd
from tableau_api_lib import TableauServerConnection
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from urllib import parse
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

# Creation of Extension class
class TableauExtension:

    # Variables for loading bar
    count_views = 0
    counter = 0
    status_percent = 0

    # init - automatically created while calling this class
    def __init__(self):
        self.status = self.status_percent
        self.connection = self.tableau_login()

    # gives back status for loading bar
    def check_status(self):
        return self.status_percent

    # current config - can be changed in the future
    @staticmethod
    def tableau_login():
        tableau_server_config = {
            'tableau_prod': {
                'server': 'http://tableau.demo.sgc.corp/',
                'api_version': '3.19',
                'personal_access_token_name': 'opv-python-token',
                'personal_access_token_secret': '+7NLXnzxR9S+UxhmvXfn8g==:e6jrJQVUDjOf8h2638VjJeuNgQ0bZS3V',
                'site_name': 'bnt-extension-poc',
                'site_url': 'bnt-extension-poc'
            }
        }
        conn = TableauServerConnection(tableau_server_config)
        conn.sign_in()
        return conn

    # giving back workbook id
    def get_workbook_id(self):
        WORKBOOK_NAME = '2023_OPV_TechStack_V1_Base'
        workbooks = self.connection.query_workbooks_for_site().json()['workbooks']['workbook']
        for workbook in workbooks:
            if workbook['name'] == WORKBOOK_NAME:
                return workbook['id']
        return workbooks.pop()['id']

    # using workbook id and gives back a list of all views - filtered to OPV
    def query_viewnames_for_workbook(self):
        conn = self.connection
        view_list = []
        workbook_id = self.get_workbook_id()
        response = conn.query_views_for_workbook(workbook_id)
        views_list_dict = response.json()['views']['view']
        for view in views_list_dict:
            view_list.append((view['id'],view['name']))
        # take only the necessary dashboards
        df_complete = pd.DataFrame(view_list,columns = ['view_id','view_name'])
        #Filterschritt
        df = df_complete[df_complete['view_name'].str.contains("OPV")]
        return df


    # function for saving the pdf file
    def save_as_pdf (self,pdf_merger):
        pdfPath = "APQR.pdf"
        
        pdfOutputFile = open(pdfPath, 'wb')
        pdf_merger.write(pdfOutputFile)
        pdfOutputFile.close()

    # creating the seperate pdf files and merge them
    def create_pdf (self):
        FILE_PREFIX = 'bnt_'
        views = self.query_viewnames_for_workbook()
        pdf_list = []

        # csv generation for pdf creation - extension must know some values
        file_content = self.create_csv().content
        file_path = 'APQR.csv'
        with open(file_path, 'wb') as file:
                file.write(file_content)
                file.close()

        # Zusammenfassung aller notwendigen Spalten
        column_names = ['Material Number','Sample Stage','Result Name']
        # csv zu Dataframe
        df = pd.read_csv('APQR.csv',delimiter=';')[column_names]
        # Sortierung nach den jeweiligen Spalten
        df_sorted = df.sort_values(by=["Material Number", "Sample Stage", "Result Name"])
        # Distinct, um nur relevante Infos zu behalten
        df_unique = df_sorted.drop_duplicates(subset=["Material Number", "Sample Stage", "Result Name"])

        #TEST
        df_filtered = df_unique[(df_unique["Material Number"] == 20000019) & (df_unique["Sample Stage"] == "DP_B-IPC6")]
            

        try:
            shutil.rmtree('./temp/')
        except:
            print('Path not available')
        os.mkdir('./temp')

        self.count_views = len(views.index)
        print(self.count_views)

        conn = self.connection
        #print(type(view_string))

            
        # For schleifen
        #for material_number in df_unique["Material Number"].unique():
        for material_number in df_filtered["Material Number"].unique():
            material_group = df_filtered[df_filtered["Material Number"] == material_number]
            #material_group = df_unique[df_unique["Material Number"] == material_number]
            # Variable für die erste Seite
            first_page_done = 0
                
            for sample_stage in material_group["Sample Stage"].unique():
                stage_group = material_group[material_group["Sample Stage"] == sample_stage]

                first_page_done = 0
                    
                for result_name in stage_group["Result Name"].unique():                        
                    for ind in views.index:
                        if first_page_done == 0 and ind == 0:
                            #Filter Variablen, welche später noch manipuliert werden müssen
                            parameter_filter_name = parse.quote('Result Name')
                            parameter_filter_value = parse.quote('')

                            sample_stage_filter_name = parse.quote('Sample Stage')
                            sample_stage_filter_value = parse.quote(str(sample_stage))

                            material_number_name = parse.quote('Material Number')
                            material_number_filter_value = parse.quote(str(material_number).split('.', 1)[0])
                            # Müssen immer wieder aktualisiert werden und durchiteriert werden
                            # Können später nach Sample Stage filtern um Größe zu verkleinern
                            pdf_params = {
                                'type': 'type=A4',
                                'orientation': 'orientation=Landscape',
                                'parameter_filter': f'vf_{parameter_filter_name}={parameter_filter_value}',
                                'sample_stage_filter': f'vf_{sample_stage_filter_name}={sample_stage_filter_value}',
                                'material_number_filter': f'vf_{material_number_name}={material_number_filter_value}'
                            }
                            #print(views['view_id'][ind])
                            self.counter = self.counter + 1
                            
                            view_string = views['view_id'][ind]

                            pdf = conn.query_view_pdf(view_id=view_string, parameter_dict=pdf_params)
                            with open('./temp/'+f'{FILE_PREFIX}{self.counter}.pdf', 'wb') as pdf_file:
                                pdf_file.write(pdf.content)
                                pdf_list.append(pdf_file.name)
                                pdf_file.close()
                                
                            self.status_percent = round(self.counter/self.count_views*100,2)
                            print(self.status_percent)
                            first_page_done = 1
                        elif first_page_done == 1 and ind == 0:
                            pass
                        else:
                            # Filter Variablen, welche später noch manipuliert werden müssen
                            parameter_filter_name = parse.quote('Result Name')
                            parameter_filter_value = parse.quote(str(result_name))

                            sample_stage_filter_name = parse.quote('Sample Stage')
                            sample_stage_filter_value = parse.quote(str(sample_stage))

                            material_number_name = parse.quote('Material Number')
                            material_number_filter_value = parse.quote(str(material_number).split('.', 1)[0])
                            # Müssen immer wieder aktualisiert werden und durchiteriert werden
                            # Können später nach Sample Stage filtern um Größe zu verkleinern
                            pdf_params = {
                                'type': 'type=A4',
                                'orientation': 'orientation=Landscape',
                                'parameter_filter': f'vf_{parameter_filter_name}={parameter_filter_value}',
                                'sample_stage_filter': f'vf_{sample_stage_filter_name}={sample_stage_filter_value}',
                                'material_number_filter': f'vf_{material_number_name}={material_number_filter_value}'
                            }
                            self.counter = self.counter + 1
                            
                            view_string = views['view_id'][ind]

                            pdf = conn.query_view_pdf(view_id=view_string, parameter_dict=pdf_params)
                            with open('./temp/'+f'{FILE_PREFIX}{self.counter}.pdf', 'wb') as pdf_file:
                                pdf_file.write(pdf.content)
                                pdf_list.append(pdf_file.name)
                                pdf_file.close()
                                
                            self.status_percent = round(self.counter/self.count_views*100,2)
                            print(self.status_percent)


        merger = PdfMerger()
        for pdf in pdf_list:
            merger.append(pdf)
        self.save_as_pdf(merger)
        merger.close()

        # adding page numbers, removing old file and renaming
        self.add_page_numgers("APQR.pdf", "APQR_numbered.pdf")
        os.remove("APQR.pdf")
        os.rename("APQR_numbered.pdf", "APQR.pdf")

        #clean up
        conn.sign_out()
        shutil.rmtree('./temp/')
        self.count_views = 0
        self.counter = 0
        self.status_percent = 0

        return 'PDF created'
    
    # using the workbook id to get the helping view
    def create_csv (self):
        conn = self.connection
        view_list = []
        workbook_id = self.get_workbook_id()
        response = conn.query_views_for_workbook(workbook_id)
        views_list_dict = response.json()['views']['view']
        for view in views_list_dict:
            view_list.append((view['id'],view['name']))
        # take only the necessary dashboards
        df_complete = pd.DataFrame(view_list,columns = ['view_id','view_name'])
        #Filterschritt
        df = df_complete[df_complete['view_name'].str.contains("Helper")]
        view_id = str(df.head(1)['view_id'][df.index.values.astype(int)[0]])
        print(view_id)
        response = conn.query_view_data(view_id)
        return response
    
    # function for saving the csv file
    def save_as_csv (self,response):
        csvPath = "APQR.csv"
        if csvPath: #If the user didn't close the dialog window
            with open(csvPath, 'wb') as csv_file:
                csv_file.write(response.content)
                csv_file.close()

    # Functions for pagenumbers
    def create_page_pdf(self, num, tmp):
        c = canvas.Canvas(tmp)
        for i in range(1, num + 1):
            # 210 x 297 -> A4
            # First Value Point from left, second point from bottom
            c.drawString((297 // 2) * mm, (4) * mm, str(i))
            c.showPage()
        c.save()

    def add_page_numgers(self, pdf_path, newpath):
        """
        Add page numbers to a pdf, save the result as a new pdf
        @param pdf_path: path to pdf
        """
        tmp = "__tmp.pdf"

        writer = PdfWriter()
        with open(pdf_path, "rb") as f:
            reader = PdfReader(f)
            n = len(reader.pages)

            # create new PDF with page numbers
            self.create_page_pdf(n, tmp)

            with open(tmp, "rb") as ftmp:
                number_pdf = PdfReader(ftmp)
                # iterarte pages
                for p in range(n):
                    page = reader.pages[p]
                    number_layer = number_pdf.pages[p]
                    # merge number page with actual page
                    page.merge_page(number_layer)
                    writer.add_page(page)

                # write result
                if len(writer.pages) > 0:
                    with open(newpath, "wb") as f:
                        writer.write(f)
            os.remove(tmp)

