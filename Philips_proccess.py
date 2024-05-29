import pandas as pd
import re
import numpy as np
import PyPDF2
import datetime

class make_excelp ():

    def extract_text_from_pdf(self,pdf_file_path):
        # Initialize an empty list to store the extracted data
        self.text = []
        with open(pdf_file_path, "rb") as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                self.text.extend(page.extract_text().split())
        return self.text

    def newdata(self,ex_data):
        self.name = []
        i = 0
        while not re.search(r'Series', str(ex_data[i])):
            self.name.append(ex_data[i])
            i += 1
        else:
            self.name.append(ex_data[i])
        print(self.name)
        self.birth = None

        self.id = []
        self.manuf = []
        self.cont = []
        self.obs = []
        self.obs.append(ex_data[10])
        self.totaldata = []
        self.rpd = []
        self.dap = []
        self.drp = []
        self.ppa = []
        self.psa = []
        self.alm =[]
        self.alt = []
        self.meusal =[]
        self.meuscop =[]
        self.xfm = []
        self.xft = []
        self.cfa = []
        self.xrftm = []
        self.material = []
        self.thick = []
        n = 0
        self.ird = []
        self.dstd = []
        print (ex_data)
        for i in range(0, len(ex_data)):
            if i<10:
                if re.match(r'\*\d|\*+\d', ex_data[i]):
                    self.birth = ex_data[i]
                elif re.findall(r'#',ex_data[i]) == ['#']:
                    self.id.append(ex_data[i])

            if re.match(r'Manufacturer', ex_data[i]) and re.match(r'Observer', ex_data[i-1]) :
                self.manuf.append(ex_data[i+2])
            elif re.match(r'Model', ex_data[i]) and re.match(r'Observer', ex_data[i-1]):
                self.manuf[0] += " " + ex_data[i+3] +" "+ex_data[i+4]
            if ex_data[i] == 'Party' and ex_data[i+1] =='responsible'and ex_data[i+2] == 'for'and ex_data[i+3] == 'servicing'and ex_data[i+4] == 'the'and ex_data[i+5] == 'device':
                self.totaldata.append(ex_data[i + 6])# (1) Dose Area Product Total
                self.totaldata.append(ex_data[i + 7])# (2) Dose RP Total
                self.totaldata.append(ex_data[i + 9])# (3) Fluoro Dose Area Product Total
                self.totaldata.append(ex_data[i + 11])# (4) Fluoro Dose (RP) Total
                self.totaldata.append(ex_data[i + 13])# (5) Total Fluoro Time
                self.totaldata.append(ex_data[i + 15])# (6) Acquisition Dose Area Product Total
                self.totaldata.append(ex_data[i + 16])# (7) Acquisition Dose (RP) Total
                self.totaldata.append(ex_data[i + 18])# (8) Total Acuisition time
                self.totaldata.append(ex_data[i + 22])# Distance
                print(self.totaldata)
            if ex_data [i] == 'Flag:UNVERIFIEDContent':
                self.cont.extend([ex_data[i+1],ex_data[i+2]])
            elif 'Observer' == ex_data[i] and ex_data[i - 1] == 'Device' and ex_data[i + 1] == 'Name':
                self.obs.append(ex_data[i+3])

            if ex_data[i] == "Orthopaedics":
                n +=1
                self.xrftm.append(None)
                self.dap.append(ex_data[i + 8])# Dose Area Product
                self.drp.append(ex_data[i + 10])# Dose RP
                self.cfa.append(ex_data[i + 11])# Collimated Field Area
            elif ex_data[i] == "ﬁlter" and ex_data[i+1]=="Copper":
                self.alm.append(ex_data[i-5]+" "+ex_data[i-4]) #aluminum compound
                self.alt.append(ex_data[i-3]) #aluminum thickness
                self.meusal.append(ex_data[i-2]) # aluminum measurment unit
                self.xfm.append(ex_data[i+3]+" "+ex_data[i+4]) #copper compound
                self.xft.append(ex_data[i + 5])# copper thickness
                self.meuscop.append(ex_data[i+6]) # copper measurment unit
                self.ird.append(ex_data[i + 7])
                self.material.append(self.alm[n-1] + " & " + self.xfm[n-1])
                self.meusal[n-1] = self.meusal[n-1].replace('X-Ray', '')
                self.thick.append(self.alt[n-1]+self.meusal[n-1] +"Al "+self.xft[n-1]+self.meuscop[n-1]+"Cu")
        print(self.manuf)
        return [self.name,self.id, self.manuf, self.cont, self.obs, self.totaldata, self.rpd, self.dap, self.drp, self.ppa, self.psa, self.cfa, self.xrftm, self.ird, self.dstd, n]



    def startpro(self, pdf_file, index):

        # Specify the path to your Word file
        pdf_file_path = pdf_file
        text_data = self.extract_text_from_pdf(pdf_file_path)

        self.ndata = self.newdata(text_data)

        self.name,self.id, self.manuf, self.cont, self.obs, self.totaldata,\
            self.rpd, self.dap, self.drp, self.ppa, self.psa,\
             self.cfa, self.xrftm, self.ird, self.dstd, self.events = self.ndata

     ###### make all list without words######
        self.process_data()
    def process_data(self):
        self.process_name_and_study()
        self.clean_and_split(self.id, 0)
        self.clean_and_split(self.manuf, 0, exclude_quotes=True)
        self.clean_and_split(self.cont, 0)
        self.contime = self.cont[1].replace('X-Ray', '')
        self.clean_and_split(self.obs, 1)
        self.calculate_age()
        self.total = self.clean_numeric_list(self.totaldata, remove_units=['Gy.m2', 'Gy.m', 'cm', 's', 'Gy'])
        self.dap1 = self.convert_list_to_numbers([item.split(' ')[0] for item in self.dap])
        self.drp1 = self.extract_and_convert(self.drp, r'\d+\.\d+e[+-]\d+')
        self.psa1 = self.convert_list_to_numbers([item.split(' ')[0] for item in self.psa])
        self.cfa1 = self.extract_and_convert(self.cfa, r'\d+\.\d+e[+-]\d+')
        self.xrftm1 = self.convert_list_to_numbers(
            [item.split(' ')[0] if isinstance(item, str) else item for item in self.xrftm])
        self.ird1 = self.extract_numbers_from_list(self.ird)
        self.dstd1 = self.extract_numbers_from_list(self.dstd)

    def process_name_and_study(self):
        i = 0
        while not re.search(r'Study', str(self.name[i])):
            if re.match(r"\(+\w", self.name[i]):
                self.sex = self.name[i].replace('(', '').replace(',', '')
            self.start_index = self.name[i]
            i += 1
        self.study = " ".join(self.name[i + 1:]).replace("^", " ").replace("/", " ")
        self.study = re.sub(r'^.*Study:|Series:+\w+', '', self.study)
        self.name[0] = self.clean_text(self.name[0])
        self.capname = self.name[0].split()

    def clean_and_split(self, data_list, index, exclude_quotes=False):
        data_list[index] = self.clean_text(data_list[index], exclude_quotes)
        setattr(self, data_list[0] + '_split', data_list[index].split())

    def calculate_age(self):
        if self.birth:
            year, month, day = map(int, self.birth.split(",")[0][1:].split("-"))
            today = datetime.date.today()
            self.age = today.year - year - ((today.month, today.day) < (month, day))
        else:
            self.age = "N/A"

    @staticmethod
    def clean_text(text, exclude_quotes=False):
        replacements = [('"', ''), ('(', ''), (')', ''), ('*', ''), (',', ''), (':', ' ')]
        for old, new in replacements:
            if not exclude_quotes or old != '"':
                text = text.replace(old, new)
        return text

    @staticmethod
    def clean_numeric_list(data_list, remove_units):
        cleaned_list = []
        for item in data_list:
            for unit in remove_units:
                item = item.replace(unit, '')
            cleaned_list.append(float(item) if item else None)
        return cleaned_list

    @staticmethod
    def convert_list_to_numbers(data_list):
        converted_list = []
        for item in data_list:
            try:
                converted_list.append(int(item))
            except ValueError:
                try:
                    converted_list.append(float(item))
                except ValueError:
                    converted_list.append(None)
        return converted_list

    @staticmethod
    def extract_and_convert(data_list, pattern):
        return [float(re.search(pattern, value).group()) for value in data_list]

    @staticmethod
    def extract_numbers_from_list(data_list):
        numbers = []
        for item in data_list:
            numbers.extend([float(seq) if '.' in seq else int(seq) for seq in re.findall(r'\b\d+(?:\.\d+)?\b', item)])
        return numbers

        # Create a pandas DataFrame from the extrac+ted data
        self.all_data = { "Dose Area Product (Gym\u00b2)":self.dap1,"Dose (RP) (Gy)":self.drp1,'Collimated Field Area (m\u00b2)':self.cfa1,
                          "X-Ray Filter Material":self.material,'X-Ray Filter Thickness':self.thick,
                          'Irradiation Duration (s)':self.ird1}


        # Use regular expressions to extract numbers from mixed elements
        max_length = max(len(self.all_data[col]) for col in self.all_data)
        for col in self.all_data:
            self.all_data[col] += [np.nan] * (max_length - len(self.all_data[col]))

        self.df = pd.DataFrame(self.all_data)
        for i in range(0,self.events):
            self.df = self.df.rename(index={i: "Event {0}".format(i+1)})
        self.df = self.df.rename_axis('Irradiation Event X-Ray Data of '+ f"Patient {index}") #self.capname[1]+" "+self.name[1])

        self.data_total = {"Patient ID": f"ID {index}", "Dose Area Product Total (Gym\u00b2)": self.total[0],
                           "Dose (RP) Total (Gy)":self.total[1],"Fluoro Dose Area Product Total (Gym\u00b2)":self.total[2],"Fluoro Dose (RP) Total (Gy)":self.total[3],
                           "Total Fluoro Time (s)":self.total[4],"Acquisition Dose Area Product Total (Gym\u00b2)":self.total[5],
                            "Acquisition Dose (RP) Total (Gy)":self.total[6],	"Total Acquisition Time (s)":self.total[7],"Reference Point Definition (mm)":self.total[8]}
        self.dft = pd.DataFrame(self.data_total, index=[f"Patient {index}"])
        self.dft = self.dft.rename_axis('Patient Name')

        self.person_data = [self.capname[1]+" "+self.name[1],f"ID {index}",self.sex, self.age,
                            self.study, self.manufacturer[0],self.content[1],self.contime,self.observer[0],self.events]

        self.dfper = pd.DataFrame(self.person_data, index=['Patient Name', 'Patient ID','Gender','Age (years)', 'Study Type', 'Manufacturer', 'Content Time', 'Content Date', 'Person Observer Name',
                                         'Number of irradiation events'], columns=[""])
        self.individual = {" Patient ID": self.name_id[0], "Gender": self.sex, "Age (years)": self.age, "Study Type": self.study,
                           "Manufacturer": self.manufacturer[0], "Content Date": self.content[1],
                           "Content Time": self.contime, "Person Observer Name": "Ν/Α"}
        self.dfin = pd.DataFrame(self.individual, index=[f"Patient {index}"])
        self.dfin = self.dfin.rename_axis('Patient Name')

        return self.df, self.dft,self.individual,self.dfin, self.person_data, self.name_id[0]

    def get_dataframes(self):
        return self.df, self.dft, self.dfper, self.name_id[0]

