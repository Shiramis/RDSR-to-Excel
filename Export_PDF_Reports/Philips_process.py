import pandas as pd
import re
import numpy as np
import PyPDF2
import datetime

class make_excelp ():

    def extract_text_from_pdf(self,pdf_file_path):
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
        self.birth = 'N/A'

        self.id = []
        self.manuf = []
        self.cont = []
        self.obs = []
        self.obs.append(ex_data[10])
        self.totaldata = []
        self.rpd = []
        self.dap = []
        self.drp = []
        self.alm =[]
        self.alt = []
        self.meusal =[]
        self.meuscop =[]
        self.xfm = []
        self.xft = []
        self.cfa = []
        self.material = []
        self.thick = []
        n = 0
        self.ird = []
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
        return [self.name,self.id, self.manuf, self.cont, self.obs, self.totaldata, self.rpd, self.dap, self.drp,
                 self.cfa, self.ird, n]



    def startpro(self, pdf_file, index):
        pdf_file_path = pdf_file
        text_data = self.extract_text_from_pdf(pdf_file_path)

        self.ndata = self.newdata(text_data)

        self.name,self.id, self.manuf, self.cont, self.obs, self.totaldata,\
            self.rpd, self.dap, self.drp,\
             self.cfa, self.ird, self.events = self.ndata

     ###### make all list without words######
        i = 0
        while not re.search(r'Study', str(self.name[i])):
            if re.match(r"\(+\w", self.name[i]):
                self.sex = self.name[i]
            self.start_index = self.name[i]
            indexs = i
            i += 1
        self.study = ""
        for i in range(indexs + 1, len(self.name)):
            self.study = self.study + " " + self.name[i]
        self.study = re.sub(r'^.*Study:', '', self.study)
        self.study = re.sub(r'Series:+\w+', '', self.study)
        self.study = self.study.replace("^", " ").replace("/", " ")
        self.sex = self.sex.replace('(', '').replace(',', '')

        self.name[0] = self.name[0].replace(':',' ').replace('(', '').replace(')', '').replace('*', '').replace(',', '')
        self.capname = self.name[0].split()

        self.id[0]= self.id[0].replace(':',' ').replace('(', '').replace(')', ' ').replace('*', '').replace(',', '')
        self.name_id = self.id[0].split()

        self.manuf[0] = self.manuf[0].replace('"','').replace('(', '').replace(')', '').replace('*', '').replace(',', '')
        self.manufacturer = self.manuf

        self.cont[0] = self.cont[0].replace(':',' ').replace('(', '').replace(')', '').replace('*', '').replace(',', '')
        self.content = self.cont[0].split()
        self.contime = self.cont[1].replace('X-Ray','')

        self.obs[1] = self.obs[1].replace('(', '').replace(')', '').replace('*', '').replace(',', '')
        self.observer = self.obs[1].split()
        if self.birth != 'N/A':
            self.birth = self.birth.split(',')[0][1:]
            year, month, day = map(int, self.birth.split("-"))
            today = datetime.date.today()
            self.age = today.year - year - ((today.month, today.day) < (month, day))
        else:
            self.age = "N/A"
        self.total = []
        for item in self.totaldata:
            value = item.replace('Gy.m2', '')
            value = value.replace('Gy.m', '').replace('cm', '').replace('s', '')
            if 'Gy' in value:
                value = value.replace('Gy', '')
            if value:
                self.total.append(float(value))
            else:
                self.total.append('N/A')
        self.moddap = [item.split(' ')[0] for item in self.dap]
        self.dap1 = []
        for item in self.moddap:
            try:
                num = int(item)
            except ValueError:
                try:
                    num = float(item)
                except ValueError:
                    num = 'N/A'
            self.dap1.append(num)
        self.drp1 = [float(re.search(r'\d+\.\d+e[+-]\d+', value).group()) for value in self.drp]


        self.cfa1 = [float(re.search(r'\d+\.\d+e[+-]\d+', value).group()) for value in self.cfa]
        self.ird1 = []
        for item in self.ird:
            numeric_sequences = re.findall(r'\b\d+(?:\.\d+)?\b', item)

            for seq in numeric_sequences:
                self.ird1.append(seq)
        self.ird1 = [float(num) if '.' in num else int(num) for num in self.ird1]

        self.all_data = { "Dose Area Product (Gym\u00b2)":self.dap1,"Dose (RP) (Gy)":self.drp1,'Collimated Field Area (m\u00b2)':self.cfa1,
                          "X-Ray Filter Material":self.material,'X-Ray Filter Thickness':self.thick,
                          'Irradiation Duration (s)':self.ird1}

        max_length = max(len(self.all_data[col]) for col in self.all_data)
        for col in self.all_data:
            self.all_data[col] += [np.nan] * (max_length - len(self.all_data[col]))

        self.df = pd.DataFrame(self.all_data)
        for i in range(0,self.events):
            self.df = self.df.rename(index={i: "Event {0}".format(i+1)})
        self.df = self.df.rename_axis('Irradiation Event X-Ray Data of '+ f"Patient {index}") #self.capname[1]+" "+self.name[1])

        self.data_total = {"Patient ID": f"Patient ID {index}", "Dose Area Product Total (Gym\u00b2)": self.total[0],
                           "Dose (RP) Total (Gy)":self.total[1],"Fluoro Dose Area Product Total (Gym\u00b2)":self.total[2],"Fluoro Dose (RP) Total (Gy)":self.total[3],
                           "Total Fluoro Time (s)":self.total[4],"Acquisition Dose Area Product Total (Gym\u00b2)":self.total[5],
                            "Acquisition Dose (RP) Total (Gy)":self.total[6],	"Total Acquisition Time (s)":self.total[7],"Reference Point Definition (mm)":self.total[8]}
        self.dft = pd.DataFrame(self.data_total, index=[f"Patient {index}"])
        self.dft = self.dft.rename_axis('Patient Name')

        self.person_data = [f"Patient {index}",f"Patient ID {index}",self.sex, self.age,
                            self.study, self.manufacturer[0],self.content[1],self.contime,self.observer[0],self.events]

        self.dfper = pd.DataFrame(self.person_data, index=['Patient Name', 'Patient ID','Gender','Age (years)', 'Study Type', 'Manufacturer', 'Content Time', 'Content Date', 'Person Observer Name',
                                         'Number of irradiation events'], columns=[""])
        self.individual = {" Patient ID": f"Patient ID {index}", "Gender": self.sex, "Age (years)": self.age, "Study Type": self.study,
                           "Manufacturer": self.manufacturer[0], "Content Date": self.content[1],
                           "Content Time": self.contime, "Person Observer Name": "Ν/Α"}
        self.dfin = pd.DataFrame(self.individual, index=[f"Patient {index}"])
        self.dfin = self.dfin.rename_axis('Patient Name')

        return self.df, self.dft,self.individual,self.dfin, self.person_data, self.name_id[0]
    def get_dataframes(self):
        return self.df, self.dft, self.dfper, self.name_id[0]

